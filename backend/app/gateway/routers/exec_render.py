"""Non-agentic report rendering endpoint.

VEPIP's document generation does not need an agent — the build scripts in
``deer-flow/skills/custom/vepip-reports/scripts/`` are fully deterministic.
This router takes the data + format, optionally generates narrative blocks
via a direct Gemini call (no LangGraph, no middleware chain), spawns the
build script as a subprocess, and streams progress back as SSE.

Auth: ``Authorization: Bearer <VEPIP_INTERNAL_SECRET>`` only. Path is
exempted from both AuthMiddleware (JWT) and CSRFMiddleware in those
modules' public-path lists.
"""
from __future__ import annotations

import asyncio
import json
import os
import re
import secrets
import subprocess
import sys
import threading
import uuid
from pathlib import Path
from typing import Any, AsyncIterator, Literal

from fastapi import APIRouter, Header, HTTPException, Request
from fastapi.responses import FileResponse, StreamingResponse
from pydantic import BaseModel, Field

router = APIRouter(prefix="/api/exec", tags=["exec"])

# Repo root: backend/app/gateway/routers/exec_render.py -> repo root is 4 parents up
_REPO_ROOT = Path(__file__).resolve().parents[4]
_BACKEND_ROOT = _REPO_ROOT / "backend"
_REPORTS_ROOT = _BACKEND_ROOT / ".deer-flow" / "reports"
_SCRIPTS_DIR = _REPO_ROOT / "skills" / "custom" / "vepip-reports" / "scripts"

# Best-effort load of deer-flow/.env so GOOGLE_AI_API_KEY / VEPIP_INTERNAL_SECRET
# are present even if the process was launched from a cwd where python-dotenv's
# find_dotenv() upward search misses it.
try:
    from dotenv import load_dotenv as _load_dotenv

    _DEERFLOW_ENV = Path(__file__).resolve().parents[4] / ".env"
    if _DEERFLOW_ENV.is_file():
        _load_dotenv(_DEERFLOW_ENV, override=False)
except ImportError:  # pragma: no cover — dotenv is a hard DeerFlow dep
    pass

_FILENAME_RE = re.compile(r"^[A-Za-z0-9_.-]+\.(pptx|docx|pdf)$")
_VIBES = {"editorial-serif", "dark-premium", "magazine-bold", "ocean-corporate"}
_SUBPROCESS_TIMEOUT_SECONDS = 300.0
_NARRATIVE_TIMEOUT_SECONDS = 60.0


def _require_secret(authorization: str | None) -> None:
    expected = os.environ.get("VEPIP_INTERNAL_SECRET", "")
    if not expected:
        raise HTTPException(status_code=503, detail="VEPIP_INTERNAL_SECRET not configured")
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(status_code=401, detail="missing bearer token")
    token = authorization[len("Bearer ") :]
    if not secrets.compare_digest(token, expected):
        raise HTTPException(status_code=401, detail="invalid token")


class RenderRequest(BaseModel):
    format: Literal["pptx", "docx", "pdf"]
    report_type: Literal["quarterly", "full"]
    project_id: str
    project_name: str
    filename: str
    period_start: str
    period_end: str
    vibe: str | None = None
    data: dict[str, Any]
    generate_narrative: bool = True
    # Default to moonshot-v1-auto: non-thinking, OpenAI-compatible `content`
    # output, auto-scales 8k/32k/128k context window based on prompt size.
    # The k2.6 / k2-thinking models return their answer in `reasoning_content`
    # (not `content`) and burn the output budget on hidden reasoning, which
    # breaks the JSON-extraction parser below — same gotcha that bit the old
    # Gemini code with `gemini-2.5-flash`.
    narrative_model: str = "moonshot-v1-auto"

    def validated_filename(self) -> str:
        if not _FILENAME_RE.match(self.filename):
            raise HTTPException(status_code=400, detail=f"invalid filename: {self.filename}")
        ext = self.filename.rsplit(".", 1)[-1]
        if ext != self.format:
            raise HTTPException(status_code=400, detail=f"filename extension {ext} != format {self.format}")
        return self.filename

    def validated_vibe(self) -> str | None:
        if self.vibe is None:
            return None
        if self.vibe not in _VIBES:
            raise HTTPException(status_code=400, detail=f"invalid vibe: {self.vibe}")
        return self.vibe


def _sse(event: str, data: dict[str, Any]) -> bytes:
    return f"event: {event}\ndata: {json.dumps(data, ensure_ascii=False)}\n\n".encode("utf-8")


def _narrative_prompt(req: RenderRequest) -> str:
    return (
        f"You are drafting funder-report narrative for Vision Empower Trust's project "
        f"'{req.project_name}' — a {req.report_type} report covering "
        f"{req.period_start} to {req.period_end}.\n\n"
        f"Project data (JSON, may be truncated):\n"
        f"{json.dumps(req.data, ensure_ascii=False)[:8000]}\n\n"
        "Return a JSON object with EXACTLY these four fields (no other fields, no markdown "
        "fences, no preamble — just the JSON):\n"
        '  "overview"      — 120-180 words. Project mission, scope, why it matters this period.\n'
        '  "achievements"  — 150-220 words. Specific quantified wins. Cite numbers from data.\n'
        '  "challenges"    — 80-140 words. Honest constraints, blockers, or risks observed.\n'
        '  "way_forward"   — 80-140 words. Concrete next steps tied to remaining deliverables.\n\n'
        "Rules: Reference concrete numbers from the data. Do not invent facts. Do not include "
        "section headings inside the field values. Professional but warm tone suitable for "
        "institutional funders (Wipro Foundation, CSR funds, etc.). Output JSON only."
    )


_NARRATIVE_KEYS = ("overview", "achievements", "challenges", "way_forward")


def _parse_narrative_json(raw: str) -> dict[str, str]:
    """Extract a 4-key narrative dict from LLM output. Lenient: tolerates fences."""
    text = raw.strip()
    # Strip ```json … ``` fences if present
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```\s*$", "", text)
    try:
        obj = json.loads(text)
    except json.JSONDecodeError:
        # Try to find a JSON object inside the text
        match = re.search(r"\{[\s\S]*\}", text)
        if not match:
            return {}
        try:
            obj = json.loads(match.group(0))
        except json.JSONDecodeError:
            return {}
    if not isinstance(obj, dict):
        return {}

    def _coerce(value: Any) -> str:
        if value is None:
            return ""
        if isinstance(value, str):
            return value.strip()
        if isinstance(value, list):
            # Gemini often returns bullet arrays for achievements / challenges /
            # way_forward. Render them as prose-friendly bullet lines.
            return "\n".join(f"• {str(item).strip()}" for item in value if str(item).strip())
        return str(value).strip()

    return {k: _coerce(obj.get(k)) for k in _NARRATIVE_KEYS}


_MOONSHOT_BASE_URL_DEFAULT = "https://api.moonshot.ai/v1"


async def _generate_narrative(req: RenderRequest) -> tuple[dict[str, str], str | None]:
    """Direct Kimi (Moonshot) call via the OpenAI-compatible endpoint —
    no LangGraph, no middleware chain.

    Routes through ChatOpenAI with `base_url` pointed at Moonshot. Honours
    `req.narrative_model` so callers can override to `moonshot-v1-32k`,
    `kimi-k2-thinking`, etc. Falls back to a clean error string when the
    SDK isn't installed or the key isn't set.

    Returns (blocks, error_message). On success error_message is None;
    on failure blocks is {} and error_message describes what went wrong.
    """
    try:
        from langchain_openai import ChatOpenAI
    except ImportError as exc:
        return {}, f"langchain_openai not installed: {exc}"

    api_key = os.environ.get("MOONSHOT_API_KEY")
    if not api_key:
        return {}, "MOONSHOT_API_KEY not set in process env"

    base_url = os.environ.get("MOONSHOT_BASE_URL") or _MOONSHOT_BASE_URL_DEFAULT

    try:
        llm = ChatOpenAI(
            model=req.narrative_model,
            api_key=api_key,
            base_url=base_url,
            temperature=0.35,
            max_tokens=4096,
            timeout=_NARRATIVE_TIMEOUT_SECONDS,
        )
        result = await asyncio.wait_for(
            llm.ainvoke(_narrative_prompt(req)),
            timeout=_NARRATIVE_TIMEOUT_SECONDS,
        )
        text = result.content if hasattr(result, "content") else str(result)
        if isinstance(text, list):
            text = "\n".join(str(part) for part in text)
        blocks = _parse_narrative_json(str(text))
        if not any(blocks.values()):
            return {}, f"narrative parse produced no fields (raw len={len(str(text))})"
        return blocks, None
    except Exception as exc:  # noqa: BLE001 — narrative is best-effort
        return {}, f"{type(exc).__name__}: {exc}"


def _narrative_to_draft(blocks: dict[str, str]) -> str:
    """Flatten narrative blocks into the legacy ``draft`` string format so old
    build-script consumers that haven't been updated yet still work."""
    parts: list[str] = []
    if blocks.get("overview"):
        parts.append(f"Overview\n\n{blocks['overview']}")
    if blocks.get("achievements"):
        parts.append(f"Achievements\n\n{blocks['achievements']}")
    if blocks.get("challenges"):
        parts.append(f"Challenges\n\n{blocks['challenges']}")
    if blocks.get("way_forward"):
        parts.append(f"Way Forward\n\n{blocks['way_forward']}")
    return "\n\n".join(parts)


async def _stream_render(req: RenderRequest, report_id: str) -> AsyncIterator[bytes]:
    filename = req.validated_filename()
    vibe = req.validated_vibe()

    out_dir = _REPORTS_ROOT / report_id
    out_dir.mkdir(parents=True, exist_ok=True)
    data_path = out_dir / "report_data.json"
    output_path = out_dir / filename

    download_url = f"/api/exec/artifact/{report_id}/{filename}"
    yield _sse("init", {"report_id": report_id, "download_url": download_url, "filename": filename})

    # ── Narrative ─────────────────────────────────────────────────────────
    narrative_blocks: dict[str, str] = {}
    draft = ""
    if req.generate_narrative:
        yield _sse("narrative-start", {"model": req.narrative_model})
        narrative_blocks, narr_err = await _generate_narrative(req)
        draft = _narrative_to_draft(narrative_blocks)
        yield _sse(
            "narrative",
            {
                "blocks": narrative_blocks,
                "text": draft,
                "length": len(draft),
                "structured": bool(narrative_blocks),
                "error": narr_err,
            },
        )

    # ── Write merged data JSON ────────────────────────────────────────────
    merged = {
        "project": {**req.data, "narrative": narrative_blocks},
        "periodStart": req.period_start,
        "periodEnd": req.period_end,
        "draft": draft,
        "narrative": narrative_blocks,
    }
    try:
        data_path.write_text(json.dumps(merged, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError as exc:
        yield _sse("error", {"message": f"failed to write data json: {exc}", "code": "write_data"})
        return

    # ── Spawn build script ────────────────────────────────────────────────
    script = _SCRIPTS_DIR / f"build_{req.format}.py"
    if not script.is_file():
        yield _sse("error", {"message": f"build script not found: {script}", "code": "script_missing"})
        return

    cmd: list[str] = [
        sys.executable,
        str(script),
        "--data",
        str(data_path),
        "--output",
        str(output_path),
        "--report-type",
        req.report_type,
        "--period-start",
        req.period_start,
        "--period-end",
        req.period_end,
    ]
    if vibe:
        cmd.extend(["--vibe", vibe])

    yield _sse("render-start", {"script": script.name, "cmd": cmd})

    # Use a thread + sync subprocess. This avoids Windows event-loop pitfalls
    # (asyncio.create_subprocess_exec raises NotImplementedError on
    # SelectorEventLoop, which silently kills the SSE stream). The thread
    # pipes stdout lines into an asyncio.Queue that the SSE generator drains.
    loop = asyncio.get_running_loop()
    line_queue: asyncio.Queue[tuple[str, str | int | None]] = asyncio.Queue()
    tail: list[str] = []

    def _put(kind: str, payload: str | int | None) -> None:
        loop.call_soon_threadsafe(line_queue.put_nowait, (kind, payload))

    def _runner() -> None:
        try:
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                cwd=str(_SCRIPTS_DIR),
                bufsize=1,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
        except FileNotFoundError as exc:
            _put("spawn_error", f"python not found: {exc}")
            _put("end", None)
            return
        except Exception as exc:  # noqa: BLE001
            _put("spawn_error", f"failed to spawn subprocess: {exc}")
            _put("end", None)
            return

        assert proc.stdout is not None
        try:
            for raw in proc.stdout:
                line = raw.rstrip("\r\n")
                _put("line", line)
            rc = proc.wait(timeout=_SUBPROCESS_TIMEOUT_SECONDS)
            _put("rc", rc)
        except subprocess.TimeoutExpired:
            proc.kill()
            try:
                proc.wait(timeout=5)
            except Exception:  # noqa: BLE001
                pass
            _put("timeout", None)
        except Exception as exc:  # noqa: BLE001
            _put("spawn_error", f"subprocess i/o error: {exc}")
        finally:
            _put("end", None)

    thread = threading.Thread(target=_runner, daemon=True)
    thread.start()

    spawn_error: str | None = None
    rc: int | None = None
    timed_out = False
    try:
        while True:
            try:
                kind, payload = await asyncio.wait_for(line_queue.get(), timeout=_SUBPROCESS_TIMEOUT_SECONDS + 30)
            except asyncio.TimeoutError:
                timed_out = True
                break
            if kind == "line":
                assert isinstance(payload, str)
                tail.append(payload)
                if len(tail) > 200:
                    tail.pop(0)
                yield _sse("render-log", {"line": payload})
            elif kind == "rc":
                assert isinstance(payload, int)
                rc = payload
            elif kind == "timeout":
                timed_out = True
            elif kind == "spawn_error":
                spawn_error = str(payload)
            elif kind == "end":
                break

        if spawn_error:
            yield _sse("error", {"message": spawn_error, "code": "spawn"})
            return
        if timed_out:
            yield _sse("error", {"message": "build script timed out", "code": "timeout", "tail": tail[-20:]})
            return
        if rc is None:
            yield _sse("error", {"message": "build script ended without exit code", "code": "no_rc", "tail": tail[-20:]})
            return
        if rc != 0:
            yield _sse("error", {"message": f"build script exited {rc}", "code": "nonzero_exit", "tail": tail[-30:]})
            return
        if not output_path.is_file():
            yield _sse("error", {"message": "build script succeeded but output missing", "code": "no_output", "tail": tail[-20:]})
            return
        size = output_path.stat().st_size
        if size == 0:
            yield _sse("error", {"message": "build produced empty file", "code": "empty_output"})
            return

        yield _sse("render-complete", {"bytes": size, "output_path": str(output_path)})
        yield _sse("done", {"report_id": report_id, "download_url": download_url, "bytes": size})
    finally:
        # Best-effort join (thread is daemon; will die with process anyway).
        if thread.is_alive():
            thread.join(timeout=2)


@router.post("/render-report")
async def render_report(
    request: Request,
    authorization: str | None = Header(None),
) -> StreamingResponse:
    _require_secret(authorization)
    body = await request.json()
    try:
        req = RenderRequest.model_validate(body)
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=400, detail=f"invalid body: {exc}") from exc

    report_id = uuid.uuid4().hex

    return StreamingResponse(
        _stream_render(req, report_id),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache, no-transform",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",
        },
    )


_MIME_TYPES = {
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "pdf": "application/pdf",
}

_REPORT_ID_RE = re.compile(r"^[a-f0-9]{32}$")


@router.get("/artifact/{report_id}/{filename}")
async def fetch_artifact(
    report_id: str,
    filename: str,
    authorization: str | None = Header(None),
) -> FileResponse:
    _require_secret(authorization)
    if not _REPORT_ID_RE.match(report_id):
        raise HTTPException(status_code=400, detail="invalid report_id")
    if not _FILENAME_RE.match(filename):
        raise HTTPException(status_code=400, detail="invalid filename")

    path = _REPORTS_ROOT / report_id / filename
    if not path.is_file():
        raise HTTPException(status_code=404, detail="artifact not found")

    ext = filename.rsplit(".", 1)[-1]
    return FileResponse(
        path=str(path),
        media_type=_MIME_TYPES.get(ext, "application/octet-stream"),
        filename=filename,
        headers={"Cache-Control": "private, no-store"},
    )
