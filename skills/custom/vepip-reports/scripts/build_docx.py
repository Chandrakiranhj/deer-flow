#!/usr/bin/env python3
"""Vision Empower premium funder report — branded DOCX builder.

Usage:
    python build_docx.py --data /path/to/data.json --output /path/to/out.docx
"""
from __future__ import annotations

import argparse
import json
import sys
import traceback
from datetime import datetime
from io import BytesIO
from pathlib import Path
from urllib.request import Request, urlopen

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _bootstrap import ensure  # type: ignore

ensure({"docx": "python-docx"})

from docx import Document  # noqa: E402
from docx.enum.table import WD_ALIGN_VERTICAL  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import Cm, Pt, RGBColor  # noqa: E402

DBROWN = RGBColor(0x2A, 0x15, 0x08)
GOLD   = RGBColor(0xC4, 0x9A, 0x32)
BODY   = RGBColor(0x3D, 0x20, 0x10)
MUTED  = RGBColor(0x9B, 0x7B, 0x5A)


def fmt_money(n) -> str:
    if n is None: return "—"
    n = float(n)
    if n >= 1e7:  return f"₹{n / 1e7:.1f}Cr"
    if n >= 1e5:  return f"₹{n / 1e5:.1f}L"
    if n >= 1e3:  return f"₹{n / 1e3:.0f}K"
    return f"₹{int(n):,}"


def fmt_date(d) -> str:
    if not d: return "—"
    try:
        dt = datetime.fromisoformat(d.replace("Z", "+00:00"))
        return dt.strftime("%-d %b %Y") if sys.platform != "win32" else dt.strftime("%#d %b %Y")
    except Exception:
        return str(d)


def fetch_image_bytes(url: str) -> bytes | None:
    if not url:
        return None
    try:
        req = Request(url, headers={"User-Agent": "vepip-report-bot/1.0"})
        with urlopen(req, timeout=8) as r:
            return r.read()
    except Exception:
        return None


def shade_cell(cell, hex_fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tc_pr.append(shd)


def styled_run(p, text: str, *, size=11, bold=False, italic=False, color: RGBColor = BODY, font="Calibri"):
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    r.font.name = font
    return r


def heading(doc, text: str, *, size=18, color: RGBColor = DBROWN):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    styled_run(p, text, size=size, bold=True, color=color, font="Georgia")
    # Thin gold rule below
    rule = doc.add_paragraph()
    rule.paragraph_format.space_after = Pt(8)
    pPr = rule._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bot = OxmlElement("w:bottom")
    bot.set(qn("w:val"), "single")
    bot.set(qn("w:sz"), "8")
    bot.set(qn("w:color"), "C49A32")
    pBdr.append(bot)
    pPr.append(pBdr)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--data", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--report-type", default="quarterly")
    parser.add_argument("--period-start", default="")
    parser.add_argument("--period-end", default="")
    parser.add_argument("--draft", default="")
    args = parser.parse_args(argv)

    data = json.loads(Path(args.data).read_text(encoding="utf-8"))
    project = data.get("project") or data
    if "budgets" not in project and project.get("budgetCategories"):
        project["budgets"] = project["budgetCategories"]
    if "activities" not in project and project.get("recentActivities"):
        project["activities"] = project["recentActivities"]
    period_start = args.period_start or data.get("periodStart") or project.get("startDate") or ""
    period_end   = args.period_end   or data.get("periodEnd")   or project.get("endDate")   or ""
    draft        = args.draft or data.get("draft") or ""

    doc = Document()
    # Page margins
    for section in doc.sections:
        section.top_margin    = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin   = Cm(2.2)
        section.right_margin  = Cm(2.2)

    # Cover block
    cover = doc.add_paragraph()
    cover.alignment = WD_ALIGN_PARAGRAPH.LEFT
    styled_run(cover, "VISION EMPOWER", size=10, bold=True, color=GOLD, font="Calibri")
    cover.add_run("\n")
    styled_run(cover, "QUARTERLY FUNDER REPORT" if args.report_type == "quarterly" else "FULL PROJECT REPORT",
               size=9, color=MUTED, font="Calibri")

    title = doc.add_paragraph()
    title.paragraph_format.space_before = Pt(8)
    styled_run(title, project.get("name", ""), size=28, bold=True, color=DBROWN, font="Georgia")

    sub = doc.add_paragraph()
    is_quarterly = args.report_type == "quarterly"
    period = (
        f"{fmt_date(period_start)} – {fmt_date(period_end)}" if is_quarterly
        else f"{fmt_date(project.get('startDate'))} – {fmt_date(project.get('endDate'))}"
    )
    styled_run(sub, f"Funder: {project.get('funderName', '')}    |    Period: {period}    |    Grant: {fmt_money(project.get('grantAmount'))}",
               size=10, italic=True, color=MUTED)

    if project.get("states"):
        st = doc.add_paragraph()
        styled_run(st, f"States: {', '.join(project['states'])}", size=10, color=MUTED)

    # ── Executive summary ──────────────────────────────────────────────────
    heading(doc, "Executive Summary")
    summary = project.get("summary") or ""
    if not summary and draft:
        lines = [l.strip() for l in draft.split("\n") if len(l.strip()) > 60]
        summary = " ".join(lines[:4])[:600]
    if not summary:
        summary = (
            f"This report summarises progress made under the {project.get('name', '')} "
            f"project funded by {project.get('funderName', '')}."
        )
    p = doc.add_paragraph()
    styled_run(p, summary, size=11, color=BODY)

    # ── Deliverables table ─────────────────────────────────────────────────
    heading(doc, "Deliverable Progress")
    deliverables = project.get("deliverables") or []
    if deliverables:
        table = doc.add_table(rows=1, cols=4)
        table.style = "Light Grid Accent 1"
        hdr = table.rows[0].cells
        for i, label in enumerate(["Deliverable", "Target", "Achieved", "% Complete"]):
            hdr[i].text = ""
            p_h = hdr[i].paragraphs[0]
            styled_run(p_h, label, size=10, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
            shade_cell(hdr[i], "2A1508")
        for d in deliverables:
            row = table.add_row().cells
            unit = d.get("unit") or ""
            row[0].text = d.get("title", "")
            row[1].text = f"{d.get('target') or '—'}{(' ' + unit) if (d.get('target') and unit) else ''}"
            row[2].text = f"{d.get('achieved') or '—'}{(' ' + unit) if (d.get('achieved') and unit) else ''}"
            target = d.get("target") or 0
            achieved = d.get("achieved") or 0
            pct = round(achieved / target * 100) if target else (100 if d.get("status") == "completed" else 0)
            row[3].text = f"{pct}%"
    else:
        styled_run(doc.add_paragraph(), "No deliverables recorded.", size=11, color=MUTED, italic=True)

    # ── Field activities ───────────────────────────────────────────────────
    heading(doc, "Field Activities")
    activities = project.get("activities") or []
    if activities:
        for act in activities[:20]:
            p = doc.add_paragraph(style="List Bullet")
            date_str = fmt_date(act.get("activityDate")) if act.get("activityDate") else ""
            head = f"{date_str} — {act.get('title', '')}" if date_str else act.get("title", "")
            styled_run(p, head + ". ", size=11, bold=True, color=DBROWN)
            reaches = []
            if act.get("teachersReached"): reaches.append(f"{act['teachersReached']} teachers")
            if act.get("studentsReached"): reaches.append(f"{act['studentsReached']} students")
            if act.get("schoolsReached"):  reaches.append(f"{act['schoolsReached']} schools")
            tail_bits = []
            if act.get("state"): tail_bits.append(act["state"])
            if reaches: tail_bits.append(", ".join(reaches))
            if act.get("notes"): tail_bits.append(act["notes"])
            if tail_bits:
                styled_run(p, " · ".join(tail_bits), size=11, color=BODY)
    else:
        styled_run(doc.add_paragraph(), "No activities recorded.", size=11, color=MUTED, italic=True)

    # ── Impact at a glance ─────────────────────────────────────────────────
    heading(doc, "Impact at a Glance")
    teachers = sum(a.get("teachersReached") or 0 for a in activities)
    students = sum(a.get("studentsReached") or 0 for a in activities)
    schools  = sum(a.get("schoolsReached") or 0 for a in activities)
    impact_table = doc.add_table(rows=1, cols=3)
    impact_table.style = "Light Grid Accent 1"
    impact_table.rows[0].cells[0].text = ""
    impact_table.rows[0].cells[1].text = ""
    impact_table.rows[0].cells[2].text = ""
    for i, (val, lbl) in enumerate([(f"{teachers:,}", "Teachers Reached"),
                                     (f"{students:,}", "Students Reached"),
                                     (f"{schools:,}",  "Schools Reached")]):
        cell = impact_table.rows[0].cells[i]
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        styled_run(cell.paragraphs[0], val, size=24, bold=True, color=DBROWN, font="Georgia")
        p2 = cell.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        styled_run(p2, lbl, size=9, color=MUTED, bold=True)
        shade_cell(cell, "F5EDD0")

    # ── Financial summary ──────────────────────────────────────────────────
    heading(doc, "Budget & Utilisation")
    approved = project.get("approvedBudget") or 0
    spent    = project.get("spentBudget") or 0
    util_pct = round(spent / approved * 100) if approved else 0
    fp = doc.add_paragraph()
    styled_run(fp, f"Approved: {fmt_money(approved)}    Spent: {fmt_money(spent)}    Utilisation: {util_pct}%",
               size=11, color=BODY)
    budgets = project.get("budgets") or []
    if budgets:
        bt = doc.add_table(rows=1, cols=4)
        bt.style = "Light Grid Accent 1"
        for i, h in enumerate(["Category", "Approved", "Spent", "Utilisation"]):
            cell = bt.rows[0].cells[i]
            cell.text = ""
            styled_run(cell.paragraphs[0], h, size=10, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
            shade_cell(cell, "2A1508")
        for b in budgets:
            row = bt.add_row().cells
            row[0].text = b.get("name", "")
            row[1].text = fmt_money(b.get("approvedAmount"))
            row[2].text = fmt_money(b.get("spentAmount"))
            ap = b.get("approvedAmount") or 0
            sp = b.get("spentAmount") or 0
            row[3].text = f"{round(sp / ap * 100) if ap else 0}%"

    # ── Stories ────────────────────────────────────────────────────────────
    testimonials = project.get("testimonials") or []
    if not testimonials:
        for a in activities:
            if a.get("testimonial"):
                testimonials.append({
                    "content": a["testimonial"],
                    "author": a.get("testimonialBy") or a.get("title", ""),
                    "role": a.get("state") or "",
                })
    if testimonials:
        heading(doc, "Stories from the Field")
        for t in testimonials[:3]:
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.5)
            styled_run(p, f'"{t.get("content", "")}"', size=12, italic=True, color=DBROWN, font="Georgia")
            p2 = doc.add_paragraph()
            p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            role = t.get("role") or ""
            styled_run(p2, f"— {t.get('author', '')}{', ' + role if role else ''}", size=10, color=MUTED)

    # ── Way forward ────────────────────────────────────────────────────────
    heading(doc, "Challenges & Next Steps")
    import re as _re

    def extract(text, start_re, stop_re):
        parts = _re.split(start_re, text, flags=_re.IGNORECASE, maxsplit=1)
        if len(parts) < 2: return ""
        content = parts[1]
        m = _re.search(stop_re, content, flags=_re.IGNORECASE)
        if m: content = content[: m.start()]
        return _re.sub(r"\s+", " ", _re.sub(r"^[\s:–\-]+", "", content)).strip()[:600]

    challenges = extract(draft, r"challenges?|barriers?|constraints?", r"next\s+steps?|way\s+forward|financial") if draft else ""
    next_steps = extract(draft, r"next\s+steps?|way\s+forward|upcoming\s+activities", r"evidence|conclusion|thank|challenges?") if draft else ""

    p = doc.add_paragraph()
    styled_run(p, "Challenges. ", size=11, bold=True, color=DBROWN)
    styled_run(p, challenges or "To be documented in consultation with the field team.", size=11, color=BODY)
    p = doc.add_paragraph()
    styled_run(p, "Next Steps. ", size=11, bold=True, color=DBROWN)
    styled_run(p, next_steps or "To be confirmed based on project progress and funder guidance.", size=11, color=BODY)

    # Footer
    doc.add_paragraph()
    foot = doc.add_paragraph()
    foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    styled_run(foot, "Vision Empower Trust  ·  www.visionempower.in", size=9, color=MUTED, italic=True)

    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    size = out.stat().st_size
    if size < 6000:
        raise RuntimeError(f"Generated DOCX is suspiciously small ({size} bytes)")
    print(f"WROTE: {out} ({size:,} bytes)")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        traceback.print_exc()
        sys.exit(1)
