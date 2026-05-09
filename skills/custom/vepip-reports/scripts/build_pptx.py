#!/usr/bin/env python3
"""Vision Empower premium funder report — 9-slide PPTX builder.

Usage:
    python build_pptx.py --data /path/to/data.json --output /path/to/out.pptx

Mirrors the pptxgenjs design used in the Next.js app: dark-brown cover
with braille texture, gold accents, deliverable progress bars, vertical
activity timeline, testimonial cards, geographic hero, financial bars,
and a dark thank-you panel on the closing slide.

The data.json schema is documented in the skill README.
"""
from __future__ import annotations

import argparse
import base64
import json
import sys
import traceback
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Any
from urllib.request import Request, urlopen

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _bootstrap import ensure  # type: ignore

ensure({"pptx": "python-pptx", "PIL": "Pillow"})

from pptx import Presentation  # noqa: E402
from pptx.dml.color import RGBColor  # noqa: E402
from pptx.enum.shapes import MSO_SHAPE  # noqa: E402
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN  # noqa: E402
from pptx.oxml.ns import qn  # noqa: E402
from pptx.util import Emu, Inches, Pt  # noqa: E402

# ── Brand palette ─────────────────────────────────────────────────────────────
DBROWN = RGBColor(0x2A, 0x15, 0x08)
GOLD   = RGBColor(0xC4, 0x9A, 0x32)
LGOLD  = RGBColor(0xED, 0xD9, 0x8A)
CREAM  = RGBColor(0xF7, 0xF3, 0xEE)
CARDF  = RGBColor(0xF5, 0xED, 0xD0)
BODY   = RGBColor(0x3D, 0x20, 0x10)
MUTED  = RGBColor(0x9B, 0x7B, 0x5A)
BORDER = RGBColor(0xE8, 0xDD, 0xD0)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
RED    = RGBColor(0xC0, 0x39, 0x2B)
GREEN  = RGBColor(0x2D, 0x7D, 0x46)
BLUE   = RGBColor(0x4A, 0x90, 0xD9)
SUCCESS= RGBColor(0x4C, 0xAF, 0x50)
WARN   = RGBColor(0xE5, 0x39, 0x35)


# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt_money(n: float | int | None) -> str:
    if n is None:
        return "—"
    n = float(n)
    if n >= 1e7:    return f"₹{n / 1e7:.1f}Cr"
    if n >= 1e5:    return f"₹{n / 1e5:.1f}L"
    if n >= 1e3:    return f"₹{n / 1e3:.0f}K"
    return f"₹{int(n):,}".replace(",", ",")


def fmt_date(d: str | None) -> str:
    if not d:
        return "—"
    try:
        dt = datetime.fromisoformat(d.replace("Z", "+00:00"))
        return dt.strftime("%-d %b %Y") if sys.platform != "win32" else dt.strftime("%#d %b %Y")
    except Exception:
        return d


def truncate(s: str | None, n: int) -> str:
    if not s:
        return ""
    return s[: n - 1] + "…" if len(s) > n else s


def project_duration_months(start: str, end: str) -> int:
    try:
        s = datetime.fromisoformat(start.replace("Z", "+00:00"))
        e = datetime.fromisoformat(end.replace("Z", "+00:00"))
        return max(1, (e.year - s.year) * 12 + (e.month - s.month))
    except Exception:
        return 0


def fetch_image(url: str) -> bytes | None:
    if not url:
        return None
    if url.startswith("data:"):
        # data:image/png;base64,...
        try:
            _, b64 = url.split(",", 1)
            return base64.b64decode(b64)
        except Exception:
            return None
    try:
        req = Request(url, headers={"User-Agent": "vepip-report-bot/1.0"})
        with urlopen(req, timeout=8) as r:
            return r.read()
    except Exception:
        return None


# ── Shape helpers ─────────────────────────────────────────────────────────────
def add_rect(slide, x, y, w, h, fill: RGBColor, line: RGBColor | None = None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = Pt(0.5)
    shape.shadow.inherit = False
    return shape


def add_ellipse(slide, x, y, w, h, fill: RGBColor):
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    return shape


def add_text(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    text: str,
    *,
    size: float = 12,
    color: RGBColor = BODY,
    bold: bool = False,
    italic: bool = False,
    font: str = "Calibri",
    align: str = "left",
    valign: str = "top",
    line_spacing: float | None = None,
    char_spacing: float | None = None,
):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    tf.vertical_anchor = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}[valign]
    p = tf.paragraphs[0]
    p.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}[align]
    if line_spacing:
        p.line_spacing = line_spacing
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    if char_spacing:
        rPr = run._r.get_or_add_rPr()
        rPr.set("spc", str(int(char_spacing * 100)))
    return box


def add_image(slide, x, y, w, h, data: bytes):
    return slide.shapes.add_picture(BytesIO(data), Inches(x), Inches(y), width=Inches(w), height=Inches(h))


def add_braille_dots(slide, sx: float, sy: float, color: RGBColor = GOLD):
    pattern = [
        (0, 0), (0.18, 0), (0.36, 0), (0.54, 0), (0.72, 0), (0.90, 0),
        (0, 0.22), (0.18, 0.22), (0.36, 0.22), (0.72, 0.22),
    ]
    for dx, dy in pattern:
        add_ellipse(slide, sx + dx, sy + dy, 0.09, 0.09, color)


# ── Slide chrome ──────────────────────────────────────────────────────────────
def add_slide_chrome(slide, eyebrow: str, title: str, num: int, total: int = 9):
    add_rect(slide, 0, 0, 0.065, 7.5, DBROWN)  # left bar
    # ghost slide number
    add_text(slide, 7.8, -0.8, 5.5, 5, str(num).zfill(2),
             size=280, color=GOLD, bold=True, font="Georgia", align="right")
    add_text(slide, 0.65, 0.22, 12, 0.28, eyebrow,
             size=8.5, color=GOLD, bold=True, char_spacing=5)
    add_text(slide, 0.65, 0.52, 11.7, 0.72, title,
             size=28, color=DBROWN, bold=True, font="Georgia")
    add_rect(slide, 0, 7.11, 13.33, 0.04, GOLD)
    add_text(slide, 0.65, 7.18, 6, 0.22, "Vision Empower", size=7.5, color=MUTED)
    add_text(slide, 11.5, 7.18, 1.6, 0.22, f"{num} / {total}", size=7.5, color=MUTED, align="right")


# ── Slide 1: Cover ────────────────────────────────────────────────────────────
def build_cover(prs: Presentation, project: dict, report_type: str,
                period_start: str, period_end: str, funder_logo: bytes | None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = DBROWN

    is_quarterly = report_type == "quarterly"
    period = (
        f"{fmt_date(period_start)} – {fmt_date(period_end)}"
        if is_quarterly
        else f"{fmt_date(project.get('startDate'))} – {fmt_date(project.get('endDate'))}"
    )

    # Braille texture
    for col in range(8):
        for row in range(10):
            add_ellipse(slide, 6.8 + col * 0.78, 0.35 + row * 0.62, 0.1, 0.1, GOLD)

    # VE logo chevron
    s1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.55), Inches(0.13), Inches(1.7))
    s1.fill.solid(); s1.fill.fore_color.rgb = CREAM; s1.line.fill.background(); s1.rotation = -22
    s2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.88), Inches(0.55), Inches(0.13), Inches(1.7))
    s2.fill.solid(); s2.fill.fore_color.rgb = CREAM; s2.line.fill.background(); s2.rotation = 22
    add_text(slide, 1.15, 0.85, 3.5, 0.32, "Vision Empower", size=10, color=CREAM)

    # Eyebrow + display words
    add_text(slide, 0.55, 3.0, 7.5, 0.3,
             "QUARTERLY FUNDER REPORT" if is_quarterly else "FULL PROJECT REPORT",
             size=9.5, color=GOLD, bold=True, char_spacing=7)
    add_text(slide, 0.45, 3.38, 6.5, 1.05, "Funder",
             size=66, color=CREAM, italic=True, font="Georgia")
    add_text(slide, 0.45, 4.5, 6.5, 1.0, "Update",
             size=66, color=LGOLD, bold=True, font="Georgia")

    # Generated date
    add_text(slide, 0.55, 5.75, 6, 0.25,
             f"Generated {datetime.now().strftime('%-d %B %Y') if sys.platform != 'win32' else datetime.now().strftime('%#d %B %Y')}",
             size=8.5, color=GOLD)

    # White info band
    add_rect(slide, 0, 6.1, 13.33, 1.4, WHITE)
    add_rect(slide, 0, 6.1, 13.33, 0.045, GOLD)
    add_text(slide, 0.6, 6.22, 7.5, 0.38, truncate(project.get("name", ""), 65),
             size=15, color=DBROWN, bold=True, font="Georgia")
    add_text(slide, 0.6, 6.62, 9, 0.28,
             f"{project.get('funderName', '')}  ·  {period}  ·  {fmt_money(project.get('grantAmount'))}",
             size=9.5, color=MUTED)
    states = project.get("states") or []
    if states:
        add_text(slide, 0.6, 6.9, 9, 0.24,
                 f"States: {', '.join(states)}", size=9, color=MUTED, italic=True)

    if funder_logo:
        try:
            add_image(slide, 10.8, 6.22, 2.1, 1.0, funder_logo)
        except Exception:
            pass

    add_braille_dots(slide, 0.55, 7.1)


# ── Slide 2: Project Overview ─────────────────────────────────────────────────
def build_overview(prs: Presentation, project: dict, draft: str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_chrome(slide, "PROJECT OVERVIEW", "About This Project", 2)

    summary = project.get("summary") or ""
    if not summary and draft:
        lines = [l.strip() for l in draft.split("\n") if len(l.strip()) > 60]
        summary = " ".join(lines[:4])[:500]
    if not summary:
        states = ", ".join(project.get("states") or [])
        summary = (
            f"This report summarises progress made under the {project.get('name', '')} project "
            f"funded by {project.get('funderName', '')}. The project operates across {states} with "
            f"a focus on inclusive education for visually impaired students."
        )

    add_rect(slide, 0.65, 1.42, 0.09, 3.6, GOLD)
    add_text(slide, 0.88, 1.42, 7.2, 3.6, truncate(summary, 500),
             size=13.5, color=BODY, line_spacing=1.55)

    add_text(slide, 0.65, 5.18, 8.2, 0.28,
             f"Beneficiaries: Visually impaired children  ·  States: {', '.join(project.get('states') or [])}",
             size=9.5, color=MUTED, italic=True)

    # Right cards
    cx, cw = 9.1, 3.6
    # Card 1 — grant
    add_rect(slide, cx, 1.42, cw, 1.7, CARDF)
    add_rect(slide, cx, 1.42, cw, 0.055, GOLD)
    add_text(slide, cx + 0.1, 1.6, cw - 0.2, 0.9, fmt_money(project.get("grantAmount")),
             size=40, color=DBROWN, bold=True, font="Georgia", align="center")
    add_text(slide, cx + 0.1, 2.55, cw - 0.2, 0.4, "TOTAL GRANT",
             size=9, color=MUTED, align="center", char_spacing=2, bold=True)
    # Card 2 — duration
    months = project_duration_months(project.get("startDate", ""), project.get("endDate", ""))
    add_rect(slide, cx, 3.32, cw, 1.7, CARDF)
    add_rect(slide, cx, 3.32, cw, 0.055, GOLD)
    add_text(slide, cx + 0.1, 3.5, cw - 0.2, 0.9, str(months),
             size=40, color=DBROWN, bold=True, font="Georgia", align="center")
    add_text(slide, cx + 0.1, 4.45, cw - 0.2, 0.4, "PROJECT DURATION (MONTHS)",
             size=9, color=MUTED, align="center", char_spacing=2, bold=True)


# ── Slide 3: Key Impact ───────────────────────────────────────────────────────
def build_key_metrics(prs: Presentation, project: dict):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_chrome(slide, "IMPACT", "Key Impact at a Glance", 3)

    deliv_done = project.get("deliverablesDone") or 0
    deliv_total = project.get("deliverablesTotal") or 0
    spent = project.get("spentBudget") or 0
    approved = project.get("approvedBudget") or 0
    deliv_pct = round(deliv_done / deliv_total * 100) if deliv_total else 0
    budget_pct = round(spent / approved * 100) if approved else 0

    activities = project.get("activities") or []
    teachers = sum(a.get("teachersReached") or 0 for a in activities)
    students = sum(a.get("studentsReached") or 0 for a in activities)

    cards = [
        (f"{deliv_pct}%", "DELIVERABLES", f"{deliv_done} of {deliv_total} complete"),
        (f"{budget_pct}%", "BUDGET USED", f"{fmt_money(spent)} of {fmt_money(approved)}"),
        (f"{teachers:,}", "TEACHERS REACHED", "across all activities"),
        (f"{students:,}", "STUDENTS REACHED", "through project activities"),
    ]
    accents = [GOLD, DBROWN, BLUE, SUCCESS]

    cw, ch, gap = 5.8, 2.6, 0.16
    sx, sy = 0.65, 1.38

    for i, ((value, label, sublabel), accent) in enumerate(zip(cards, accents)):
        col, row = i % 2, i // 2
        x = sx + col * (cw + gap)
        y = sy + row * (ch + gap)
        add_rect(slide, x, y, cw, ch, CARDF)
        add_rect(slide, x, y, 0.07, ch, accent)
        add_text(slide, x + 0.18, y + 0.28, cw - 0.18, 1.35, value,
                 size=60, color=DBROWN, bold=True, font="Georgia", align="center")
        add_text(slide, x + 0.18, y + 1.72, cw - 0.18, 0.65, label,
                 size=11, color=BODY, align="center")
        add_text(slide, x + 0.18, y + 2.28, cw - 0.18, 0.22, sublabel,
                 size=8.5, color=MUTED, align="center", char_spacing=1)


# ── Slide 4: Deliverables ─────────────────────────────────────────────────────
def build_deliverables(prs: Presentation, project: dict):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_chrome(slide, "DELIVERABLES", "Progress Against Plan", 4)

    items = project.get("deliverables") or []
    if not items:
        add_text(slide, 0.65, 2.5, 12, 0.5, "No deliverables recorded for this project.",
                 size=13, color=MUTED)
        return

    add_text(slide, 0.65, 1.38, 8, 0.32,
             f"{project.get('deliverablesDone') or 0} of {project.get('deliverablesTotal') or 0} deliverables complete",
             size=11, color=MUTED)

    add_rect(slide, 0.65, 1.78, 12.03, 0.38, DBROWN)
    headers = [
        ("DELIVERABLE", 0.75, 5.2, "left"),
        ("TARGET",      6.0,  1.6, "center"),
        ("ACHIEVED",    7.65, 1.7, "center"),
        ("PROGRESS",    9.4,  2.0, "center"),
        ("STATUS",      11.45,1.1, "center"),
    ]
    for label, x, w, align in headers:
        add_text(slide, x, 1.82, w, 0.3, label,
                 size=8.5, color=LGOLD, bold=True, align=align, char_spacing=1)

    status_color = {
        "completed": SUCCESS, "in_progress": GOLD,
        "overdue": WARN, "not_started": MUTED,
    }
    for i, d in enumerate(items[:8]):
        row_y = 2.16 + i * 0.56
        achieved = d.get("achieved") or 0
        target = d.get("target") or 0
        pct = round(achieved / target * 100) if target else (100 if d.get("status") == "completed" else 0)
        bg = WHITE if i % 2 == 0 else CREAM
        sc = status_color.get(d.get("status", ""), MUTED)
        unit = d.get("unit", "")

        add_rect(slide, 0.65, row_y, 12.03, 0.55, bg)
        add_text(slide, 0.75, row_y + 0.1, 5.1, 0.35, truncate(d.get("title", ""), 40),
                 size=11, color=DBROWN, bold=(d.get("status") == "completed"))
        add_text(slide, 6.0, row_y + 0.1, 1.6, 0.35,
                 f"{target}{(' ' + unit) if unit else ''}" if target else "—",
                 size=11, color=BODY, align="center")
        ach_str = f"{achieved}{(' ' + unit) if unit else ''}" if achieved else "—"
        add_text(slide, 7.65, row_y + 0.1, 1.7, 0.35, ach_str,
                 size=11, color=GOLD if pct >= 100 else BODY, bold=(pct >= 100), align="center")

        # Progress bar
        add_rect(slide, 9.4, row_y + 0.38, 2.0, 0.09, BORDER)
        fill_w = min(pct / 100, 1) * 2.0
        if fill_w > 0:
            add_rect(slide, 9.4, row_y + 0.38, fill_w, 0.09, GOLD if pct >= 100 else DBROWN)
        add_text(slide, 11.1, row_y + 0.12, 0.32, 0.3, f"{pct}%",
                 size=8.5, color=GOLD if pct >= 100 else MUTED, align="right")
        add_ellipse(slide, 11.5, row_y + 0.2, 0.15, 0.15, sc)


# ── Slide 5: Activities ───────────────────────────────────────────────────────
def build_activities(prs: Presentation, project: dict, gallery_imgs: list[bytes | None]):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_chrome(slide, "ACTIVITIES", "Field Activities & Reach", 5)

    activities = (project.get("activities") or [])[:6]
    if not activities:
        add_text(slide, 0.65, 2.5, 12, 0.5, "No activities recorded for this project.",
                 size=13, color=MUTED)
        return

    has_photos = any(g for g in gallery_imgs)
    list_w = 6.5 if has_photos else 11.5
    max_acts = 5 if has_photos else 6

    for i, act in enumerate(activities[:max_acts]):
        item_y = 1.4 + i * 1.05
        is_last = i == min(len(activities), max_acts) - 1

        add_ellipse(slide, 0.65, item_y + 0.17, 0.22, 0.22, GOLD)
        if not is_last:
            add_rect(slide, 0.735, item_y + 0.38, 0.07, 1.05, BORDER)

        add_text(slide, 1.05, item_y, list_w - 0.45, 0.38,
                 truncate(act.get("title", ""), 60),
                 size=12, color=DBROWN, bold=True)

        meta = "  ·  ".join([p for p in [
            fmt_date(act.get("activityDate")) if act.get("activityDate") else "",
            act.get("state") or act.get("location") or "",
        ] if p])
        if meta:
            add_text(slide, 1.05, item_y + 0.38, list_w - 0.45, 0.22, meta,
                     size=9, color=GOLD)

        reaches = []
        if act.get("teachersReached"): reaches.append(f"{act['teachersReached']} teachers")
        if act.get("studentsReached"): reaches.append(f"{act['studentsReached']} students")
        if act.get("schoolsReached"):  reaches.append(f"{act['schoolsReached']} schools")
        reach_str = "  ·  ".join(reaches) if reaches else truncate(act.get("notes") or "", 65)
        if reach_str:
            add_text(slide, 1.05, item_y + 0.6, list_w - 0.45, 0.22, reach_str,
                     size=10, color=BODY)

    if has_photos:
        slots = [g for g in gallery_imgs if g][:4]
        positions = [(7.4, 1.38), (10.28, 1.38), (7.4, 3.55), (10.28, 3.55)]
        gallery_meta = (project.get("gallery") or [])
        for i, img in enumerate(slots):
            ix, iy = positions[i]
            try:
                add_image(slide, ix, iy, 2.65, 2.0, img)
            except Exception:
                continue
            cap = (gallery_meta[i] if i < len(gallery_meta) else {}).get("caption") if gallery_meta else None
            if cap:
                add_rect(slide, ix, iy + 2.0 - 0.3, 2.65, 0.3, DBROWN)
                add_text(slide, ix + 0.05, iy + 2.0 - 0.28, 2.65 - 0.1, 0.28,
                         truncate(cap, 40), size=7.5, color=WHITE, align="center")


# ── Slide 6: Stories ──────────────────────────────────────────────────────────
def build_testimonials(prs: Presentation, project: dict):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_chrome(slide, "STORIES", "Stories from the Field", 6)

    testimonials = project.get("testimonials") or []
    if not testimonials:
        for a in (project.get("activities") or []):
            if a.get("testimonial"):
                testimonials.append({
                    "content": a["testimonial"],
                    "author": a.get("testimonialBy") or a.get("title", ""),
                    "role": a.get("state") or "",
                })
                if len(testimonials) >= 2:
                    break

    if not testimonials:
        activities = project.get("activities") or []
        totals = [
            (sum(a.get("teachersReached") or 0 for a in activities), "Teachers Reached"),
            (sum(a.get("studentsReached") or 0 for a in activities), "Students Reached"),
            (sum(a.get("schoolsReached") or 0 for a in activities), "Schools Reached"),
        ]
        for i, (v, l) in enumerate(totals):
            cx = 0.65 + i * 4.22
            add_rect(slide, cx, 2.2, 3.85, 2.2, CARDF)
            add_rect(slide, cx, 2.2, 3.85, 0.055, GOLD)
            add_text(slide, cx + 0.1, 2.38, 3.65, 1.2, f"{v:,}",
                     size=52, color=DBROWN, bold=True, font="Georgia", align="center")
            add_text(slide, cx + 0.1, 3.62, 3.65, 0.5, l,
                     size=10, color=BODY, align="center")
        return

    add_text(slide, 0.2, -0.3, 3.5, 3.5, "“",
             size=240, color=GOLD, font="Georgia")

    primary = testimonials[0]
    add_rect(slide, 0.65, 1.38, 12.03, 3.5, CARDF)
    add_rect(slide, 0.65, 1.38, 0.09, 3.5, GOLD)
    add_text(slide, 0.9, 1.58, 11.5, 2.5, truncate(primary.get("content", ""), 320),
             size=15.5, color=DBROWN, italic=True, font="Georgia",
             line_spacing=1.6, valign="middle")
    add_rect(slide, 0.9, 4.45, 11.2, 0.03, GOLD)
    role = primary.get("role")
    add_text(slide, 0.9, 4.55, 11.5, 0.3,
             f"— {primary.get('author', '')}{', ' + role if role else ''}",
             size=10.5, color=MUTED, align="right")

    if len(testimonials) > 1:
        q2 = testimonials[1]
        add_rect(slide, 0.65, 5.0, 12.03, 1.7, CREAM)
        add_rect(slide, 0.65, 5.0, 0.09, 1.7, GOLD)
        add_text(slide, 0.9, 5.1, 11.5, 0.9, truncate(q2.get("content", ""), 200),
                 size=11.5, color=BODY, italic=True, line_spacing=1.4)
        role2 = q2.get("role")
        add_text(slide, 0.9, 5.98, 11.5, 0.3,
                 f"— {q2.get('author', '')}{', ' + role2 if role2 else ''}",
                 size=9.5, color=MUTED, align="right")


# ── Slide 7: Geographic Reach ─────────────────────────────────────────────────
def build_geographic(prs: Presentation, project: dict):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_chrome(slide, "REACH", "Geographic Coverage", 7)

    states = project.get("states") or []
    add_text(slide, 0.65, 1.38, 3.8, 2.5, str(len(states)),
             size=160, color=GOLD, bold=True, font="Georgia",
             align="center", valign="middle")
    add_text(slide, 0.65, 3.9, 3.8, 0.3, "STATES COVERED",
             size=9, color=MUTED, align="center", char_spacing=2, bold=True)

    add_rect(slide, 4.6, 1.38, 0.03, 5.4, BORDER)

    per_row, box_w, box_h, gap_x, gap_y = 3, 2.5, 0.68, 0.18, 0.15
    grid_x, grid_y = 4.9, 1.38
    for i, s in enumerate(states):
        col, row = i % per_row, i // per_row
        bx = grid_x + col * (box_w + gap_x)
        by = grid_y + row * (box_h + gap_y)
        if by + box_h > 6.7:
            break
        add_rect(slide, bx, by, box_w, box_h, WHITE, line=BORDER)
        add_rect(slide, bx, by, 0.07, box_h, GOLD)
        add_text(slide, bx + 0.18, by + 0.14, box_w - 0.28, 0.4, s,
                 size=13, color=DBROWN, bold=True)

    activities = project.get("activities") or []
    teachers = sum(a.get("teachersReached") or 0 for a in activities)
    students = sum(a.get("studentsReached") or 0 for a in activities)
    schools  = sum(a.get("schoolsReached") or 0 for a in activities)

    chips = [
        (f"{teachers:,}", "TEACHERS"),
        (f"{students:,}", "STUDENTS"),
        (f"{schools:,}",  "SCHOOLS"),
        (str(len(activities)), "ACTIVITIES"),
    ]
    for i, (n, l) in enumerate(chips):
        cx = 4.9 + i * 2.1
        add_text(slide, cx, 6.25, 1.8, 0.38, n,
                 size=18, color=GOLD, bold=True, font="Georgia", align="center")
        add_text(slide, cx, 6.63, 1.8, 0.22, l,
                 size=7.5, color=MUTED, align="center", char_spacing=1, bold=True)
        if i < len(chips) - 1:
            add_rect(slide, cx + 1.8 + 0.12, 6.3, 0.015, 0.55, GOLD)


# ── Slide 8: Financial Summary ────────────────────────────────────────────────
def build_financials(prs: Presentation, project: dict):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_chrome(slide, "FINANCIALS", "Budget & Utilisation", 8)

    budgets = project.get("budgets") or []
    approved = project.get("approvedBudget") or 0
    spent    = project.get("spentBudget") or 0
    balance  = approved - spent
    util_pct = round(spent / approved * 100) if approved else 0

    cw, ch = 2.85, 1.55
    gap = (13.33 - 0.65 * 2 - 4 * cw) / 3
    sx, cy = 0.65, 1.38

    cards = [
        (fmt_money(approved), "APPROVED BUDGET"),
        (fmt_money(spent),    "AMOUNT SPENT"),
        (fmt_money(balance),  "BALANCE REMAINING"),
        (f"{util_pct}%",      "UTILISATION %"),
    ]
    for i, (v, l) in enumerate(cards):
        cx = sx + i * (cw + gap)
        add_rect(slide, cx, cy, cw, ch, CARDF)
        add_rect(slide, cx, cy, cw, 0.055, GOLD)
        add_text(slide, cx + 0.1, cy + 0.18, cw - 0.2, 0.82, v,
                 size=28, color=DBROWN, bold=True, font="Georgia", align="center")
        add_text(slide, cx + 0.1, cy + 1.06, cw - 0.2, 0.38, l,
                 size=8, color=MUTED, align="center", char_spacing=1, bold=True)

    add_text(slide, 0.65, 3.12, 12, 0.28, "Budget breakdown by category",
             size=10, color=MUTED)

    if not budgets:
        add_text(slide, 0.65, 3.5, 12, 0.5, "No budget categories recorded.",
                 size=13, color=MUTED)
        return

    for i, b in enumerate(budgets[:8]):
        row_y = 3.42 + i * 0.52
        appr  = b.get("approvedAmount") or 0
        sp    = b.get("spentAmount") or 0
        utilp = round(sp / appr * 100) if appr else 0

        add_text(slide, 0.65, row_y, 3.8, 0.38, truncate(b.get("name", ""), 30),
                 size=11, color=DBROWN)
        add_text(slide, 4.55, row_y, 2.3, 0.38,
                 f"{fmt_money(sp)} / {fmt_money(appr)}", size=9, color=MUTED, align="right")
        add_rect(slide, 7.0, row_y + 0.14, 5.2, 0.18, BORDER)
        fw = min(utilp / 100, 1) * 5.2
        if fw > 0:
            add_rect(slide, 7.0, row_y + 0.14, fw, 0.18, GOLD if utilp >= 90 else DBROWN)
        add_text(slide, 12.3, row_y, 0.7, 0.38, f"{utilp}%",
                 size=10, color=GOLD if utilp >= 90 else BODY, align="right",
                 bold=(utilp >= 90))


# ── Slide 9: Way Forward ──────────────────────────────────────────────────────
def build_way_forward(prs: Presentation, project: dict, draft: str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    def extract(text: str, start_re: str, stop_re: str) -> str:
        import re as _re
        parts = _re.split(start_re, text, flags=_re.IGNORECASE, maxsplit=1)
        if len(parts) < 2:
            return ""
        content = parts[1]
        m = _re.search(stop_re, content, flags=_re.IGNORECASE)
        if m:
            content = content[: m.start()]
        content = _re.sub(r"^[\s:–\-]+", "", content)
        content = _re.sub(r"\s+", " ", content).strip()
        return content[:380]

    challenges = extract(draft, r"challenges?|barriers?|constraints?|difficulties",
                         r"next\s+steps?|way\s+forward|financial|evidence|recommendation") if draft else ""
    next_steps = extract(draft, r"next\s+steps?|way\s+forward|plan\s+for\s+next|upcoming\s+activities",
                         r"evidence\s+gaps?|conclusion|closing|financial|thank|challenges?") if draft else ""
    challenges = challenges or "To be documented in consultation with the field team."
    next_steps = next_steps or "To be confirmed based on project progress and funder guidance."

    # Left chrome
    add_rect(slide, 0, 0, 0.065, 7.5, DBROWN)
    add_text(slide, 0.65, 0.22, 12, 0.28, "WAY FORWARD",
             size=8.5, color=GOLD, bold=True, char_spacing=5)
    add_text(slide, 0.65, 0.52, 11.7, 0.72, "Challenges & Next Steps",
             size=28, color=DBROWN, bold=True, font="Georgia")

    col_w, col_h = 3.3, 4.55
    col1_x, col2_x, col_y = 0.65, 4.15, 1.42

    # Challenges
    add_rect(slide, col1_x, col_y, col_w, col_h, CREAM)
    add_rect(slide, col1_x, col_y, col_w, 0.055, RED)
    add_text(slide, col1_x + 0.15, col_y + 0.12, col_w - 0.2, 0.45, "Challenges",
             size=16, color=DBROWN, bold=True, font="Georgia")
    add_rect(slide, col1_x + 0.15, col_y + 0.62, col_w - 0.3, 0.025, BORDER)
    add_text(slide, col1_x + 0.15, col_y + 0.75, col_w - 0.25, col_h - 0.92, challenges,
             size=11.5, color=BODY, line_spacing=1.5)

    # Next steps
    add_rect(slide, col2_x, col_y, col_w, col_h, CREAM)
    add_rect(slide, col2_x, col_y, col_w, 0.055, GREEN)
    add_text(slide, col2_x + 0.15, col_y + 0.12, col_w - 0.2, 0.45, "Next Steps",
             size=16, color=DBROWN, bold=True, font="Georgia")
    add_rect(slide, col2_x + 0.15, col_y + 0.62, col_w - 0.3, 0.025, BORDER)
    add_text(slide, col2_x + 0.15, col_y + 0.75, col_w - 0.25, col_h - 0.92, next_steps,
             size=11.5, color=BODY, line_spacing=1.5)

    # Footer
    add_rect(slide, 0, 7.11, 7.63, 0.04, GOLD)
    add_text(slide, 0.65, 7.18, 5, 0.22, "Vision Empower", size=7.5, color=MUTED)
    add_text(slide, 6.3, 7.18, 1.2, 0.22, "9 / 9", size=7.5, color=MUTED, align="right")

    # Right dark panel
    add_rect(slide, 7.63, 0, 5.7, 7.5, DBROWN)
    for col in range(3):
        for row in range(6):
            add_ellipse(slide, 8.7 + col * 0.6, 0.5 + row * 0.85, 0.1, 0.1, GOLD)

    add_text(slide, 7.75, 1.6, 5.4, 1.2, "Thank You",
             size=52, color=CREAM, italic=True, font="Georgia", align="center")
    add_text(slide, 7.75, 2.9, 5.4, 0.38, "for your continued partnership",
             size=11, color=GOLD, align="center")
    add_text(slide, 7.75, 3.3, 5.4, 0.7, truncate(project.get("name", ""), 55),
             size=15, color=CREAM, bold=True, font="Georgia", align="center")
    add_rect(slide, 8.3, 4.1, 4.3, 0.025, CREAM)

    add_text(slide, 7.75, 4.25, 5.4, 0.4, "Vision Empower Trust",
             size=11, color=CREAM, bold=True, align="center")
    add_text(slide, 7.75, 4.6, 5.4, 0.4, "Enabling inclusive education for visually impaired children",
             size=10, color=CREAM, italic=True, align="center")
    add_text(slide, 7.75, 5.1, 5.4, 0.32, "www.visionempower.in",
             size=10, color=GOLD, bold=True, align="center")
    add_text(slide, 7.75, 5.75, 5.4, 0.3, f"Funded by {project.get('funderName', '')}",
             size=9, color=MUTED, align="center")


# ── Main ──────────────────────────────────────────────────────────────────────
def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--data", required=True, help="JSON file with project + report data")
    parser.add_argument("--output", required=True, help="Path to write .pptx")
    parser.add_argument("--report-type", default="quarterly", choices=["quarterly", "full"])
    parser.add_argument("--period-start", default="")
    parser.add_argument("--period-end", default="")
    parser.add_argument("--draft", default="", help="Optional narrative draft for challenges/next-steps extraction")
    args = parser.parse_args(argv)

    data: dict[str, Any] = json.loads(Path(args.data).read_text(encoding="utf-8"))
    project = data.get("project") or data  # accept either { project: {...} } or flat

    # Period for quarterly reports — fallback
    period_start = args.period_start or data.get("periodStart") or project.get("startDate") or ""
    period_end   = args.period_end   or data.get("periodEnd")   or project.get("endDate")   or ""
    draft        = args.draft or data.get("draft") or ""

    funder_logo = fetch_image(project.get("funderLogoUrl", ""))
    gallery_imgs = [fetch_image((g or {}).get("url", "")) for g in (project.get("gallery") or [])[:4]]

    prs = Presentation()
    prs.slide_width  = Inches(13.333)
    prs.slide_height = Inches(7.5)
    prs.core_properties.title = f"{project.get('name', 'VEPIP Report')} — Funder Report"
    prs.core_properties.author = "Vision Empower"

    build_cover(prs, project, args.report_type, period_start, period_end, funder_logo)
    build_overview(prs, project, draft)
    build_key_metrics(prs, project)
    build_deliverables(prs, project)
    build_activities(prs, project, gallery_imgs)
    build_testimonials(prs, project)
    build_geographic(prs, project)
    build_financials(prs, project)
    build_way_forward(prs, project, draft)

    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    size = out.stat().st_size
    if size < 8000:
        raise RuntimeError(f"Generated PPTX is suspiciously small ({size} bytes) — likely corrupt")
    print(f"WROTE: {out} ({size:,} bytes, 9 slides)")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        traceback.print_exc()
        sys.exit(1)
