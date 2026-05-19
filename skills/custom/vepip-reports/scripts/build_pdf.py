#!/usr/bin/env python3
"""Vision Empower premium funder report — vibe-driven PDF builder (reportlab).

The selected vibe drives palette + typography + chart aesthetic + section
gating. Every page renders with a printed header (project name + funder),
a printed footer (page numbers + Vision Empower mark), a cover page with
a colored vibe band, charts (donut for budget utilisation, horizontal bar
for deliverable progress), and KPI tiles drawn as table cells so they
print without aliasing.

Usage:
    python build_pdf.py --data data.json --output out.pdf --vibe dark-premium
"""
from __future__ import annotations

import argparse
import json
import sys
import traceback
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _bootstrap import ensure  # type: ignore

ensure({"reportlab": "reportlab"})

from reportlab.graphics.charts.barcharts import HorizontalBarChart  # noqa: E402
from reportlab.graphics.charts.piecharts import Pie  # noqa: E402
from reportlab.graphics.shapes import Drawing, Rect, String  # noqa: E402
from reportlab.lib.colors import HexColor, Color  # noqa: E402
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT  # noqa: E402
from reportlab.lib.pagesizes import A4  # noqa: E402
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet  # noqa: E402
from reportlab.lib.units import cm, mm  # noqa: E402
from reportlab.pdfbase import pdfmetrics  # noqa: E402
from reportlab.pdfbase.ttfonts import TTFont  # noqa: E402
from reportlab.pdfgen import canvas as pdfcanvas  # noqa: E402
from reportlab.platypus import (  # noqa: E402
    BaseDocTemplate, Frame, PageBreak, PageTemplate, Paragraph, Spacer,
    Table, TableStyle, KeepTogether, Flowable,
)

from vibes import Vibe, get_vibe  # noqa: E402


WHITE = HexColor("#FFFFFF")
BLACK = HexColor("#000000")


# ── ReportLab font mapping ────────────────────────────────────────────────────
# ReportLab only ships the base14 Type1 fonts. Attempt to register Inter /
# Source Serif if they're present in common system font dirs (best-effort,
# silent fail), else fall back to base14 Helvetica/Times preserving intent.
_SERIF_BASE = ("Times-Roman", "Times-Bold", "Times-Italic", "Times-BoldItalic")
_SANS_BASE = ("Helvetica", "Helvetica-Bold", "Helvetica-Oblique", "Helvetica-BoldOblique")

_REGISTERED: dict[str, bool] = {}


def _try_register_font(name: str, candidates: list[str]) -> bool:
    if name in _REGISTERED:
        return _REGISTERED[name]
    for path in candidates:
        try:
            if Path(path).exists():
                pdfmetrics.registerFont(TTFont(name, path))
                _REGISTERED[name] = True
                return True
        except Exception:
            continue
    _REGISTERED[name] = False
    return False


def map_fonts(vibe: Vibe) -> dict[str, str]:
    """Pick ReportLab-safe fonts for this vibe. Honours intent (serif vs sans,
    bold/italic) while degrading gracefully when designer fonts are missing."""
    head = (vibe.typography.headline or "").lower()
    body = (vibe.typography.body or "").lower()

    def family_for(name: str) -> tuple[str, str, str, str]:
        if any(s in name for s in ("georgia", "fraunces", "times", "serif")):
            return _SERIF_BASE
        return _SANS_BASE

    h = family_for(head)
    b = family_for(body)
    return {
        "head_regular": h[0],
        "head_bold": h[1],
        "head_italic": h[2],
        "head_bold_italic": h[3],
        "body_regular": b[0],
        "body_bold": b[1],
        "body_italic": b[2],
    }


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


def _mix(c1: Color, c2: Color, t: float) -> Color:
    return Color(c1.red * (1 - t) + c2.red * t,
                 c1.green * (1 - t) + c2.green * t,
                 c1.blue * (1 - t) + c2.blue * t)


# ── Custom drawn flowables ────────────────────────────────────────────────────
class ColoredBand(Flowable):
    """A full-width coloured band used as the cover hero strip."""

    def __init__(self, width: float, height: float, fill: Color,
                 accent: Color, motif: str = "grid"):
        super().__init__()
        self.width = width
        self.height = height
        self.fill = fill
        self.accent = accent
        self.motif = motif

    def wrap(self, *_):
        return self.width, self.height

    def draw(self):
        c = self.canv
        c.saveState()
        c.setFillColor(self.fill)
        c.setStrokeColor(self.fill)
        c.rect(0, 0, self.width, self.height, stroke=0, fill=1)

        # Motif: subtle grid of dots or stripes for visual texture.
        c.setFillColor(_mix(self.fill, self.accent, 0.18))
        if self.motif == "grid":
            step = 14
            r = 1.0
            for y in range(int(step / 2), int(self.height), step):
                for x in range(int(step / 2), int(self.width), step):
                    c.circle(x, y, r, stroke=0, fill=1)
        elif self.motif == "stripes":
            c.setStrokeColor(_mix(self.fill, self.accent, 0.12))
            c.setLineWidth(0.4)
            for x in range(0, int(self.width), 10):
                c.line(x, 0, x + self.height, self.height)
        else:
            for y in range(6, int(self.height), 10):
                c.setFillColor(_mix(self.fill, self.accent, 0.10))
                c.rect(0, y, self.width, 0.6, stroke=0, fill=1)

        # Accent bar across the top
        c.setFillColor(self.accent)
        c.rect(0, self.height - 4, self.width, 4, stroke=0, fill=1)
        c.restoreState()


def make_donut(approved: float, spent: float, accent_hex: str,
               soft_hex: str, muted_hex: str, label_hex: str,
               body_font: str, label_font: str) -> Drawing:
    """A small donut chart showing budget utilisation. Caller supplies hex
    strings (without #) to keep colour decisions in the vibe layer."""
    d = Drawing(160, 160)
    spent = max(0, min(spent, approved or 0))
    remaining = max(0, (approved or 0) - spent)
    if (approved or 0) <= 0:
        # No budget data — draw an empty grey ring with a dash in the centre.
        pie = Pie()
        pie.x, pie.y = 20, 12
        pie.width = pie.height = 120
        pie.data = [1]
        pie.simpleLabels = 1
        pie.slices[0].fillColor = HexColor("#" + soft_hex)
        pie.slices[0].strokeColor = WHITE
        pie.slices[0].strokeWidth = 2
        d.add(pie)
        d.add(String(80, 80, "—", textAnchor="middle",
                     fontSize=18, fontName=label_font,
                     fillColor=HexColor("#" + label_hex)))
        return d
    pct = spent / approved * 100 if approved else 0
    pie = Pie()
    pie.x, pie.y = 20, 12
    pie.width = pie.height = 120
    pie.data = [spent, remaining]
    pie.simpleLabels = 1
    pie.labels = ["", ""]
    pie.slices[0].fillColor = HexColor("#" + accent_hex)
    pie.slices[0].strokeColor = WHITE
    pie.slices[0].strokeWidth = 2
    pie.slices[1].fillColor = HexColor("#" + soft_hex)
    pie.slices[1].strokeColor = WHITE
    pie.slices[1].strokeWidth = 2
    d.add(pie)
    # Cut out the centre with the page bg colour by overlaying a circle.
    d.add(Rect(60, 52, 40, 40, fillColor=WHITE, strokeColor=None,
               strokeWidth=0))  # square stand-in; we'll use a circle below
    # Add an actual circle hole — Rect placeholder above keeps reportlab
    # from clipping the labels; real hole drawn with a smaller white pie:
    hole = Pie()
    hole.x, hole.y = 44, 36
    hole.width = hole.height = 72
    hole.data = [1]
    hole.simpleLabels = 1
    hole.slices[0].fillColor = WHITE
    hole.slices[0].strokeColor = WHITE
    d.add(hole)
    # Percentage label in the centre.
    d.add(String(80, 78, f"{pct:.0f}%", textAnchor="middle",
                 fontSize=20, fontName=label_font,
                 fillColor=HexColor("#" + label_hex)))
    d.add(String(80, 62, "utilised", textAnchor="middle",
                 fontSize=8, fontName=body_font,
                 fillColor=HexColor("#" + muted_hex)))
    return d


def make_deliverable_bars(deliverables: list[dict], accent_hex: str,
                          track_hex: str, label_hex: str, body_font: str,
                          max_bars: int = 6) -> Drawing | None:
    """Horizontal bar chart showing % complete per deliverable."""
    items = []
    for d in deliverables[:max_bars]:
        target = d.get("target") or 0
        achieved = d.get("achieved") or 0
        pct = min(100, round(achieved / target * 100)) if target else (
            100 if d.get("status") == "completed" else 0
        )
        items.append((d.get("title", "")[:40], pct))
    if not items:
        return None
    n = len(items)
    height = max(80, n * 24 + 20)
    width = 460
    drawing = Drawing(width, height)
    bar_left = 180
    bar_max_w = width - bar_left - 50
    track_h = 8
    y = height - 18
    for label, pct in items:
        # Label on the left.
        drawing.add(String(bar_left - 8, y - 3, label, textAnchor="end",
                           fontSize=9, fontName=body_font,
                           fillColor=HexColor("#" + label_hex)))
        # Track.
        drawing.add(Rect(bar_left, y - track_h + 2, bar_max_w, track_h,
                         fillColor=HexColor("#" + track_hex),
                         strokeColor=None, rx=4, ry=4))
        # Filled portion.
        w = bar_max_w * pct / 100
        drawing.add(Rect(bar_left, y - track_h + 2, w, track_h,
                         fillColor=HexColor("#" + accent_hex),
                         strokeColor=None, rx=4, ry=4))
        # Percentage label at the right of the bar.
        drawing.add(String(bar_left + w + 6, y - 3, f"{pct}%",
                           fontSize=9, fontName=body_font,
                           fillColor=HexColor("#" + label_hex)))
        y -= 24
    return drawing


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--data", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--report-type", default="quarterly")
    parser.add_argument("--period-start", default="")
    parser.add_argument("--period-end", default="")
    parser.add_argument("--draft", default="")
    parser.add_argument("--vibe", default="editorial-serif")
    args = parser.parse_args(argv)

    data = json.loads(Path(args.data).read_text(encoding="utf-8"))
    project = data.get("project") or data
    if "budgets" not in project and project.get("budgetCategories"):
        project["budgets"] = project["budgetCategories"]
    if "activities" not in project and project.get("recentActivities"):
        project["activities"] = project["recentActivities"]
    if "deliverablesTotal" not in project or "deliverablesDone" not in project:
        delivs = project.get("deliverables") or []
        project.setdefault("deliverablesTotal", len(delivs))
        project.setdefault(
            "deliverablesDone",
            sum(1 for d in delivs if (d.get("achieved") or 0) >= (d.get("target") or 0) > 0),
        )
    if "approvedBudget" not in project or "spentBudget" not in project:
        cats = project.get("budgetCategories") or project.get("budgets") or []
        project.setdefault(
            "approvedBudget",
            sum((c.get("approvedAmount") or c.get("approved") or 0) for c in cats),
        )
        project.setdefault(
            "spentBudget",
            sum((c.get("spentAmount") or c.get("spent") or 0) for c in cats),
        )
    period_start = args.period_start or data.get("periodStart") or project.get("startDate") or ""
    period_end = args.period_end or data.get("periodEnd") or project.get("endDate") or ""
    draft = args.draft or data.get("draft") or ""

    vibe = get_vibe(args.vibe)
    fonts = map_fonts(vibe)
    pal = vibe.palette
    TEXT_PRIMARY = HexColor("#" + pal.as_hex("text_primary"))
    TEXT_SECONDARY = HexColor("#" + pal.as_hex("text_secondary"))
    TEXT_MUTED = HexColor("#" + pal.as_hex("text_muted"))
    ACCENT = HexColor("#" + pal.as_hex("accent"))
    ACCENT_SOFT = HexColor("#" + pal.as_hex("accent_soft"))
    SURFACE = HexColor("#" + pal.as_hex("surface"))
    SURFACE_ALT = HexColor("#" + pal.as_hex("surface_alt"))
    BORDER = HexColor("#" + pal.as_hex("border"))
    BG = HexColor("#" + pal.as_hex("bg"))

    # ── Section gating mirrors PPTX manifest ──────────────────────────────
    narrative = project.get("narrative") if isinstance(project.get("narrative"), dict) else {}
    summary_text = ((narrative.get("overview") or "").strip() if narrative else "") or project.get("summary") or ""
    deliverables = project.get("deliverables") or []
    activities = project.get("activities") or []
    testimonials = list(project.get("testimonials") or [])
    if not testimonials:
        for a in activities:
            if a.get("testimonial"):
                testimonials.append({
                    "content": a["testimonial"],
                    "author": a.get("testimonialBy") or a.get("title", ""),
                    "role": a.get("state") or "",
                })

    quote_threshold = 40 if vibe.layout.prefers_quotes else 100
    overview_min = 80 if vibe.layout.prefers_prose else 180
    way_forward_min_draft = 300 if vibe.layout.prefers_prose else 600

    has_overview = bool(summary_text and len(summary_text) >= overview_min) or bool(
        draft and len(draft) >= max(200, overview_min * 2)
    )
    has_deliverables = len(deliverables) >= 1
    has_activities = len(activities) >= 1
    has_stories = any(
        (t.get("content") or "").strip() and len((t.get("content") or "").strip()) >= quote_threshold
        for t in testimonials
    )
    approved = project.get("approvedBudget") or 0
    spent = project.get("spentBudget") or 0
    has_financials = approved > 0 or (vibe.layout.prefers_charts and (spent > 0 or bool(project.get("budgets"))))
    has_narrative_forward = (
        (narrative.get("challenges") or "").strip()
        or (narrative.get("way_forward") or "").strip()
    ) if narrative else False
    has_way_forward = has_narrative_forward or bool(draft and len(draft) >= way_forward_min_draft)

    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)

    # ── Paragraph styles ──────────────────────────────────────────────────
    base = getSampleStyleSheet()
    h_eyebrow = ParagraphStyle("Eyebrow", parent=base["Normal"],
                               textColor=ACCENT, fontSize=8.5, spaceAfter=2,
                               fontName=fonts["body_bold"])
    h_cover_eyebrow_dark = ParagraphStyle(
        "CoverEyebrow", parent=base["Normal"], textColor=ACCENT_SOFT,
        fontSize=10, fontName=fonts["body_bold"], spaceAfter=12,
    )
    h_cover_title_dark = ParagraphStyle(
        "CoverTitle", parent=base["Title"], textColor=WHITE, fontSize=36,
        leading=42,
        fontName=fonts["head_bold_italic"] if vibe.typography.headline_italic
        else fonts["head_bold"],
        alignment=TA_LEFT, spaceAfter=14,
    )
    h_cover_sub_dark = ParagraphStyle(
        "CoverSub", parent=base["Normal"], textColor=HexColor("#E7E5E4"),
        fontSize=12, leading=18, fontName=fonts["body_regular"],
    )
    h_cover_meta_dark = ParagraphStyle(
        "CoverMeta", parent=base["Normal"], textColor=HexColor("#A8A29E"),
        fontSize=9.5, leading=14, fontName=fonts["body_regular"], spaceBefore=18,
    )
    h_title = ParagraphStyle("VTitle", parent=base["Title"],
                             textColor=TEXT_PRIMARY, fontSize=24, leading=28,
                             fontName=fonts["head_bold_italic"] if vibe.typography.headline_italic else fonts["head_bold"],
                             alignment=TA_LEFT)
    h_sub = ParagraphStyle("Sub", parent=base["Italic"],
                           textColor=TEXT_MUTED, fontSize=10, spaceAfter=12,
                           fontName=fonts["body_italic"])
    h_section = ParagraphStyle("Section", parent=base["Heading1"],
                               textColor=TEXT_PRIMARY, fontSize=15, leading=18,
                               fontName=fonts["head_bold"],
                               spaceBefore=14, spaceAfter=2)
    h_section_rule = ParagraphStyle(
        "Rule", parent=base["Normal"], spaceBefore=0, spaceAfter=10,
        textColor=ACCENT, fontSize=9, fontName=fonts["body_bold"],
    )
    h_body = ParagraphStyle("Body", parent=base["BodyText"],
                            textColor=TEXT_SECONDARY, fontSize=10.5, leading=15,
                            fontName=fonts["body_regular"])
    h_muted = ParagraphStyle("Muted", parent=base["Italic"],
                             textColor=TEXT_MUTED, fontSize=9.5,
                             fontName=fonts["body_italic"])
    h_quote = ParagraphStyle("Quote", parent=base["Italic"],
                             textColor=TEXT_PRIMARY, fontSize=12, leading=18,
                             leftIndent=12, fontName=fonts["head_italic"])
    h_attrib = ParagraphStyle("Attrib", parent=base["Normal"],
                              textColor=TEXT_MUTED, fontSize=10, alignment=TA_RIGHT,
                              fontName=fonts["body_regular"])
    h_metric = ParagraphStyle("Metric", parent=base["Normal"],
                              textColor=TEXT_PRIMARY, fontSize=26, leading=28,
                              fontName=fonts["head_bold"], alignment=TA_CENTER)
    h_metric_label = ParagraphStyle("MetricLabel", parent=base["Normal"],
                                    textColor=TEXT_MUTED, fontSize=8,
                                    alignment=TA_CENTER,
                                    fontName=fonts["body_bold"], spaceAfter=4)

    # ── Document with custom page templates ──────────────────────────────
    is_quarterly = args.report_type == "quarterly"
    project_name = project.get("name", "")
    funder = project.get("funderName") or "Vision Empower"
    period = (
        f"{fmt_date(period_start)} – {fmt_date(period_end)}" if is_quarterly
        else f"{fmt_date(project.get('startDate'))} – {fmt_date(project.get('endDate'))}"
    )
    title_pdf = f"{project_name} — {'Quarterly' if is_quarterly else 'Project'} Report"

    page_w, page_h = A4
    margin_x = 2.0 * cm
    margin_top = 2.2 * cm
    margin_bottom = 2.4 * cm

    def _cover_page(canvas: pdfcanvas.Canvas, _doc):
        canvas.saveState()
        # Vibe-coloured band across the upper third — drives identity at first
        # glance even before any text.
        band_h = 8.8 * cm
        band_y = page_h - band_h - 0.5 * cm
        # Choose a dark band on light vibes and vice versa.
        band_fill = TEXT_PRIMARY if pal.bg.lower() not in ("#000000", "#0f1419") else SURFACE
        canvas.setFillColor(band_fill)
        canvas.rect(0, band_y, page_w, band_h, stroke=0, fill=1)
        # Top accent rule.
        canvas.setFillColor(ACCENT)
        canvas.rect(0, page_h - 0.35 * cm, page_w, 0.35 * cm, stroke=0, fill=1)
        # Eyebrow / brand mark on the band.
        canvas.setFillColor(ACCENT_SOFT)
        canvas.setFont(fonts["body_bold"], 10)
        canvas.drawString(margin_x, band_y + band_h - 1.4 * cm, "VISION EMPOWER")
        canvas.setFillColor(HexColor("#D6D3D1"))
        canvas.setFont(fonts["body_bold"], 8.5)
        canvas.drawString(
            margin_x, band_y + band_h - 1.8 * cm,
            "QUARTERLY FUNDER REPORT" if is_quarterly else "FULL PROJECT REPORT",
        )
        # Bottom-of-band metadata strip.
        canvas.setFillColor(HexColor("#A8A29E"))
        canvas.setFont(fonts["body_regular"], 9)
        canvas.drawString(
            margin_x, band_y + 0.8 * cm,
            f"Period: {period}    |    Funder: {funder}    |    Grant: {fmt_money(project.get('grantAmount'))}",
        )
        # Footer mark on cover.
        canvas.setFillColor(TEXT_MUTED)
        canvas.setFont(fonts["body_italic"], 8)
        canvas.drawCentredString(
            page_w / 2, 1.3 * cm,
            "Vision Empower Trust  ·  Enabling inclusive education for visually impaired children  ·  www.visionempower.in",
        )
        canvas.restoreState()

    def _content_page(canvas: pdfcanvas.Canvas, doc):
        canvas.saveState()
        # Top header band — thin coloured strip + project tag + funder.
        canvas.setFillColor(ACCENT)
        canvas.rect(0, page_h - 0.18 * cm, page_w, 0.18 * cm, stroke=0, fill=1)
        canvas.setFillColor(TEXT_MUTED)
        canvas.setFont(fonts["body_bold"], 8)
        canvas.drawString(margin_x, page_h - 1.0 * cm, project_name[:60].upper())
        canvas.drawRightString(page_w - margin_x, page_h - 1.0 * cm, funder.upper()[:50])
        # Footer: thin rule, page number, brand mark.
        canvas.setStrokeColor(BORDER)
        canvas.setLineWidth(0.4)
        canvas.line(margin_x, 1.6 * cm, page_w - margin_x, 1.6 * cm)
        canvas.setFillColor(TEXT_MUTED)
        canvas.setFont(fonts["body_regular"], 8)
        canvas.drawString(margin_x, 1.05 * cm, f"VEPIP  ·  {period}")
        canvas.drawRightString(
            page_w - margin_x, 1.05 * cm,
            f"Page {doc.page - 1}",  # cover doesn't count
        )
        canvas.restoreState()

    cover_frame = Frame(
        margin_x, margin_bottom,
        page_w - 2 * margin_x,
        page_h - margin_bottom - 11 * cm,  # leaves room for the cover band
        leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
        id="cover_frame",
    )
    content_frame = Frame(
        margin_x, margin_bottom,
        page_w - 2 * margin_x,
        page_h - margin_bottom - margin_top,
        leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
        id="content_frame",
    )

    doc = BaseDocTemplate(
        str(out), pagesize=A4,
        leftMargin=margin_x, rightMargin=margin_x,
        topMargin=margin_top, bottomMargin=margin_bottom,
        title=title_pdf,
        author="Vision Empower Trust",
    )
    doc.addPageTemplates([
        PageTemplate(id="cover", frames=[cover_frame], onPage=_cover_page),
        PageTemplate(id="content", frames=[content_frame], onPage=_content_page),
    ])

    story: list = []

    # ── Cover page content (sits below the painted band) ─────────────────
    cover_title = ParagraphStyle(
        "CoverProjectTitle", parent=base["Title"], textColor=TEXT_PRIMARY,
        fontSize=30, leading=34,
        fontName=fonts["head_bold_italic"] if vibe.typography.headline_italic
        else fonts["head_bold"],
        alignment=TA_LEFT, spaceBefore=4, spaceAfter=10,
    )
    cover_subhead = ParagraphStyle(
        "CoverSubhead", parent=base["Normal"], textColor=TEXT_SECONDARY,
        fontSize=12, leading=18, fontName=fonts["body_regular"],
    )
    cover_meta = ParagraphStyle(
        "CoverMetaLight", parent=base["Normal"], textColor=TEXT_MUTED,
        fontSize=10, leading=14, fontName=fonts["body_regular"], spaceBefore=14,
    )

    story.append(Paragraph(project_name, cover_title))
    if summary_text:
        story.append(Paragraph(summary_text[:380], cover_subhead))
    elif project.get("summary"):
        story.append(Paragraph(project["summary"][:380], cover_subhead))
    if project.get("states"):
        story.append(Paragraph("States: " + ", ".join(project["states"]), cover_meta))
    if project.get("startDate") or project.get("endDate"):
        story.append(Paragraph(
            f"Project term: {fmt_date(project.get('startDate'))} → {fmt_date(project.get('endDate'))}",
            cover_meta,
        ))
    # Decorative accent rule across the cover.
    story.append(Spacer(1, 0.6 * cm))
    rule_tbl = Table([[""]], colWidths=[2.5 * cm], rowHeights=[0.18 * cm])
    rule_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), ACCENT),
        ("BOX", (0, 0), (-1, -1), 0, WHITE),
    ]))
    story.append(rule_tbl)

    # ── Move to the content template ─────────────────────────────────────
    from reportlab.platypus.flowables import PageBreakIfNotEmpty
    from reportlab.platypus import NextPageTemplate
    story.append(NextPageTemplate("content"))
    story.append(PageBreak())

    # ── Executive summary (gated) ────────────────────────────────────────
    if has_overview:
        story.append(Paragraph("Executive Summary", h_section))
        story.append(Paragraph("OVERVIEW", h_section_rule))
        summary = summary_text
        if not summary and draft:
            lines = [l.strip() for l in draft.split("\n") if len(l.strip()) > 60]
            summary = " ".join(lines[:4])[:600]
        if summary:
            story.append(Paragraph(summary, h_body))

    # ── Impact tiles ─────────────────────────────────────────────────────
    teachers = sum(a.get("teachersReached") or 0 for a in activities)
    students = sum(a.get("studentsReached") or 0 for a in activities)
    schools = sum(a.get("schoolsReached") or 0 for a in activities)
    if teachers or students or schools:
        story.append(Paragraph("Impact at a Glance", h_section))
        story.append(Paragraph("THIS PERIOD", h_section_rule))
        cell_w = (page_w - 2 * margin_x) / 3
        impact_data = [
            [Paragraph("TEACHERS", h_metric_label),
             Paragraph("STUDENTS", h_metric_label),
             Paragraph("SCHOOLS", h_metric_label)],
            [Paragraph(f"{teachers:,}", h_metric),
             Paragraph(f"{students:,}", h_metric),
             Paragraph(f"{schools:,}", h_metric)],
            [Paragraph("reached this period", h_metric_label),
             Paragraph("reached this period", h_metric_label),
             Paragraph("reached this period", h_metric_label)],
        ]
        impact_tbl = Table(
            impact_data,
            colWidths=[cell_w - 6, cell_w - 6, cell_w - 6],
            rowHeights=[0.6 * cm, 1.4 * cm, 0.55 * cm],
        )
        impact_tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), SURFACE_ALT),
            ("LINEABOVE", (0, 0), (-1, 0), 2, ACCENT),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("BOX", (0, 0), (-1, -1), 0, WHITE),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
        story.append(impact_tbl)
        story.append(Spacer(1, 0.2 * cm))

    # ── Deliverables: table + bar chart ──────────────────────────────────
    if has_deliverables:
        story.append(Paragraph("Deliverable Progress", h_section))
        story.append(Paragraph("WHAT THE GRANT COMMITTED TO", h_section_rule))
        rows = [["Deliverable", "Target", "Achieved", "% Complete"]]
        for d in deliverables:
            unit = d.get("unit") or ""
            target = d.get("target") or 0
            achieved = d.get("achieved") or 0
            pct = round(achieved / target * 100) if target else (100 if d.get("status") == "completed" else 0)
            rows.append([
                d.get("title", ""),
                f"{target}{(' ' + unit) if (target and unit) else ''}" if target else "—",
                f"{achieved}{(' ' + unit) if (achieved and unit) else ''}" if achieved else "—",
                f"{pct}%",
            ])
        d_tbl = Table(rows, colWidths=[8 * cm, 2.6 * cm, 2.6 * cm, 2.5 * cm], hAlign="LEFT")
        d_tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), TEXT_PRIMARY),
            ("TEXTCOLOR", (0, 0), (-1, 0), ACCENT_SOFT),
            ("FONTNAME", (0, 0), (-1, 0), fonts["body_bold"]),
            ("FONTNAME", (0, 1), (-1, -1), fonts["body_regular"]),
            ("FONTSIZE", (0, 0), (-1, -1), 9.5),
            ("TEXTCOLOR", (0, 1), (-1, -1), TEXT_SECONDARY),
            ("ALIGN", (1, 0), (-1, -1), "CENTER"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, SURFACE_ALT]),
            ("LINEBELOW", (0, 0), (-1, 0), 0, WHITE),
            ("LINEBELOW", (0, 1), (-1, -1), 0.25, BORDER),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
        story.append(d_tbl)
        # Visual bars beneath the table.
        bars = make_deliverable_bars(
            deliverables,
            accent_hex=pal.as_hex("accent"),
            track_hex=pal.as_hex("surface_alt"),
            label_hex=pal.as_hex("text_secondary"),
            body_font=fonts["body_regular"],
        )
        if bars is not None:
            story.append(Spacer(1, 0.3 * cm))
            story.append(bars)

    # ── Activities ───────────────────────────────────────────────────────
    if has_activities:
        story.append(Paragraph("Field Activities", h_section))
        story.append(Paragraph("ON-GROUND WORK", h_section_rule))
        for act in activities[:15]:
            d = fmt_date(act.get("activityDate")) if act.get("activityDate") else ""
            head = f"<b>{d} — {act.get('title', '')}</b>" if d else f"<b>{act.get('title', '')}</b>"
            reaches = []
            if act.get("teachersReached"): reaches.append(f"{act['teachersReached']} teachers")
            if act.get("studentsReached"): reaches.append(f"{act['studentsReached']} students")
            if act.get("schoolsReached"):  reaches.append(f"{act['schoolsReached']} schools")
            tail_bits = []
            if act.get("state"): tail_bits.append(act["state"])
            if reaches: tail_bits.append(", ".join(reaches))
            if act.get("notes"): tail_bits.append(act["notes"])
            tail = " · ".join(tail_bits)
            story.append(Paragraph(f"{head}. {tail}", h_body))

    # ── Financials: donut + table ────────────────────────────────────────
    if has_financials:
        story.append(Paragraph("Budget & Utilisation", h_section))
        story.append(Paragraph("WHERE THE GRANT IS GOING", h_section_rule))
        # Donut on the left, summary numbers on the right.
        donut = make_donut(
            approved=approved, spent=spent,
            accent_hex=pal.as_hex("accent"),
            soft_hex=pal.as_hex("surface_alt"),
            muted_hex=pal.as_hex("text_muted"),
            label_hex=pal.as_hex("text_primary"),
            body_font=fonts["body_regular"],
            label_font=fonts["head_bold"],
        )
        util_pct = round(spent / approved * 100) if approved else 0
        summary_para = Paragraph(
            f"<font name='{fonts['body_bold']}' color='#{pal.as_hex('text_primary')}'>Approved</font><br/>"
            f"<font size='14' name='{fonts['head_bold']}' color='#{pal.as_hex('text_primary')}'>{fmt_money(approved)}</font>"
            f"<br/><br/>"
            f"<font name='{fonts['body_bold']}' color='#{pal.as_hex('text_primary')}'>Spent</font><br/>"
            f"<font size='14' name='{fonts['head_bold']}' color='#{pal.as_hex('text_primary')}'>{fmt_money(spent)}</font>"
            f"<br/><br/>"
            f"<font name='{fonts['body_bold']}' color='#{pal.as_hex('text_primary')}'>Utilisation</font><br/>"
            f"<font size='14' name='{fonts['head_bold']}' color='#{pal.as_hex('accent')}'>{util_pct}%</font>",
            h_body,
        )
        finance_tbl = Table(
            [[donut, summary_para]],
            colWidths=[5 * cm, page_w - 2 * margin_x - 5 * cm],
        )
        finance_tbl.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (1, 0), (1, 0), 14),
            ("BOX", (0, 0), (-1, -1), 0, WHITE),
        ]))
        story.append(finance_tbl)
        story.append(Spacer(1, 0.3 * cm))
        budgets = project.get("budgets") or []
        if budgets:
            rows = [["Category", "Approved", "Spent", "Utilisation"]]
            for b in budgets:
                ap = b.get("approvedAmount") or 0
                sp = b.get("spentAmount") or 0
                rows.append([
                    b.get("name", ""),
                    fmt_money(ap), fmt_money(sp),
                    f"{round(sp / ap * 100) if ap else 0}%",
                ])
            b_tbl = Table(rows, colWidths=[8 * cm, 2.6 * cm, 2.6 * cm, 2.5 * cm], hAlign="LEFT")
            b_tbl.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), TEXT_PRIMARY),
                ("TEXTCOLOR", (0, 0), (-1, 0), ACCENT_SOFT),
                ("FONTNAME", (0, 0), (-1, 0), fonts["body_bold"]),
                ("FONTNAME", (0, 1), (-1, -1), fonts["body_regular"]),
                ("FONTSIZE", (0, 0), (-1, -1), 9.5),
                ("TEXTCOLOR", (0, 1), (-1, -1), TEXT_SECONDARY),
                ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, SURFACE_ALT]),
                ("LINEBELOW", (0, 1), (-1, -1), 0.25, BORDER),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]))
            story.append(b_tbl)

    # ── Stories (gated; threshold matches vibe) ──────────────────────────
    if has_stories:
        strong = [t for t in testimonials
                  if (t.get("content") or "").strip()
                  and len((t.get("content") or "").strip()) >= quote_threshold]
        if strong:
            story.append(Paragraph("Stories from the Field", h_section))
            story.append(Paragraph("IN THEIR OWN WORDS", h_section_rule))
            for t in strong[:3]:
                story.append(Paragraph(f"&ldquo;{t.get('content', '')}&rdquo;", h_quote))
                role = t.get("role") or ""
                story.append(Paragraph(
                    f"— {t.get('author', '')}{', ' + role if role else ''}", h_attrib))
                story.append(Spacer(1, 0.25 * cm))

    # ── Achievements (if structured narrative provided) ──────────────────
    if narrative and (narrative.get("achievements") or "").strip():
        story.append(Paragraph("Key Achievements", h_section))
        story.append(Paragraph("WHAT WE'RE PROUD OF", h_section_rule))
        story.append(Paragraph(narrative["achievements"].strip()[:1200], h_body))
        story.append(Spacer(1, 0.3 * cm))

    # ── Way forward (gated; no filler) ──────────────────────────────────
    if has_way_forward:
        challenges = (narrative.get("challenges") or "").strip() if narrative else ""
        next_steps = (narrative.get("way_forward") or "").strip() if narrative else ""

        if (not challenges or not next_steps) and draft:
            import re as _re

            def extract(text, start_re, stop_re):
                parts = _re.split(start_re, text, flags=_re.IGNORECASE, maxsplit=1)
                if len(parts) < 2: return ""
                content = parts[1]
                m = _re.search(stop_re, content, flags=_re.IGNORECASE)
                if m: content = content[: m.start()]
                return _re.sub(r"\s+", " ", _re.sub(r"^[\s:–\-]+", "", content)).strip()[:600]

            if not challenges:
                challenges = extract(draft, r"challenges?|barriers?|constraints?",
                                     r"next\s+steps?|way\s+forward|financial")
            if not next_steps:
                next_steps = extract(draft, r"next\s+steps?|way\s+forward|upcoming\s+activities",
                                     r"evidence|conclusion|thank|challenges?")

        story.append(Paragraph("Challenges & Next Steps", h_section))
        story.append(Paragraph("LOOKING AHEAD", h_section_rule))
        if challenges:
            story.append(Paragraph(f"<b>Challenges.</b> {challenges}", h_body))
        if next_steps:
            story.append(Paragraph(f"<b>Next Steps.</b> {next_steps}", h_body))

    doc.build(story)
    size = out.stat().st_size
    if size < 2500:
        raise RuntimeError(f"Generated PDF is suspiciously small ({size} bytes)")
    print(f"WROTE: {out} ({size:,} bytes, vibe={vibe.key})")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        traceback.print_exc()
        sys.exit(1)
