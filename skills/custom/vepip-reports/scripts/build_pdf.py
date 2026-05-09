#!/usr/bin/env python3
"""Vision Empower premium funder report — branded PDF builder (reportlab).

Usage:
    python build_pdf.py --data /path/to/data.json --output /path/to/out.pdf
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

from reportlab.lib import colors  # noqa: E402
from reportlab.lib.colors import HexColor  # noqa: E402
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT  # noqa: E402
from reportlab.lib.pagesizes import A4  # noqa: E402
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet  # noqa: E402
from reportlab.lib.units import cm  # noqa: E402
from reportlab.platypus import (  # noqa: E402
    Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)

DBROWN = HexColor("#2A1508")
GOLD   = HexColor("#C49A32")
CARDF  = HexColor("#F5EDD0")
CREAM  = HexColor("#F7F3EE")
BODY   = HexColor("#3D2010")
MUTED  = HexColor("#9B7B5A")
WHITE  = HexColor("#FFFFFF")


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
    period_start = args.period_start or data.get("periodStart") or project.get("startDate") or ""
    period_end   = args.period_end   or data.get("periodEnd")   or project.get("endDate")   or ""
    draft        = args.draft or data.get("draft") or ""

    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)

    base = getSampleStyleSheet()
    h_eyebrow = ParagraphStyle("Eyebrow", parent=base["Normal"],
                               textColor=GOLD, fontSize=8.5, spaceAfter=2,
                               fontName="Helvetica-Bold")
    h_title = ParagraphStyle("VTitle", parent=base["Title"],
                             textColor=DBROWN, fontSize=24, leading=28,
                             fontName="Times-Bold", alignment=TA_LEFT)
    h_sub = ParagraphStyle("Sub", parent=base["Italic"],
                           textColor=MUTED, fontSize=10, spaceAfter=12)
    h_section = ParagraphStyle("Section", parent=base["Heading1"],
                               textColor=DBROWN, fontSize=15, leading=18,
                               fontName="Times-Bold", spaceBefore=14, spaceAfter=6)
    h_body = ParagraphStyle("Body", parent=base["BodyText"],
                            textColor=BODY, fontSize=10.5, leading=15)
    h_muted = ParagraphStyle("Muted", parent=base["Italic"],
                             textColor=MUTED, fontSize=9.5)
    h_quote = ParagraphStyle("Quote", parent=base["Italic"],
                             textColor=DBROWN, fontSize=12, leading=18,
                             leftIndent=12, fontName="Times-Italic")
    h_attrib = ParagraphStyle("Attrib", parent=base["Normal"],
                              textColor=MUTED, fontSize=10, alignment=TA_RIGHT)
    h_metric = ParagraphStyle("Metric", parent=base["Normal"],
                              textColor=DBROWN, fontSize=24, leading=26,
                              fontName="Times-Bold", alignment=TA_CENTER)
    h_metric_label = ParagraphStyle("MetricLabel", parent=base["Normal"],
                                    textColor=MUTED, fontSize=8, alignment=TA_CENTER,
                                    fontName="Helvetica-Bold", spaceAfter=4)

    doc = SimpleDocTemplate(
        str(out), pagesize=A4,
        leftMargin=2.0 * cm, rightMargin=2.0 * cm,
        topMargin=1.8 * cm, bottomMargin=1.8 * cm,
        title=f"{project.get('name', 'VEPIP Report')} — Funder Report",
    )

    story: list = []

    is_quarterly = args.report_type == "quarterly"
    period = (
        f"{fmt_date(period_start)} – {fmt_date(period_end)}" if is_quarterly
        else f"{fmt_date(project.get('startDate'))} – {fmt_date(project.get('endDate'))}"
    )

    # Cover block
    story.append(Paragraph(
        "QUARTERLY FUNDER REPORT" if is_quarterly else "FULL PROJECT REPORT", h_eyebrow))
    story.append(Paragraph("VISION EMPOWER", ParagraphStyle("V", parent=base["Normal"],
                                                            fontSize=8.5, textColor=GOLD, fontName="Helvetica-Bold",
                                                            spaceAfter=8)))
    story.append(Paragraph(project.get("name", ""), h_title))
    story.append(Paragraph(
        f"Funder: {project.get('funderName', '')} | Period: {period} | Grant: {fmt_money(project.get('grantAmount'))}",
        h_sub,
    ))
    if project.get("states"):
        story.append(Paragraph(f"States: {', '.join(project['states'])}", h_muted))

    # Executive summary
    story.append(Spacer(1, 0.3 * cm))
    story.append(Paragraph("Executive Summary", h_section))
    summary = project.get("summary") or ""
    if not summary and draft:
        lines = [l.strip() for l in draft.split("\n") if len(l.strip()) > 60]
        summary = " ".join(lines[:4])[:600]
    if not summary:
        summary = (
            f"This report summarises progress made under the {project.get('name', '')} "
            f"project funded by {project.get('funderName', '')}."
        )
    story.append(Paragraph(summary, h_body))

    # Impact tiles (3-up)
    activities = project.get("activities") or []
    teachers = sum(a.get("teachersReached") or 0 for a in activities)
    students = sum(a.get("studentsReached") or 0 for a in activities)
    schools  = sum(a.get("schoolsReached") or 0 for a in activities)

    story.append(Paragraph("Impact at a Glance", h_section))
    impact_data = [
        [Paragraph(f"{teachers:,}", h_metric),
         Paragraph(f"{students:,}", h_metric),
         Paragraph(f"{schools:,}", h_metric)],
        [Paragraph("TEACHERS", h_metric_label),
         Paragraph("STUDENTS", h_metric_label),
         Paragraph("SCHOOLS", h_metric_label)],
    ]
    impact_tbl = Table(impact_data, colWidths=[5.2 * cm] * 3, rowHeights=[1.4 * cm, 0.5 * cm])
    impact_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), CARDF),
        ("LINEABOVE",  (0, 0), (-1, 0), 2, GOLD),
        ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
        ("BOX",        (0, 0), (-1, -1), 0, WHITE),
    ]))
    story.append(impact_tbl)

    # Deliverables table
    story.append(Paragraph("Deliverable Progress", h_section))
    deliverables = project.get("deliverables") or []
    if deliverables:
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
        d_tbl = Table(rows, colWidths=[7.5 * cm, 3 * cm, 3 * cm, 2.5 * cm], hAlign="LEFT")
        d_tbl.setStyle(TableStyle([
            ("BACKGROUND",  (0, 0), (-1, 0), DBROWN),
            ("TEXTCOLOR",   (0, 0), (-1, 0), HexColor("#EDD98A")),
            ("FONTNAME",    (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE",    (0, 0), (-1, -1), 9.5),
            ("ALIGN",       (1, 0), (-1, -1), "CENTER"),
            ("ROWBACKGROUNDS",(0, 1), (-1, -1), [WHITE, CREAM]),
            ("GRID",        (0, 0), (-1, -1), 0.25, HexColor("#E8DDD0")),
            ("VALIGN",      (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING",(0, 0), (-1, -1), 6),
            ("TOPPADDING",  (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING",(0, 0), (-1, -1), 5),
        ]))
        story.append(d_tbl)
    else:
        story.append(Paragraph("No deliverables recorded.", h_muted))

    # Activities
    story.append(Paragraph("Field Activities", h_section))
    if activities:
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
    else:
        story.append(Paragraph("No activities recorded.", h_muted))

    # Financials
    story.append(Paragraph("Budget & Utilisation", h_section))
    approved = project.get("approvedBudget") or 0
    spent    = project.get("spentBudget") or 0
    util_pct = round(spent / approved * 100) if approved else 0
    story.append(Paragraph(
        f"<b>Approved:</b> {fmt_money(approved)} &nbsp;&nbsp; "
        f"<b>Spent:</b> {fmt_money(spent)} &nbsp;&nbsp; "
        f"<b>Utilisation:</b> {util_pct}%",
        h_body,
    ))
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
        b_tbl = Table(rows, colWidths=[7 * cm, 3 * cm, 3 * cm, 3 * cm], hAlign="LEFT")
        b_tbl.setStyle(TableStyle([
            ("BACKGROUND",   (0, 0), (-1, 0), DBROWN),
            ("TEXTCOLOR",    (0, 0), (-1, 0), HexColor("#EDD98A")),
            ("FONTNAME",     (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE",     (0, 0), (-1, -1), 9.5),
            ("ALIGN",        (1, 0), (-1, -1), "RIGHT"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, CREAM]),
            ("GRID",         (0, 0), (-1, -1), 0.25, HexColor("#E8DDD0")),
            ("LEFTPADDING",  (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ]))
        story.append(b_tbl)

    # Stories
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
        story.append(Paragraph("Stories from the Field", h_section))
        for t in testimonials[:3]:
            story.append(Paragraph(f"&ldquo;{t.get('content', '')}&rdquo;", h_quote))
            role = t.get("role") or ""
            story.append(Paragraph(f"— {t.get('author', '')}{', ' + role if role else ''}", h_attrib))
            story.append(Spacer(1, 0.25 * cm))

    # Way forward
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

    story.append(Paragraph("Challenges & Next Steps", h_section))
    story.append(Paragraph(f"<b>Challenges.</b> {challenges or 'To be documented in consultation with the field team.'}", h_body))
    story.append(Paragraph(f"<b>Next Steps.</b> {next_steps or 'To be confirmed based on project progress and funder guidance.'}", h_body))

    story.append(Spacer(1, 0.5 * cm))
    story.append(Paragraph(
        "<i>Vision Empower Trust  ·  Enabling inclusive education for visually impaired children  ·  www.visionempower.in</i>",
        h_muted,
    ))

    doc.build(story)
    size = out.stat().st_size
    if size < 4000:
        raise RuntimeError(f"Generated PDF is suspiciously small ({size} bytes)")
    print(f"WROTE: {out} ({size:,} bytes)")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        traceback.print_exc()
        sys.exit(1)
