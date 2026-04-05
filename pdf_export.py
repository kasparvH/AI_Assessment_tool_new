import io
import textwrap
from datetime import date
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    BaseDocTemplate, Frame, PageTemplate, Paragraph,
    Spacer, Table, TableStyle, Image as RLImage, HRFlowable,
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.pdfbase import pdfmetrics

NAVY = colors.HexColor("#1B3B6F")
ORANGE = colors.HexColor("#F7941D")
LIGHT_GREY = colors.HexColor("#F8F9FA")
WHITE = colors.white
RED = colors.HexColor("#dc3545")
YELLOW = colors.HexColor("#ffc107")
GREEN = colors.HexColor("#28a745")

PAGE_W, PAGE_H = A4
MARGIN = 20 * mm


def _score_color(pct: float):
    if pct < 35:
        return RED
    if pct < 60:
        return ORANGE
    if pct < 80:
        return YELLOW
    return GREEN


def create_radar_chart(dim_summary: pd.DataFrame, output_path: str):
    dimensions = dim_summary["dimension"].tolist()
    scores = dim_summary["score_0_5"].tolist()
    N = len(dimensions)

    angles = np.linspace(0, 2 * np.pi, N, endpoint=False).tolist()
    scores_plot = scores + [scores[0]]
    angles_plot = angles + angles[:1]

    fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(polar=True))
    ax.plot(angles_plot, scores_plot, "o-", linewidth=2.5, color="#F7941D")
    ax.fill(angles_plot, scores_plot, alpha=0.2, color="#1B3B6F")
    ax.set_ylim(0, 5)
    ax.set_yticks([1, 2, 3, 4, 5])
    ax.set_yticklabels(["1", "2", "3", "4", "5"], fontsize=9, color="gray")
    ax.set_xticks(angles)

    labels = [textwrap.fill(d, 14) for d in dimensions]
    ax.set_xticklabels(labels, fontsize=9, color="#1B3B6F", fontweight="bold")
    ax.set_facecolor("#F8F9FA")
    ax.grid(color="#cccccc", linestyle="--", linewidth=0.7)
    ax.set_title("AI Maturity by Dimension", fontsize=14,
                 color="#1B3B6F", fontweight="bold", pad=20)
    plt.tight_layout()
    plt.savefig(output_path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close()


def _footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(colors.grey)
    text = f"Confidential — Straightable Innovation & Strategy — {date.today().isoformat()}"
    canvas.drawCentredString(PAGE_W / 2, 10 * mm, text)
    canvas.drawRightString(PAGE_W - MARGIN, 10 * mm, f"Page {doc.page}")
    canvas.restoreState()


def export_pdf(
    report_text: str,
    dim_summary: pd.DataFrame,
    overall: dict,
    compliance: dict,
    org_name: str,
    respondent_name: str,
    logo_path: str,
    output_path: str,
):
    styles = getSampleStyleSheet()
    h1 = ParagraphStyle("H1", parent=styles["Heading1"], textColor=NAVY,
                         fontSize=20, spaceAfter=6)
    h2 = ParagraphStyle("H2", parent=styles["Heading2"], textColor=NAVY,
                         fontSize=13, spaceAfter=4, spaceBefore=10)
    body = ParagraphStyle("Body", parent=styles["Normal"], fontSize=9,
                           leading=14, spaceAfter=4)
    small = ParagraphStyle("Small", parent=styles["Normal"], fontSize=8,
                            textColor=colors.grey)

    doc = BaseDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=MARGIN, bottomMargin=18 * mm,
    )
    frame = Frame(MARGIN, 18 * mm, PAGE_W - 2 * MARGIN, PAGE_H - MARGIN - 18 * mm)
    template = PageTemplate(id="main", frames=[frame], onPage=_footer)
    doc.addPageTemplates([template])

    story = []

    # Cover
    if Path(logo_path).exists():
        story.append(RLImage(logo_path, width=60 * mm, height=20 * mm))
        story.append(Spacer(1, 8 * mm))

    story.append(Paragraph(f"AI Readiness Assessment Report", h1))
    story.append(Paragraph(f"<b>{org_name}</b>", styles["Heading2"]))
    story.append(Spacer(1, 4 * mm))
    story.append(Paragraph(f"Respondent: {respondent_name}", body))
    story.append(Paragraph(f"Date: {date.today().strftime('%d %B %Y')}", body))
    story.append(Spacer(1, 4 * mm))

    # Maturity badge
    tier = overall["maturity_tier"]
    score = overall["overall_score_0_5"]
    pct = overall["overall_pct"]
    badge_data = [[
        Paragraph(f"<b>Overall Maturity: {tier}</b>", ParagraphStyle(
            "badge", fontSize=13, textColor=WHITE, alignment=TA_CENTER)),
        Paragraph(f"<b>{score:.1f} / 5.0 &nbsp; ({pct:.0f}%)</b>", ParagraphStyle(
            "badge2", fontSize=11, textColor=WHITE, alignment=TA_CENTER)),
    ]]
    badge_table = Table(badge_data, colWidths=[(PAGE_W - 2 * MARGIN) * 0.6,
                                               (PAGE_W - 2 * MARGIN) * 0.4])
    badge_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), NAVY),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("ROUNDEDCORNERS", [6]),
    ]))
    story.append(badge_table)
    story.append(Spacer(1, 6 * mm))

    # Radar chart
    radar_path = str(Path(output_path).parent / "radar_tmp.png")
    create_radar_chart(dim_summary, radar_path)
    story.append(RLImage(radar_path, width=120 * mm, height=120 * mm))
    story.append(Spacer(1, 4 * mm))

    # Dimension scores table
    story.append(Paragraph("Dimension Scores", h2))
    dim_rows = [["Dimension", "Score", "%", "Maturity"]]
    for _, row in dim_summary.sort_values("score_pct", ascending=False).iterrows():
        dim_rows.append([
            row["dimension"],
            f"{row['score_0_5']:.1f} / 5.0",
            f"{row['score_pct']:.0f}%",
            _tier_label(row["score_pct"]),
        ])
    dim_table = Table(dim_rows, colWidths=[90 * mm, 30 * mm, 20 * mm, 30 * mm])
    dim_style = [
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, LIGHT_GREY]),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.lightgrey),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]
    for i, (_, row) in enumerate(dim_summary.sort_values("score_pct", ascending=False).iterrows(), 1):
        c = _score_color(row["score_pct"])
        dim_style.append(("TEXTCOLOR", (3, i), (3, i), c))
        dim_style.append(("FONTNAME", (3, i), (3, i), "Helvetica-Bold"))
    dim_table.setStyle(TableStyle(dim_style))
    story.append(dim_table)
    story.append(Spacer(1, 6 * mm))

    # Report text
    story.append(HRFlowable(width="100%", color=ORANGE, thickness=1.5))
    story.append(Spacer(1, 3 * mm))
    for line in report_text.split("\n"):
        line = line.rstrip()
        if line.startswith("## "):
            story.append(Spacer(1, 3 * mm))
            story.append(Paragraph(line[3:], h2))
        elif line.startswith("**") and line.endswith("**"):
            story.append(Paragraph(f"<b>{line[2:-2]}</b>", body))
        elif line.strip() == "":
            story.append(Spacer(1, 2 * mm))
        else:
            story.append(Paragraph(line, body))

    # Compliance table
    story.append(Spacer(1, 6 * mm))
    story.append(Paragraph("Compliance & Framework Coverage", h2))
    comp_rows = [["Framework", "Questions covered", "Avg. Score"]]
    for fw, data in compliance.items():
        comp_rows.append([
            fw.replace("_", " "),
            f"{data['covered']} / {data['total']}",
            f"{data['avg_score']:.1f} / 5.0",
        ])
    comp_table = Table(comp_rows, colWidths=[70 * mm, 60 * mm, 40 * mm])
    comp_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, LIGHT_GREY]),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.lightgrey),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(comp_table)

    doc.build(story)

    # Clean up temp radar
    try:
        Path(radar_path).unlink()
    except Exception:
        pass


def _tier_label(pct: float) -> str:
    if pct < 35:
        return "Emerging"
    if pct < 60:
        return "Developing"
    if pct < 80:
        return "Accelerating"
    return "Leading"
