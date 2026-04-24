from __future__ import annotations

import io
from typing import Iterable

import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from .data_utils import build_report_summary, build_rule_based_narrative, build_people_summary


def _set_title(slide, title: str, subtitle: str | None = None):
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(12.3), Inches(0.5))
    p = title_box.text_frame.paragraphs[0]
    p.text = title
    p.font.size = Pt(24)
    p.font.bold = True
    if subtitle:
        sub = slide.shapes.add_textbox(Inches(0.52), Inches(0.86), Inches(12), Inches(0.35))
        p2 = sub.text_frame.paragraphs[0]
        p2.text = subtitle
        p2.font.size = Pt(11)


def _add_metric(slide, x, y, label, value):
    box = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(2.3), Inches(0.9))
    box.text_frame.clear()
    p = box.text_frame.paragraphs[0]
    p.text = str(value)
    p.font.size = Pt(24)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER
    p2 = box.text_frame.add_paragraph()
    p2.text = label
    p2.font.size = Pt(10)
    p2.alignment = PP_ALIGN.CENTER


def _add_bullets(slide, x, y, w, h, lines: Iterable[str], font_size=14):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.font.size = Pt(font_size)
        p.level = 0


def _add_bar_chart(slide, x, y, w, h, title: str, data: dict):
    chart_data = CategoryChartData()
    items = list(data.items())[:8]
    if not items:
        items = [("No data", 0)]
    chart_data.categories = [str(k)[:30] for k, _ in items]
    chart_data.add_series(title, [int(v) for _, v in items])
    chart = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(x), Inches(y), Inches(w), Inches(h), chart_data).chart
    chart.has_legend = False
    chart.chart_title.has_text_frame = True
    chart.chart_title.text_frame.text = title


def _add_table(slide, x, y, w, h, df: pd.DataFrame, max_rows=8):
    view = df.head(max_rows).copy()
    if view.empty:
        view = pd.DataFrame({"Message": ["No records found for selected filters."]})
    rows, cols = view.shape[0] + 1, view.shape[1]
    table = slide.shapes.add_table(rows, cols, Inches(x), Inches(y), Inches(w), Inches(h)).table
    for j, col in enumerate(view.columns):
        table.cell(0, j).text = str(col)
    for i, (_, row) in enumerate(view.iterrows(), start=1):
        for j, col in enumerate(view.columns):
            val = row[col]
            if pd.isna(val):
                val = ""
            table.cell(i, j).text = str(val)[:120]
    for r in range(rows):
        for c in range(cols):
            for p in table.cell(r, c).text_frame.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(7 if r else 8)
                    run.font.bold = r == 0


def generate_achievement_ppt(df: pd.DataFrame, report_title: str, period_label: str) -> bytes:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    summary = build_report_summary(df)
    narrative = build_rule_based_narrative(summary, period_label)

    # Slide 1
    slide = prs.slides.add_slide(blank)
    _set_title(slide, report_title)
    _add_bullets(slide, 0.7, 1.7, 10.8, 1.2, [f"Reporting Period: {period_label}", "Generated from Achievement Tracker", "Template: DeptFlow Achievement Reporter MVP"], font_size=20)
    _add_metric(slide, 0.8, 4.1, "Achievements", summary["total_achievements"])
    _add_metric(slide, 3.4, 4.1, "Studies / Workstreams", summary["unique_studies"])
    _add_metric(slide, 6.0, 4.1, "Contributors", summary["unique_contributors"])
    _add_metric(slide, 8.6, 4.1, "CF-related Count", summary["cf_total"])

    # Slide 2
    slide = prs.slides.add_slide(blank)
    _set_title(slide, "Executive Summary", period_label)
    _add_bullets(slide, 0.7, 1.25, 11.8, 1.8, [narrative], font_size=16)
    bullets = [
        f"Total achievements: {summary['total_achievements']}",
        f"Role breakdown: {', '.join([f'{k}: {v}' for k, v in summary['role_breakdown'].items()]) or 'N/A'}",
        f"Top activity types: {', '.join(list(summary['activity_breakdown'].keys())[:5]) or 'N/A'}",
        "Recommended control: LM confirmation before final leadership distribution.",
    ]
    _add_bullets(slide, 0.9, 3.3, 11, 2.6, bullets, font_size=15)

    # Slide 3
    slide = prs.slides.add_slide(blank)
    _set_title(slide, "Achievement Overview Dashboard", period_label)
    _add_metric(slide, 0.65, 1.15, "Achievements", summary["total_achievements"])
    _add_metric(slide, 3.2, 1.15, "Studies / Workstreams", summary["unique_studies"])
    _add_metric(slide, 5.75, 1.15, "Contributors", summary["unique_contributors"])
    _add_metric(slide, 8.3, 1.15, "CF Total", summary["cf_total"])
    _add_bar_chart(slide, 0.7, 2.65, 5.7, 3.8, "Activity Type Breakdown", summary["activity_breakdown"])
    _add_bar_chart(slide, 6.8, 2.65, 5.7, 3.8, "Slide Category Breakdown", summary["category_breakdown"])

    # Slide 4
    slide = prs.slides.add_slide(blank)
    _set_title(slide, "Project / Study Highlights", period_label)
    study_cols = ["study_id", "ta_area", "activity_type", "submitted_by_name", "comments"]
    study_df = df[study_cols].copy() if not df.empty else pd.DataFrame()
    study_df.columns = ["Study ID", "TA", "Activity", "Owner", "Comments"] if not study_df.empty else []
    _add_table(slide, 0.55, 1.15, 12.2, 5.8, study_df, max_rows=9)

    # Slide 5
    slide = prs.slides.add_slide(blank)
    _set_title(slide, "People Contribution View", period_label)
    people_df = build_people_summary(df)
    if not people_df.empty:
        people_df = people_df[["person_name", "achievement_count", "main_activity"]]
        people_df.columns = ["Person", "Achievement Count", "Main Activity"]
    _add_table(slide, 0.8, 1.2, 11.7, 5.6, people_df, max_rows=12)

    # Slide 6
    slide = prs.slides.add_slide(blank)
    _set_title(slide, "Appendix: Achievement Detail", period_label)
    appendix_cols = ["delivery_date", "role", "study_id", "activity_type", "submitted_by_name", "team_members_display", "comments"]
    appendix = df[appendix_cols].copy() if not df.empty else pd.DataFrame()
    appendix.columns = ["Delivery", "Role", "Study", "Activity", "Owner", "Team", "Comments"] if not appendix.empty else []
    _add_table(slide, 0.35, 1.05, 12.65, 6.0, appendix, max_rows=10)

    stream = io.BytesIO()
    prs.save(stream)
    return stream.getvalue()
