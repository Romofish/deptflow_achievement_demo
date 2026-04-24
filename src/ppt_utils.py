from __future__ import annotations

import io
from typing import Any, Iterable

import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from .ai_utils import build_rule_based_narrative
from .data_utils import PRESENTATION_COLUMNS, build_people_summary
from .metrics_utils import (
    build_quality_complexity_highlights,
    build_study_highlights,
    calculate_report_metrics,
)


TEMPLATE_VERSION = "deptflow-fixed-v2"

FIXED_TEMPLATE_SLIDES = [
    {"Slide": 1, "Name": "Title Page", "Content": "Report title, period, template metadata, and KPI cards"},
    {"Slide": 2, "Name": "Executive Summary", "Content": "Editable narrative plus calculated metric callouts"},
    {"Slide": 3, "Name": "Achievement Overview Dashboard", "Content": "Core KPIs, category mix, and impact mix"},
    {"Slide": 4, "Name": "Activity Type Breakdown", "Content": "Top activity types from filtered data"},
    {"Slide": 5, "Name": "Project / Study Highlights", "Content": "Study-level summary table"},
    {"Slide": 6, "Name": "People Contribution View", "Content": "Contributor summary from filtered data"},
    {"Slide": 7, "Name": "Quality / Complexity Highlights", "Content": "High-impact, quality, and CF indicators"},
    {"Slide": 8, "Name": "Appendix Detail Table", "Content": "Filtered achievement details"},
]

BRAND_DARK = RGBColor(22, 36, 54)
BRAND_BLUE = RGBColor(35, 102, 168)
BRAND_TEAL = RGBColor(21, 139, 132)
BRAND_ORANGE = RGBColor(222, 118, 44)
BRAND_LIGHT = RGBColor(244, 247, 250)
BRAND_MUTED = RGBColor(92, 105, 121)
WHITE = RGBColor(255, 255, 255)


def _set_fill(shape, color: RGBColor) -> None:
    shape.fill.solid()
    shape.fill.fore_color.rgb = color


def _set_line(shape, color: RGBColor, width: float = 0.75) -> None:
    shape.line.color.rgb = color
    shape.line.width = Pt(width)


def _set_paragraph_font(paragraph, size: int, color: RGBColor = BRAND_DARK, bold: bool = False) -> None:
    paragraph.font.size = Pt(size)
    paragraph.font.color.rgb = color
    paragraph.font.bold = bold


def _add_footer(slide, generated_at: str | None = None, ai_provider: str | None = None) -> None:
    footer = slide.shapes.add_textbox(Inches(0.55), Inches(7.1), Inches(12.1), Inches(0.22))
    text = f"{TEMPLATE_VERSION}"
    if generated_at:
        text += f" | Generated: {generated_at}"
    if ai_provider:
        text += f" | Narrative: {ai_provider}"
    p = footer.text_frame.paragraphs[0]
    p.text = text
    p.font.size = Pt(7)
    p.font.color.rgb = BRAND_MUTED


def _set_title(slide, title: str, subtitle: str | None = None) -> None:
    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.18))
    _set_fill(band, BRAND_BLUE)
    band.line.fill.background()

    title_box = slide.shapes.add_textbox(Inches(0.55), Inches(0.38), Inches(12.2), Inches(0.52))
    p = title_box.text_frame.paragraphs[0]
    p.text = title
    _set_paragraph_font(p, 24, BRAND_DARK, bold=True)
    if subtitle:
        sub = slide.shapes.add_textbox(Inches(0.57), Inches(0.88), Inches(11.8), Inches(0.35))
        p2 = sub.text_frame.paragraphs[0]
        p2.text = subtitle
        _set_paragraph_font(p2, 10, BRAND_MUTED)


def _add_metric(slide, x: float, y: float, label: str, value: Any, accent: RGBColor = BRAND_BLUE) -> None:
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(2.55), Inches(0.92))
    _set_fill(box, BRAND_LIGHT)
    _set_line(box, RGBColor(218, 226, 235))
    box.text_frame.clear()
    p = box.text_frame.paragraphs[0]
    p.text = str(value)
    p.alignment = PP_ALIGN.CENTER
    _set_paragraph_font(p, 22, accent, bold=True)
    p2 = box.text_frame.add_paragraph()
    p2.text = label
    p2.alignment = PP_ALIGN.CENTER
    _set_paragraph_font(p2, 8, BRAND_MUTED)


def _add_text_block(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    text: str,
    font_size: int = 13,
    color: RGBColor = BRAND_DARK,
) -> None:
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.clear()
    paragraphs = [part.strip() for part in str(text).splitlines() if part.strip()] or [""]
    for idx, paragraph_text in enumerate(paragraphs):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = paragraph_text
        _set_paragraph_font(p, font_size, color)
        if idx:
            p.space_before = Pt(5)


def _add_bullets(slide, x: float, y: float, w: float, h: float, lines: Iterable[str], font_size: int = 12) -> None:
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.clear()
    for idx, line in enumerate(lines):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = str(line)
        p.level = 0
        _set_paragraph_font(p, font_size, BRAND_DARK)


def _add_bar_chart(slide, x: float, y: float, w: float, h: float, title: str, data: dict[str, int]) -> None:
    chart_data = CategoryChartData()
    items = list(data.items())[:8]
    if not items:
        items = [("No data", 0)]
    chart_data.categories = [str(key)[:34] for key, _ in items]
    chart_data.add_series(title, [int(value) for _, value in items])

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h),
        chart_data,
    ).chart
    chart.has_legend = False
    chart.chart_title.has_text_frame = True
    chart.chart_title.text_frame.text = title
    chart.chart_title.text_frame.paragraphs[0].font.size = Pt(12)
    chart.chart_title.text_frame.paragraphs[0].font.bold = True
    chart.value_axis.tick_labels.font.size = Pt(8)
    chart.category_axis.tick_labels.font.size = Pt(8)


def _add_column_chart(slide, x: float, y: float, w: float, h: float, title: str, data: dict[str, int]) -> None:
    chart_data = CategoryChartData()
    items = list(data.items())[:8]
    if not items:
        items = [("No data", 0)]
    chart_data.categories = [str(key)[:18] for key, _ in items]
    chart_data.add_series(title, [int(value) for _, value in items])

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h),
        chart_data,
    ).chart
    chart.has_legend = False
    chart.chart_title.has_text_frame = True
    chart.chart_title.text_frame.text = title
    chart.chart_title.text_frame.paragraphs[0].font.size = Pt(12)
    chart.chart_title.text_frame.paragraphs[0].font.bold = True
    chart.value_axis.tick_labels.font.size = Pt(8)
    chart.category_axis.tick_labels.font.size = Pt(8)


def _truncate(value: Any, max_chars: int = 110) -> str:
    if pd.isna(value):
        return ""
    text = str(value).replace("\n", " ").strip()
    return text if len(text) <= max_chars else text[: max_chars - 1] + "..."


def _add_table(slide, x: float, y: float, w: float, h: float, df: pd.DataFrame, max_rows: int = 8) -> None:
    view = df.head(max_rows).copy()
    if view.empty:
        view = pd.DataFrame({"Message": ["No records found for selected filters."]})

    rows, cols = view.shape[0] + 1, view.shape[1]
    table = slide.shapes.add_table(rows, cols, Inches(x), Inches(y), Inches(w), Inches(h)).table

    for col_idx, col in enumerate(view.columns):
        cell = table.cell(0, col_idx)
        cell.text = str(col)
        _set_fill(cell, BRAND_DARK)
        for paragraph in cell.text_frame.paragraphs:
            _set_paragraph_font(paragraph, 7, WHITE, bold=True)

    for row_idx, (_, row) in enumerate(view.iterrows(), start=1):
        for col_idx, col in enumerate(view.columns):
            cell = table.cell(row_idx, col_idx)
            cell.text = _truncate(row[col])
            for paragraph in cell.text_frame.paragraphs:
                _set_paragraph_font(paragraph, 6, BRAND_DARK)


def _activity_table(metrics: dict[str, Any]) -> pd.DataFrame:
    return pd.DataFrame(metrics.get("activity_breakdown", {}).items(), columns=["Activity Type", "Count"])


def _appendix_table(df: pd.DataFrame) -> pd.DataFrame:
    cols = [col for col in PRESENTATION_COLUMNS if col in df.columns]
    appendix = df[cols].copy() if cols else pd.DataFrame()
    return appendix.rename(columns={
        "delivery_date": "Delivery",
        "role": "Role",
        "ta_area": "TA",
        "study_id": "Study",
        "activity_type": "Activity",
        "submitted_by_name": "Owner",
        "team_members_display": "Team",
        "comments": "Comments",
        "slide_category": "Category",
        "impact_level": "Impact",
    })


def generate_achievement_ppt(
    df: pd.DataFrame,
    report_title: str,
    period_label: str,
    narrative_text: str | None = None,
    ai_provider: str = "rule_based",
    ai_model: str = "",
    generated_at: str | None = None,
    filters: dict[str, Any] | None = None,
    metrics: dict[str, Any] | None = None,
) -> bytes:
    """Generate the fixed-template PPTX from filtered achievement data only."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    metrics = metrics or calculate_report_metrics(df)
    summary = metrics["summary"]
    quality = metrics.get("quality_complexity", {})
    narrative = narrative_text or build_rule_based_narrative(metrics, period_label)

    props = prs.core_properties
    props.title = report_title
    props.subject = f"{TEMPLATE_VERSION}; ai_provider={ai_provider}; generated_at={generated_at or ''}"
    props.comments = f"Filter trace: {filters or {}}"

    # Slide 1
    slide = prs.slides.add_slide(blank)
    _set_title(slide, report_title, "Achievement report generated from filtered tracker data")
    _add_text_block(
        slide,
        0.7,
        1.45,
        10.8,
        0.82,
        f"Reporting Period: {period_label}\nTemplate Version: {TEMPLATE_VERSION}",
        font_size=16,
    )
    _add_metric(slide, 0.75, 3.95, "Achievements", summary["total_achievements"], BRAND_BLUE)
    _add_metric(slide, 3.55, 3.95, "Studies / Workstreams", summary["unique_studies"], BRAND_TEAL)
    _add_metric(slide, 6.35, 3.95, "Contributors", summary["unique_contributors"], BRAND_ORANGE)
    _add_metric(slide, 9.15, 3.95, "CF-related Count", summary["cf_total"], BRAND_BLUE)
    _add_footer(slide, generated_at, ai_provider)

    # Slide 2
    slide = prs.slides.add_slide(blank)
    _set_title(slide, "Executive Summary", period_label)
    _add_text_block(slide, 0.7, 1.25, 8.1, 4.25, narrative, font_size=14)
    _add_metric(slide, 9.25, 1.35, "Achievements", summary["total_achievements"], BRAND_BLUE)
    _add_metric(slide, 9.25, 2.55, "High Impact", quality.get("high_impact_count", 0), BRAND_ORANGE)
    _add_metric(slide, 9.25, 3.75, "Quality / Inspection", quality.get("quality_inspection_count", 0), BRAND_TEAL)
    _add_bullets(
        slide,
        0.9,
        5.65,
        11.2,
        0.8,
        [
            "Narrative text is editable and should be reviewed before final distribution.",
            "All numeric values are sourced from application-calculated metrics.",
        ],
        font_size=10,
    )
    _add_footer(slide, generated_at, ai_provider)

    # Slide 3
    slide = prs.slides.add_slide(blank)
    _set_title(slide, "Achievement Overview Dashboard", period_label)
    _add_metric(slide, 0.65, 1.15, "Achievements", summary["total_achievements"], BRAND_BLUE)
    _add_metric(slide, 3.2, 1.15, "Studies / Workstreams", summary["unique_studies"], BRAND_TEAL)
    _add_metric(slide, 5.75, 1.15, "Contributors", summary["unique_contributors"], BRAND_ORANGE)
    _add_metric(slide, 8.3, 1.15, "CF Total", summary["cf_total"], BRAND_BLUE)
    _add_bar_chart(slide, 0.7, 2.55, 5.75, 3.95, "Slide Category Breakdown", metrics.get("category_breakdown", {}))
    _add_column_chart(slide, 7.0, 2.55, 5.2, 3.95, "Impact Level Breakdown", metrics.get("impact_breakdown", {}))
    _add_footer(slide, generated_at, ai_provider)

    # Slide 4
    slide = prs.slides.add_slide(blank)
    _set_title(slide, "Activity Type Breakdown", period_label)
    _add_bar_chart(slide, 0.7, 1.25, 6.25, 5.3, "Activity Type Count", metrics.get("activity_breakdown", {}))
    _add_table(slide, 7.35, 1.35, 5.25, 4.75, _activity_table(metrics), max_rows=10)
    _add_footer(slide, generated_at, ai_provider)

    # Slide 5
    slide = prs.slides.add_slide(blank)
    _set_title(slide, "Project / Study Highlights", period_label)
    _add_table(slide, 0.45, 1.18, 12.4, 5.75, build_study_highlights(df, limit=9), max_rows=9)
    _add_footer(slide, generated_at, ai_provider)

    # Slide 6
    slide = prs.slides.add_slide(blank)
    _set_title(slide, "People Contribution View", period_label)
    people_df = build_people_summary(df)
    if not people_df.empty:
        people_df = people_df[["person_name", "achievement_count", "main_activity"]].copy()
        people_df.columns = ["Person", "Achievement Count", "Main Activity"]
    _add_table(slide, 0.65, 1.22, 6.0, 5.55, people_df, max_rows=12)
    people_chart = dict(zip(people_df.get("Person", []), people_df.get("Achievement Count", []))) if not people_df.empty else {}
    _add_bar_chart(slide, 7.1, 1.55, 5.25, 4.8, "Top Contributors", people_chart)
    _add_footer(slide, generated_at, ai_provider)

    # Slide 7
    slide = prs.slides.add_slide(blank)
    _set_title(slide, "Quality / Complexity Highlights", period_label)
    _add_metric(slide, 0.7, 1.15, "High Impact", quality.get("high_impact_count", 0), BRAND_ORANGE)
    _add_metric(slide, 3.5, 1.15, "Quality / Inspection", quality.get("quality_inspection_count", 0), BRAND_TEAL)
    _add_metric(slide, 6.3, 1.15, "Custom Function", quality.get("custom_function_count", 0), BRAND_BLUE)
    _add_metric(slide, 9.1, 1.15, "CF Major", quality.get("cf_major", 0), BRAND_ORANGE)
    _add_table(slide, 0.45, 2.55, 12.4, 4.15, build_quality_complexity_highlights(df, limit=7), max_rows=7)
    _add_footer(slide, generated_at, ai_provider)

    # Slide 8
    slide = prs.slides.add_slide(blank)
    _set_title(slide, "Appendix Detail Table", period_label)
    _add_table(slide, 0.28, 1.05, 12.8, 6.0, _appendix_table(df), max_rows=10)
    _add_footer(slide, generated_at, ai_provider)

    stream = io.BytesIO()
    prs.save(stream)
    return stream.getvalue()
