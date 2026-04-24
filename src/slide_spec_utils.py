from __future__ import annotations

import copy
import json
from typing import Any

import pandas as pd

from .ai_utils import build_rule_based_narrative
from .data_utils import PRESENTATION_COLUMNS, build_people_summary
from .metrics_utils import build_quality_complexity_highlights, build_study_highlights


SLIDE_SPEC_VERSION = "slide-spec-v1"
FIXED_SLIDE_IDS = [1, 2, 3, 4, 5, 6, 7, 8]

STYLE_PRESETS = {
    "atlas": {
        "label": "Atlas Boardroom",
        "accent": "#2566A8",
        "accent_2": "#158B84",
        "accent_3": "#DE762C",
        "ink": "#162436",
        "muted": "#5C6979",
        "paper": "#F4F7FA",
        "surface": "#FFFFFF",
    },
    "onyx": {
        "label": "Onyx Control Room",
        "accent": "#66D9E8",
        "accent_2": "#95D5B2",
        "accent_3": "#FFB703",
        "ink": "#EAF2F8",
        "muted": "#9CB3C9",
        "paper": "#08111F",
        "surface": "#101B2D",
    },
    "linen": {
        "label": "Linen Editorial",
        "accent": "#9B5D32",
        "accent_2": "#376D63",
        "accent_3": "#C59339",
        "ink": "#2D261F",
        "muted": "#766B61",
        "paper": "#F7F1E8",
        "surface": "#FFFDF8",
    },
}

ALLOWED_CHART_TYPES = {"none", "bar", "column"}
ALLOWED_LAYOUT_VARIANTS = {
    "hero",
    "split",
    "dashboard",
    "chart_table",
    "table_full",
    "table_chart",
    "cards_table",
}
PATCHABLE_SLIDE_FIELDS = {
    "title",
    "subtitle",
    "narrative",
    "fields",
    "top_n",
    "sort_by",
    "chart_type",
    "layout_variant",
    "data_source",
}

FIELD_LABELS = {
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
    "cf_total": "CF Total",
    "achievement_count": "Achievements",
    "top_activity": "Top Activity",
    "high_impact_count": "High Impact",
    "latest_delivery": "Latest Delivery",
    "person_name": "Person",
    "main_activity": "Main Activity",
    "person_email": "Email",
}

DATA_SOURCE_FIELDS = {
    "none": [],
    "summary": [],
    "activity_breakdown": ["Activity Type", "Count"],
    "category_breakdown": ["Category", "Count"],
    "impact_breakdown": ["Impact", "Count"],
    "study_highlights": ["Study ID", "TA", "Achievements", "Top Activity", "High Impact", "CF Total", "Latest Delivery"],
    "people_summary": ["Person", "Achievement Count", "Main Activity"],
    "quality_highlights": ["Delivery", "Study ID", "Activity", "Category", "Impact", "CF Total", "Comments"],
    "appendix": [
        "delivery_date",
        "role",
        "ta_area",
        "study_id",
        "activity_type",
        "submitted_by_name",
        "team_members_display",
        "comments",
        "slide_category",
        "impact_level",
    ],
}


def normalize_style_preset(value: str | None) -> str:
    return value if value in STYLE_PRESETS else "atlas"


def allowed_fields() -> list[str]:
    fields = set()
    for values in DATA_SOURCE_FIELDS.values():
        fields.update(values)
    return sorted(fields)


def build_default_slide_spec(
    metrics: dict[str, Any],
    filtered_df: pd.DataFrame,
    *,
    report_title: str = "CDO China Achievement Report",
    period_label: str = "selected period",
    narrative_text: str | None = None,
    style_preset: str = "atlas",
) -> dict[str, Any]:
    narrative = narrative_text or build_rule_based_narrative(metrics, period_label)
    return {
        "version": SLIDE_SPEC_VERSION,
        "style_preset": normalize_style_preset(style_preset),
        "slides": [
            {
                "id": 1,
                "title": report_title,
                "subtitle": "Achievement report generated from filtered tracker data",
                "kind": "title",
                "data_source": "summary",
                "fields": [],
                "top_n": 4,
                "sort_by": "",
                "chart_type": "none",
                "layout_variant": "hero",
                "narrative": f"Reporting Period: {period_label}",
            },
            {
                "id": 2,
                "title": "Executive Summary",
                "subtitle": period_label,
                "kind": "narrative",
                "data_source": "summary",
                "fields": [],
                "top_n": 3,
                "sort_by": "",
                "chart_type": "none",
                "layout_variant": "split",
                "narrative": narrative,
            },
            {
                "id": 3,
                "title": "Achievement Overview Dashboard",
                "subtitle": period_label,
                "kind": "dashboard",
                "data_source": "category_breakdown",
                "fields": ["Category", "Count"],
                "top_n": 8,
                "sort_by": "Count",
                "chart_type": "bar",
                "layout_variant": "dashboard",
                "narrative": "",
            },
            {
                "id": 4,
                "title": "Activity Type Breakdown",
                "subtitle": period_label,
                "kind": "chart_table",
                "data_source": "activity_breakdown",
                "fields": ["Activity Type", "Count"],
                "top_n": 10,
                "sort_by": "Count",
                "chart_type": "bar",
                "layout_variant": "chart_table",
                "narrative": "",
            },
            {
                "id": 5,
                "title": "Project / Study Highlights",
                "subtitle": period_label,
                "kind": "table",
                "data_source": "study_highlights",
                "fields": ["Study ID", "TA", "Achievements", "Top Activity", "High Impact", "CF Total", "Latest Delivery"],
                "top_n": 9,
                "sort_by": "Achievements",
                "chart_type": "none",
                "layout_variant": "table_full",
                "narrative": "",
            },
            {
                "id": 6,
                "title": "People Contribution View",
                "subtitle": period_label,
                "kind": "people",
                "data_source": "people_summary",
                "fields": ["Person", "Achievement Count", "Main Activity"],
                "top_n": 12,
                "sort_by": "Achievement Count",
                "chart_type": "bar",
                "layout_variant": "table_chart",
                "narrative": "",
            },
            {
                "id": 7,
                "title": "Quality / Complexity Highlights",
                "subtitle": period_label,
                "kind": "quality",
                "data_source": "quality_highlights",
                "fields": ["Delivery", "Study ID", "Activity", "Category", "Impact", "CF Total", "Comments"],
                "top_n": 7,
                "sort_by": "CF Total",
                "chart_type": "none",
                "layout_variant": "cards_table",
                "narrative": "",
            },
            {
                "id": 8,
                "title": "Appendix Detail Table",
                "subtitle": period_label,
                "kind": "appendix",
                "data_source": "appendix",
                "fields": [
                    "delivery_date",
                    "role",
                    "ta_area",
                    "study_id",
                    "activity_type",
                    "submitted_by_name",
                    "team_members_display",
                    "comments",
                    "slide_category",
                    "impact_level",
                ],
                "top_n": 10,
                "sort_by": "delivery_date",
                "chart_type": "none",
                "layout_variant": "table_full",
                "narrative": "",
            },
        ],
    }


def sync_slide_spec_context(
    slide_spec: dict[str, Any],
    *,
    report_title: str,
    period_label: str,
    default_narrative: str,
) -> dict[str, Any]:
    spec = copy.deepcopy(slide_spec)
    spec["version"] = SLIDE_SPEC_VERSION
    spec["style_preset"] = normalize_style_preset(spec.get("style_preset"))
    slides = {int(slide.get("id", 0)): slide for slide in spec.get("slides", [])}
    if set(slides) != set(FIXED_SLIDE_IDS):
        spec = build_default_slide_spec({}, pd.DataFrame(), report_title=report_title, period_label=period_label, narrative_text=default_narrative)
        slides = {int(slide.get("id", 0)): slide for slide in spec["slides"]}

    slides[1]["title"] = report_title
    slides[1]["narrative"] = f"Reporting Period: {period_label}"
    for slide_id in FIXED_SLIDE_IDS:
        slides[slide_id]["subtitle"] = period_label if slide_id > 1 else slides[slide_id].get("subtitle", "")
    if not slides[2].get("narrative"):
        slides[2]["narrative"] = default_narrative
    spec["slides"] = [slides[slide_id] for slide_id in FIXED_SLIDE_IDS]
    return spec


def slide_spec_signature(slide_spec: dict[str, Any]) -> str:
    return json.dumps(slide_spec, ensure_ascii=False, sort_keys=True, default=str)


def get_slide(slide_spec: dict[str, Any], slide_id: int) -> dict[str, Any]:
    for slide in slide_spec.get("slides", []):
        if int(slide.get("id", 0)) == int(slide_id):
            return slide
    raise ValueError(f"Unknown slide id: {slide_id}")


def validate_slide_patch(patch: dict[str, Any], allowed_field_names: list[str] | None = None) -> tuple[bool, list[str]]:
    errors: list[str] = []
    allowed_field_names = allowed_field_names or allowed_fields()
    allowed_field_set = set(allowed_field_names)

    if not isinstance(patch, dict):
        return False, ["Patch must be a JSON object."]
    if "slides" not in patch and "style_preset" not in patch:
        return False, ["Patch must include 'slides' or 'style_preset'."]

    if "style_preset" in patch and patch["style_preset"] not in STYLE_PRESETS:
        errors.append(f"style_preset must be one of: {', '.join(STYLE_PRESETS)}.")

    slides = patch.get("slides", [])
    if slides and not isinstance(slides, list):
        errors.append("slides must be a list.")
        return False, errors

    for change in slides:
        if not isinstance(change, dict):
            errors.append("Each slide patch must be an object.")
            continue
        raw_slide_id = change.get("id")
        try:
            slide_id = int(raw_slide_id)
        except (TypeError, ValueError):
            slide_id = -1
        if slide_id not in FIXED_SLIDE_IDS:
            errors.append(f"Invalid slide id {raw_slide_id}; fixed slides are 1-8 and cannot be added or removed.")
        for key in change:
            if key == "id":
                continue
            if key not in PATCHABLE_SLIDE_FIELDS:
                errors.append(f"Field '{key}' is not patchable.")
        if "fields" in change:
            fields = change["fields"]
            if not isinstance(fields, list):
                errors.append("fields must be a list.")
            else:
                invalid = [field for field in fields if field not in allowed_field_set]
                if invalid:
                    errors.append(f"Unsupported fields: {', '.join(invalid)}.")
        if "top_n" in change:
            top_n = change["top_n"]
            if not isinstance(top_n, int) or top_n < 1 or top_n > 25:
                errors.append("top_n must be an integer from 1 to 25.")
        if "chart_type" in change and change["chart_type"] not in ALLOWED_CHART_TYPES:
            errors.append(f"chart_type must be one of: {', '.join(sorted(ALLOWED_CHART_TYPES))}.")
        if "layout_variant" in change and change["layout_variant"] not in ALLOWED_LAYOUT_VARIANTS:
            errors.append(f"layout_variant must be one of: {', '.join(sorted(ALLOWED_LAYOUT_VARIANTS))}.")
        if "data_source" in change and change["data_source"] not in DATA_SOURCE_FIELDS:
            errors.append(f"data_source must be one of: {', '.join(DATA_SOURCE_FIELDS)}.")
    return not errors, errors


def apply_slide_patch(slide_spec: dict[str, Any], patch: dict[str, Any]) -> dict[str, Any]:
    valid, errors = validate_slide_patch(patch)
    if not valid:
        raise ValueError("; ".join(errors))

    out = copy.deepcopy(slide_spec)
    if "style_preset" in patch:
        out["style_preset"] = normalize_style_preset(patch["style_preset"])

    slides = {int(slide["id"]): slide for slide in out.get("slides", [])}
    for change in patch.get("slides", []):
        slide = slides[int(change["id"])]
        for key, value in change.items():
            if key != "id":
                slide[key] = value

    out["slides"] = [slides[slide_id] for slide_id in FIXED_SLIDE_IDS]
    return out


def get_data_frame_for_slide(slide: dict[str, Any], df: pd.DataFrame, metrics: dict[str, Any]) -> pd.DataFrame:
    data_source = slide.get("data_source", "none")
    top_n = int(slide.get("top_n") or 8)

    if data_source == "activity_breakdown":
        table = pd.DataFrame(metrics.get("activity_breakdown", {}).items(), columns=["Activity Type", "Count"])
    elif data_source == "category_breakdown":
        table = pd.DataFrame(metrics.get("category_breakdown", {}).items(), columns=["Category", "Count"])
    elif data_source == "impact_breakdown":
        table = pd.DataFrame(metrics.get("impact_breakdown", {}).items(), columns=["Impact", "Count"])
    elif data_source == "study_highlights":
        table = build_study_highlights(df, limit=max(top_n, 1))
    elif data_source == "people_summary":
        people_df = build_people_summary(df)
        if not people_df.empty:
            people_df = people_df[["person_name", "achievement_count", "main_activity"]].copy()
            people_df.columns = ["Person", "Achievement Count", "Main Activity"]
        table = people_df
    elif data_source == "quality_highlights":
        table = build_quality_complexity_highlights(df, limit=max(top_n, 1))
    elif data_source == "appendix":
        cols = [col for col in PRESENTATION_COLUMNS if col in df.columns]
        table = df[cols].copy() if cols else pd.DataFrame()
        table = table.rename(columns={col: FIELD_LABELS.get(col, col) for col in table.columns})
    else:
        return pd.DataFrame()

    fields = slide.get("fields") or []
    if fields:
        renamed = table.rename(columns={col: FIELD_LABELS.get(col, col) for col in table.columns})
        available = [field for field in fields if field in renamed.columns]
        table = renamed[available] if available else renamed

    sort_by = slide.get("sort_by")
    if sort_by and sort_by in table.columns:
        table = table.sort_values(sort_by, ascending=False, kind="stable")

    return table.head(top_n).reset_index(drop=True)
