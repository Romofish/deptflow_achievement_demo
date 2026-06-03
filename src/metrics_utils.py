from __future__ import annotations

import csv
import json
from pathlib import Path
from typing import Any, Iterable

import pandas as pd

from .data_utils import build_report_summary


REPORT_HISTORY_COLUMNS = [
    "generated_at",
    "report_title",
    "template_version",
    "ai_provider",
    "ai_model",
    "rows_loaded",
    "rows_filtered",
    "period_label",
    "filter_json",
    "metrics_json",
    "slide_spec_json",
    "style_preset",
    "preview_generated",
    "preview_renderer",
    "narrative_chars",
]


def _clean_list(values: Iterable[Any] | None) -> list[str]:
    if not values:
        return []
    return [str(value) for value in values if str(value).strip()]


def _value_counts(
    df: pd.DataFrame,
    column: str,
    limit: int | None = None,
    drop_blank: bool = True,
) -> dict[str, int]:
    if df.empty or column not in df.columns:
        return {}

    series = df[column].fillna("").astype(str).str.strip()
    if drop_blank:
        series = series[series.ne("")]

    counts = series.value_counts()
    if limit:
        counts = counts.head(limit)
    return {str(key): int(value) for key, value in counts.items()}


def _sum_int(df: pd.DataFrame, column: str) -> int:
    if df.empty or column not in df.columns:
        return 0
    return int(pd.to_numeric(df[column], errors="coerce").fillna(0).sum())


def _text_series(df: pd.DataFrame, column: str) -> pd.Series:
    if column not in df.columns:
        return pd.Series([""] * len(df), index=df.index)
    return df[column].fillna("").astype(str).str.strip()


def _numeric_series(df: pd.DataFrame, column: str) -> pd.Series:
    if column not in df.columns:
        return pd.Series([0] * len(df), index=df.index)
    return pd.to_numeric(df[column], errors="coerce").fillna(0)


def _mode_text(values: pd.Series) -> str:
    values = values.fillna("").astype(str).str.strip()
    values = values[values.ne("")]
    if values.empty:
        return ""
    return str(values.value_counts().index[0])


def calculate_report_metrics(df: pd.DataFrame) -> dict[str, Any]:
    """Calculate all report metrics from the already-filtered achievement data."""
    summary = build_report_summary(df)
    quality_complexity = {
        "high_impact_count": int(df["impact_level"].eq("High").sum()) if "impact_level" in df else 0,
        "medium_impact_count": int(df["impact_level"].eq("Medium").sum()) if "impact_level" in df else 0,
        "low_impact_count": int(df["impact_level"].eq("Low").sum()) if "impact_level" in df else 0,
        "quality_inspection_count": int(df["slide_category"].eq("Quality / Inspection").sum()) if "slide_category" in df else 0,
        "custom_function_count": int(df["slide_category"].eq("Custom Function").sum()) if "slide_category" in df else 0,
        "database_delivery_count": int(df["slide_category"].eq("Database Delivery").sum()) if "slide_category" in df else 0,
        "cf_standards_minor": _sum_int(df, "cf_standards_minor"),
        "cf_major": _sum_int(df, "cf_major"),
        "cf_study_specific": _sum_int(df, "cf_study_specific"),
        "cf_total": _sum_int(df, "cf_total"),
    }

    return {
        "summary": summary,
        "activity_breakdown": _value_counts(df, "activity_type", limit=10),
        "category_breakdown": _value_counts(df, "slide_category", limit=10),
        "impact_breakdown": _value_counts(df, "impact_level", limit=10),
        "role_breakdown": _value_counts(df, "role", limit=10),
        "ta_breakdown": _value_counts(df, "ta_area", limit=10),
        "monthly_trend": _value_counts(df, "month", limit=24),
        "top_studies": _value_counts(df, "study_id", limit=8),
        "top_people": summary.get("top_people", {}),
        "quality_complexity": quality_complexity,
    }


def compact_metrics_for_ai(metrics: dict[str, Any]) -> dict[str, Any]:
    """Return the only calculated facts the AI is allowed to reference."""
    summary = metrics.get("summary", {})
    return {
        "total_achievements": summary.get("total_achievements", 0),
        "unique_studies": summary.get("unique_studies", 0),
        "unique_contributors": summary.get("unique_contributors", 0),
        "cf_total": summary.get("cf_total", 0),
        "activity_breakdown": metrics.get("activity_breakdown", {}),
        "category_breakdown": metrics.get("category_breakdown", {}),
        "impact_breakdown": metrics.get("impact_breakdown", {}),
        "top_studies": metrics.get("top_studies", {}),
        "quality_complexity": metrics.get("quality_complexity", {}),
    }


def build_study_highlights(df: pd.DataFrame, limit: int = 8) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Study ID", "TA", "Achievements", "Top Activity", "High Impact", "CF Total", "Latest Delivery"])

    work = df.copy()
    work["study_id"] = work["study_id"].fillna("").astype(str).str.strip().replace("", "Unspecified")
    grouped = work.groupby("study_id", dropna=False).agg(
        ta_area=("ta_area", _mode_text),
        achievement_count=("study_id", "size"),
        top_activity=("activity_type", _mode_text),
        high_impact_count=("impact_level", lambda s: int(s.eq("High").sum())),
        cf_total=("cf_total", "sum"),
        latest_delivery=("delivery_date", lambda s: str(pd.to_datetime(s, errors="coerce").max().date()) if pd.to_datetime(s, errors="coerce").notna().any() else ""),
    )
    grouped = grouped.sort_values(["achievement_count", "high_impact_count"], ascending=False).head(limit).reset_index()
    grouped.columns = ["Study ID", "TA", "Achievements", "Top Activity", "High Impact", "CF Total", "Latest Delivery"]
    grouped["CF Total"] = grouped["CF Total"].fillna(0).astype(int)
    return grouped


def build_quality_complexity_highlights(df: pd.DataFrame, limit: int = 8) -> pd.DataFrame:
    columns = ["Delivery", "Study ID", "Activity", "Category", "Impact", "CF Total", "Comments"]
    if df.empty:
        return pd.DataFrame(columns=columns)

    cf_total = _numeric_series(df, "cf_total")
    mask = (
        _text_series(df, "impact_level").eq("High")
        | _text_series(df, "slide_category").isin(["Quality / Inspection", "Custom Function", "Database Delivery"])
        | cf_total.gt(0)
    )
    view = df.loc[mask].copy()
    if view.empty:
        view = df.copy()

    cols = ["delivery_date", "study_id", "activity_type", "slide_category", "impact_level", "cf_total", "comments"]
    view = view[[col for col in cols if col in view.columns]].head(limit)
    view = view.rename(columns={
        "delivery_date": "Delivery",
        "study_id": "Study ID",
        "activity_type": "Activity",
        "slide_category": "Category",
        "impact_level": "Impact",
        "cf_total": "CF Total",
        "comments": "Comments",
    })
    for col in columns:
        if col not in view.columns:
            view[col] = ""
    return view[columns]


def build_filter_trace(
    *,
    source_label: str,
    start_date: Any,
    end_date: Any,
    roles: Iterable[Any] | None,
    ta_areas: Iterable[Any] | None,
    studies: Iterable[Any] | None,
    people: Iterable[Any] | None,
    activity_types: Iterable[Any] | None,
    categories: Iterable[Any] | None,
    impact_levels: Iterable[Any] | None,
) -> dict[str, Any]:
    return {
        "source": source_label,
        "start_date": str(start_date),
        "end_date": str(end_date),
        "roles": _clean_list(roles),
        "ta_areas": _clean_list(ta_areas),
        "studies": _clean_list(studies),
        "people": _clean_list(people),
        "activity_types": _clean_list(activity_types),
        "categories": _clean_list(categories),
        "impact_levels": _clean_list(impact_levels),
    }


def trace_signature(filter_trace: dict[str, Any], metrics: dict[str, Any]) -> str:
    payload = {
        "filters": filter_trace,
        "metrics": compact_metrics_for_ai(metrics),
    }
    return json.dumps(payload, ensure_ascii=False, sort_keys=True, default=str)


def build_report_history_record(
    *,
    generated_at: str,
    report_title: str,
    template_version: str,
    ai_provider: str,
    ai_model: str,
    rows_loaded: int,
    rows_filtered: int,
    period_label: str,
    filter_trace: dict[str, Any],
    metrics: dict[str, Any],
    narrative_text: str,
    slide_spec: dict[str, Any] | None = None,
    style_preset: str = "",
    preview_generated: bool = False,
    preview_renderer: str = "html",
) -> dict[str, Any]:
    return {
        "generated_at": generated_at,
        "report_title": report_title,
        "template_version": template_version,
        "ai_provider": ai_provider,
        "ai_model": ai_model,
        "rows_loaded": int(rows_loaded),
        "rows_filtered": int(rows_filtered),
        "period_label": period_label,
        "filter_json": json.dumps(filter_trace, ensure_ascii=False, sort_keys=True, default=str),
        "metrics_json": json.dumps(compact_metrics_for_ai(metrics), ensure_ascii=False, sort_keys=True, default=str),
        "slide_spec_json": json.dumps(slide_spec or {}, ensure_ascii=False, sort_keys=True, default=str),
        "style_preset": style_preset,
        "preview_generated": bool(preview_generated),
        "preview_renderer": preview_renderer,
        "narrative_chars": len(narrative_text or ""),
    }


def append_report_history(history_path: str | Path, record: dict[str, Any]) -> None:
    path = Path(history_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists() and path.stat().st_size > 0:
        existing = pd.read_csv(path)
        if list(existing.columns) != REPORT_HISTORY_COLUMNS:
            for col in REPORT_HISTORY_COLUMNS:
                if col not in existing.columns:
                    existing[col] = ""
            existing = existing[REPORT_HISTORY_COLUMNS]
            existing.to_csv(path, index=False, encoding="utf-8-sig")

    write_header = not path.exists() or path.stat().st_size == 0

    with path.open("a", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=REPORT_HISTORY_COLUMNS, extrasaction="ignore")
        if write_header:
            writer.writeheader()
        writer.writerow(record)


def load_report_history(history_path: str | Path) -> pd.DataFrame:
    path = Path(history_path)
    if not path.exists() or path.stat().st_size == 0:
        return pd.DataFrame(columns=REPORT_HISTORY_COLUMNS)
    return pd.read_csv(path)
