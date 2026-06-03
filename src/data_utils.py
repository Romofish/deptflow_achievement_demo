from __future__ import annotations

import io
import re
from dataclasses import dataclass
from typing import Iterable

import pandas as pd

RAW_TO_STD_COLUMNS = {
    "ID": "achievement_id",
    "Role": "role",
    "TA / Non-Project": "ta_area",
    "Study ID": "study_id",
    "Activity Type": "activity_type",
    "PPC Category (Multiple Choice)": "ppc_category",
    "Start Date of Activity": "start_date",
    "Delivery Date": "delivery_date",
    "User Submitted": "submitted_by",
    "Team Members": "team_members_raw",
    "Comments or Deliverable Details": "comments",
    "CF%23 Standards/minor update": "cf_standards_minor",
    "CF%23 Major update": "cf_major",
    "CF%23 Study Specific": "cf_study_specific",
    "Status (Confirm by LM only)": "lm_status",
}

PRESENTATION_COLUMNS = [
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
]

SLIDE_CATEGORY_RULES = [
    ("Quality / Inspection", ["inspection", "nmpa", "submission", "acr f", "acrf", "cfdi"]),
    ("Database Delivery", ["db go-live", "go-live", "final dbl", "snapshot", "setup", "programming support"]),
    ("PPC Delivery", ["ppc", "post production"]),
    ("Custom Function", ["custom function", "cf programming", "cf peer"]),
    ("SME / Training", ["sme", "training", "tpd", "data review"]),
    ("Rule File", ["rule file", "up-version"]),
]


def _read_bytes_to_text(uploaded_or_path) -> str:
    if hasattr(uploaded_or_path, "read"):
        if hasattr(uploaded_or_path, "seek"):
            uploaded_or_path.seek(0)
        content = uploaded_or_path.read()
        if isinstance(content, str):
            return content
        return content.decode("utf-8-sig", errors="replace")
    with open(uploaded_or_path, "r", encoding="utf-8-sig", errors="replace") as f:
        return f.read()


def read_achievement_csv(uploaded_or_path) -> pd.DataFrame:
    """Read SharePoint-exported CSV. If the first row is ListSchema, skip it."""
    text = _read_bytes_to_text(uploaded_or_path)
    first_line = text.splitlines()[0] if text.splitlines() else ""
    skiprows = 1 if first_line.startswith("ListSchema=") else 0
    return pd.read_csv(io.StringIO(text), skiprows=skiprows)


def email_to_name(value: str | float | None) -> str:
    if pd.isna(value) or value is None:
        return ""
    value = str(value).strip()
    if not value:
        return ""
    local = value.split("@")[0]
    local = local.replace(".", " ").replace("-", " ").replace("_", " ")
    return " ".join(part.capitalize() for part in local.split() if part)


def split_people(value):
    """
    Split people fields such as:
    - user1@company.com;user2@company.com
    - user1@company.com, user2@company.com
    - user1@company.com
    """
    if pd.isna(value):
        return []

    text = str(value).strip()
    if not text:
        return []

    # Split by semicolon, comma, newline, or pipe
    parts = re.split(r"[;,\n|]+", text)

    people = []
    for part in parts:
        person = part.strip()
        if person:
            people.append(person)

    return people


def infer_slide_category(row: pd.Series) -> str:
    text = " ".join([
        str(row.get("activity_type", "")),
        str(row.get("comments", "")),
        str(row.get("ppc_category", "")),
    ]).lower()
    for category, keywords in SLIDE_CATEGORY_RULES:
        if any(k in text for k in keywords):
            return category
    return "Other Delivery"


def infer_impact_level(row: pd.Series) -> str:
    text = " ".join([
        str(row.get("activity_type", "")),
        str(row.get("comments", "")),
        str(row.get("ppc_category", "")),
    ]).lower()
    high_terms = ["inspection", "nmpa", "cfdi", "go-live", "final dbl", "submission", "major"]
    medium_terms = ["ppc", "custom function", "rule file", "setup", "snapshot", "sme"]
    if any(t in text for t in high_terms):
        return "High"
    if any(t in text for t in medium_terms):
        return "Medium"
    return "Low"


def normalize_achievements(df_raw: pd.DataFrame) -> pd.DataFrame:
    df = df_raw.rename(columns=RAW_TO_STD_COLUMNS).copy()
    for col in RAW_TO_STD_COLUMNS.values():
        if col not in df.columns:
            df[col] = pd.NA

    for date_col in ["start_date", "delivery_date"]:
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce", utc=True).dt.tz_convert(None).dt.date

    for num_col in ["cf_standards_minor", "cf_major", "cf_study_specific"]:
        df[num_col] = pd.to_numeric(df[num_col], errors="coerce").fillna(0).astype(int)

    for text_col in ["role", "ta_area", "study_id", "activity_type", "ppc_category", "submitted_by", "team_members_raw", "comments", "lm_status"]:
        df[text_col] = df[text_col].fillna("").astype(str).str.strip()

    df["submitted_by_name"] = df["submitted_by"].apply(email_to_name)
    df["team_members_list"] = df["team_members_raw"].apply(split_people)
    df["team_members_display"] = df["team_members_list"].apply(lambda xs: "; ".join(email_to_name(x) for x in xs))
    df["all_people"] = df.apply(lambda r: sorted(set([r["submitted_by"]] + r["team_members_list"])), axis=1)
    df["all_people_display"] = df["all_people"].apply(lambda xs: "; ".join(email_to_name(x) for x in xs if x))
    df["slide_category"] = df.apply(infer_slide_category, axis=1)
    df["impact_level"] = df.apply(infer_impact_level, axis=1)
    df["reportable_flag"] = True
    df["month"] = pd.to_datetime(df["delivery_date"], errors="coerce").dt.to_period("M").astype(str).replace("NaT", "")
    df["quarter"] = pd.to_datetime(df["delivery_date"], errors="coerce").dt.to_period("Q").astype(str).replace("NaT", "")
    df["cf_total"] = df[["cf_standards_minor", "cf_major", "cf_study_specific"]].sum(axis=1)
    return df


def filter_achievements(
    df: pd.DataFrame,
    start_date=None,
    end_date=None,
    roles: Iterable[str] | None = None,
    ta_areas: Iterable[str] | None = None,
    studies: Iterable[str] | None = None,
    people: Iterable[str] | None = None,
    activity_types: Iterable[str] | None = None,
    categories: Iterable[str] | None = None,
    impact_levels: Iterable[str] | None = None,
) -> pd.DataFrame:
    out = df.copy()
    if start_date:
        out = out[pd.to_datetime(out["delivery_date"], errors="coerce") >= pd.to_datetime(start_date)]
    if end_date:
        out = out[pd.to_datetime(out["delivery_date"], errors="coerce") <= pd.to_datetime(end_date)]
    filters = [
        ("role", roles),
        ("ta_area", ta_areas),
        ("study_id", studies),
        ("activity_type", activity_types),
        ("slide_category", categories),
        ("impact_level", impact_levels),
    ]
    for col, values in filters:
        if values:
            out = out[out[col].isin(values)]
    if people:
        people = set(people)
        out = out[out["all_people"].apply(lambda xs: bool(people.intersection(set(xs))))]
    return out.reset_index(drop=True)


def build_report_summary(df: pd.DataFrame) -> dict:
    if df.empty:
        return {
            "total_achievements": 0,
            "unique_studies": 0,
            "unique_contributors": 0,
            "role_breakdown": {},
            "activity_breakdown": {},
            "category_breakdown": {},
            "top_studies": {},
            "top_people": {},
            "cf_total": 0,
        }
    people = []
    for xs in df["all_people"]:
        people.extend([x for x in xs if x])
    return {
        "total_achievements": int(len(df)),
        "unique_studies": int(df.loc[df["study_id"].ne(""), "study_id"].nunique()),
        "unique_contributors": int(pd.Series(people).nunique()) if people else 0,
        "role_breakdown": df["role"].value_counts().to_dict(),
        "activity_breakdown": df["activity_type"].value_counts().head(10).to_dict(),
        "category_breakdown": df["slide_category"].value_counts().to_dict(),
        "top_studies": df["study_id"].value_counts().head(8).to_dict(),
        "top_people": pd.Series(people).value_counts().head(8).to_dict() if people else {},
        "cf_total": int(df["cf_total"].sum()),
    }


def build_rule_based_narrative(summary: dict, period_label: str = "selected period") -> str:
    total = summary.get("total_achievements", 0)
    studies = summary.get("unique_studies", 0)
    contributors = summary.get("unique_contributors", 0)
    top_activities = list(summary.get("activity_breakdown", {}).keys())[:3]
    top_categories = list(summary.get("category_breakdown", {}).keys())[:3]
    activity_text = ", ".join(top_activities) if top_activities else "no major activity category"
    category_text = ", ".join(top_categories) if top_categories else "no categorized delivery area"
    return (
        f"During {period_label}, the team recorded {total} achievements across {studies} studies/non-project workstreams, "
        f"with contributions from {contributors} colleagues. The main delivery focus areas were {activity_text}. "
        f"From a reporting perspective, the achievements are mainly grouped under {category_text}. "
        "These outputs can be reviewed by line managers and converted into a standardized leadership report through the fixed slide template."
    )


def build_people_summary(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, row in df.iterrows():
        for person in row["all_people"]:
            if person:
                rows.append({
                    "person_email": person,
                    "person_name": email_to_name(person),
                    "achievement_id": row["achievement_id"],
                    "role": row["role"],
                    "study_id": row["study_id"],
                    "activity_type": row["activity_type"],
                    "slide_category": row["slide_category"],
                })
    if not rows:
        return pd.DataFrame(columns=["person_name", "achievement_count", "main_activity"])
    people_df = pd.DataFrame(rows)
    result = people_df.groupby(["person_email", "person_name"], as_index=False).agg(
        achievement_count=("achievement_id", "nunique"),
        main_activity=("activity_type", lambda s: s.value_counts().index[0] if not s.empty else ""),
    )
    return result.sort_values("achievement_count", ascending=False)
