from __future__ import annotations

import datetime as dt
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from src.data_utils import (
    PRESENTATION_COLUMNS,
    build_people_summary,
    build_report_summary,
    build_rule_based_narrative,
    filter_achievements,
    normalize_achievements,
    read_achievement_csv,
)
from src.ppt_utils import generate_achievement_ppt

APP_DIR = Path(__file__).parent
SAMPLE_CSV = APP_DIR / "sample_data" / "2026-CDO China Achievements sample.csv"

st.set_page_config(
    page_title="DeptFlow Achievement Reporter",
    page_icon="📊",
    layout="wide",
)

st.title("DeptFlow Achievement Reporter")
st.caption("Streamlit MVP · Filter achievements · Preview metrics · Generate fixed-template PPT")

with st.sidebar:
    st.header("1. Data Source")
    uploaded = st.file_uploader("Upload SharePoint achievement CSV", type=["csv"])
    use_sample = st.toggle("Use bundled sample data", value=uploaded is None and SAMPLE_CSV.exists())

    st.header("2. Report Settings")
    report_title = st.text_input("Report title", "CDO China Achievement Report")

try:
    if uploaded is not None:
        source = uploaded
    elif use_sample and SAMPLE_CSV.exists():
        source = SAMPLE_CSV
    else:
        st.info("Upload a SharePoint achievement CSV to build the dashboard and PPTX report.")
        st.stop()

    raw_df = read_achievement_csv(source)
    df = normalize_achievements(raw_df)
except Exception as exc:
    st.error(f"Failed to load achievement data: {exc}")
    st.stop()

min_date = pd.to_datetime(df["delivery_date"], errors="coerce").min()
max_date = pd.to_datetime(df["delivery_date"], errors="coerce").max()
if pd.isna(min_date) or pd.isna(max_date):
    default_range = (dt.date.today().replace(day=1), dt.date.today())
else:
    default_range = (min_date.date(), max_date.date())

with st.sidebar:
    st.header("3. Filters")
    date_range = st.date_input("Delivery date range", value=default_range)
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
    else:
        start_date, end_date = default_range

    roles = st.multiselect("Role", sorted([x for x in df["role"].unique() if x]))
    ta_areas = st.multiselect("TA / Non-Project", sorted([x for x in df["ta_area"].unique() if x]))
    studies = st.multiselect("Study ID", sorted([x for x in df["study_id"].unique() if x]))
    activity_types = st.multiselect("Activity Type", sorted([x for x in df["activity_type"].unique() if x]))
    categories = st.multiselect("Slide Category", sorted([x for x in df["slide_category"].unique() if x]))
    impact_levels = st.multiselect("Impact Level", ["High", "Medium", "Low"])

    all_people = sorted({person for people in df["all_people"] for person in people if person})
    people = st.multiselect("Person", all_people, format_func=lambda email: email.split("@")[0].replace(".", " ").title())

filtered = filter_achievements(
    df,
    start_date=start_date,
    end_date=end_date,
    roles=roles,
    ta_areas=ta_areas,
    studies=studies,
    people=people,
    activity_types=activity_types,
    categories=categories,
    impact_levels=impact_levels,
)
summary = build_report_summary(filtered)
period_label = f"{start_date} to {end_date}"

kpi1, kpi2, kpi3, kpi4 = st.columns(4)
kpi1.metric("Achievements", summary["total_achievements"])
kpi2.metric("Studies / Workstreams", summary["unique_studies"])
kpi3.metric("Contributors", summary["unique_contributors"])
kpi4.metric("CF-related Count", summary["cf_total"])

st.divider()

tab_overview, tab_detail, tab_people, tab_report = st.tabs([
    "Overview Dashboard",
    "Achievement Review Table",
    "People View",
    "PPT Builder",
])

with tab_overview:
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Activity Type Breakdown")
        activity_df = pd.DataFrame(summary["activity_breakdown"].items(), columns=["Activity Type", "Count"])
        if not activity_df.empty:
            st.plotly_chart(px.bar(activity_df, x="Count", y="Activity Type", orientation="h"), use_container_width=True)
        else:
            st.info("No data for selected filters.")
    with c2:
        st.subheader("Slide Category Breakdown")
        category_df = pd.DataFrame(summary["category_breakdown"].items(), columns=["Slide Category", "Count"])
        if not category_df.empty:
            st.plotly_chart(px.pie(category_df, names="Slide Category", values="Count", hole=0.45), use_container_width=True)
        else:
            st.info("No data for selected filters.")

    st.subheader("Monthly Trend")
    trend = filtered.groupby("month", as_index=False).size().rename(columns={"size": "Count"}) if not filtered.empty else pd.DataFrame(columns=["month", "Count"])
    if not trend.empty:
        st.plotly_chart(px.line(trend, x="month", y="Count", markers=True), use_container_width=True)
    else:
        st.info("No monthly trend available.")

with tab_detail:
    st.subheader("Achievement Review Table")
    st.caption("This table is the human-review layer before formal PPT generation. In a later version, LM Confirmed / Reportable Flag can be edited here.")
    display_cols = [col for col in PRESENTATION_COLUMNS if col in filtered.columns]
    view = filtered[display_cols].copy()
    view = view.rename(columns={
        "delivery_date": "Delivery Date",
        "role": "Role",
        "ta_area": "TA / Non-Project",
        "study_id": "Study ID",
        "activity_type": "Activity Type",
        "submitted_by_name": "Owner",
        "team_members_display": "Team Members",
        "comments": "Comments / Deliverable Details",
        "slide_category": "Slide Category",
        "impact_level": "Impact Level",
    })
    st.dataframe(view, use_container_width=True, hide_index=True)
    csv_bytes = view.to_csv(index=False).encode("utf-8-sig")
    st.download_button("Download filtered CSV", csv_bytes, file_name="filtered_achievements.csv", mime="text/csv")

with tab_people:
    st.subheader("People Contribution View")
    people_df = build_people_summary(filtered)
    st.dataframe(people_df, use_container_width=True, hide_index=True)
    if not people_df.empty:
        st.plotly_chart(px.bar(people_df.head(10), x="achievement_count", y="person_name", orientation="h"), use_container_width=True)

with tab_report:
    st.subheader("PPT Builder")
    st.caption("The MVP uses a fixed slide template. AI narrative can be added later as a controlled, editable draft layer.")

    narrative = build_rule_based_narrative(summary, period_label)
    st.text_area("Executive summary draft", narrative, height=140)

    slide_plan = pd.DataFrame([
        {"Slide": 1, "Name": "Title Page", "Content": "Report title, period, key metrics"},
        {"Slide": 2, "Name": "Executive Summary", "Content": "Summary narrative and key highlights"},
        {"Slide": 3, "Name": "Achievement Overview Dashboard", "Content": "Metrics + activity/category charts"},
        {"Slide": 4, "Name": "Project / Study Highlights", "Content": "Top study-level details"},
        {"Slide": 5, "Name": "People Contribution View", "Content": "Contributor summary"},
        {"Slide": 6, "Name": "Appendix Detail Table", "Content": "Filtered achievement details"},
    ])
    st.dataframe(slide_plan, use_container_width=True, hide_index=True)

    ppt_bytes = generate_achievement_ppt(filtered, report_title=report_title, period_label=period_label)
    st.download_button(
        "Generate and download PPTX",
        data=ppt_bytes,
        file_name="deptflow_achievement_report.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )

with st.expander("Data quality checks"):
    checks = {
        "Rows loaded": len(df),
        "Rows after filters": len(filtered),
        "Missing delivery date": int(df["delivery_date"].isna().sum()),
        "Missing Study ID": int(df["study_id"].eq("").sum()),
        "Missing comments": int(df["comments"].eq("").sum()),
        "Missing submitted by": int(df["submitted_by"].eq("").sum()),
    }
    st.json(checks)
