from __future__ import annotations

import datetime as dt
import os
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from src.ai_utils import generate_narrative, get_provider_status, load_env_file, normalize_provider
from src.data_utils import (
    PRESENTATION_COLUMNS,
    build_people_summary,
    filter_achievements,
    normalize_achievements,
    read_achievement_csv,
)
from src.metrics_utils import (
    append_report_history,
    build_filter_trace,
    build_report_history_record,
    calculate_report_metrics,
    compact_metrics_for_ai,
    load_report_history,
    trace_signature,
)
from src.ppt_utils import FIXED_TEMPLATE_SLIDES, TEMPLATE_VERSION, generate_achievement_ppt

APP_DIR = Path(__file__).parent
SAMPLE_CSV = APP_DIR / "sample_data" / "2026-CDO China Achievements sample.csv"
ENV_PATH = APP_DIR / ".env"
REPORT_HISTORY_CSV = APP_DIR / "outputs" / "report_history.csv"

load_env_file(ENV_PATH)

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
    report_title = st.text_input("Report title", os.getenv("DEFAULT_REPORT_TITLE", "CDO China Achievement Report"))

try:
    if uploaded is not None:
        source = uploaded
        source_label = f"uploaded:{uploaded.name}"
    elif use_sample and SAMPLE_CSV.exists():
        source = SAMPLE_CSV
        source_label = "bundled_sample"
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
metrics = calculate_report_metrics(filtered)
summary = metrics["summary"]
period_label = f"{start_date} to {end_date}"
filter_trace = build_filter_trace(
    source_label=source_label,
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

narrative_signature = f"{trace_signature(filter_trace, metrics)}|report_title={report_title}"
if st.session_state.get("narrative_signature") != narrative_signature:
    result = generate_narrative(
        metrics=metrics,
        period_label=period_label,
        provider="rule_based",
        report_title=report_title,
        env_path=ENV_PATH,
    )
    st.session_state["narrative_text"] = result.text
    st.session_state["narrative_provider"] = result.provider
    st.session_state["narrative_model"] = result.model
    st.session_state["narrative_warning"] = result.warning
    st.session_state["narrative_signature"] = narrative_signature
    st.session_state.pop("ppt_bytes", None)
    st.session_state.pop("ppt_generated_at", None)
    st.session_state.pop("ppt_file_name", None)
    st.session_state.pop("ppt_signature", None)
    st.session_state.pop("ppt_narrative_text", None)

kpi1, kpi2, kpi3, kpi4 = st.columns(4)
kpi1.metric("Achievements", summary["total_achievements"])
kpi2.metric("Studies / Workstreams", summary["unique_studies"])
kpi3.metric("Contributors", summary["unique_contributors"])
kpi4.metric("CF-related Count", summary["cf_total"])

st.divider()

tab_overview, tab_detail, tab_people, tab_ai, tab_report = st.tabs([
    "Overview Dashboard",
    "Achievement Review Table",
    "People View",
    "AI Narrative Assistant",
    "PPT Builder",
])

with tab_overview:
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Activity Type Breakdown")
        activity_df = pd.DataFrame(metrics["activity_breakdown"].items(), columns=["Activity Type", "Count"])
        if not activity_df.empty:
            st.plotly_chart(px.bar(activity_df, x="Count", y="Activity Type", orientation="h"), use_container_width=True)
        else:
            st.info("No data for selected filters.")
    with c2:
        st.subheader("Slide Category Breakdown")
        category_df = pd.DataFrame(metrics["category_breakdown"].items(), columns=["Slide Category", "Count"])
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

with tab_ai:
    st.subheader("AI Narrative Assistant")
    st.caption("AI only drafts editable text. It receives application-calculated metrics, not raw authority to change data.")

    provider_status = get_provider_status(ENV_PATH)
    provider_codes = ["rule_based", "openai", "gemini"]
    default_provider = normalize_provider(os.getenv("AI_PROVIDER", "rule_based"))
    default_index = provider_codes.index(default_provider) if default_provider in provider_codes else 0

    def _provider_label(code: str) -> str:
        status = provider_status[code]
        availability = "available" if status["available"] else "no API key"
        return f"{status['label']} ({availability})"

    selected_provider = st.selectbox(
        "Narrative provider",
        provider_codes,
        index=default_index,
        format_func=_provider_label,
    )

    if st.button("Generate narrative from calculated metrics", type="primary"):
        result = generate_narrative(
            metrics=metrics,
            period_label=period_label,
            provider=selected_provider,
            report_title=report_title,
            env_path=ENV_PATH,
        )
        st.session_state["narrative_text"] = result.text
        st.session_state["narrative_provider"] = result.provider
        st.session_state["narrative_model"] = result.model
        st.session_state["narrative_warning"] = result.warning

    st.text_area("Editable narrative used in PPT", key="narrative_text", height=220)

    current_provider = st.session_state.get("narrative_provider", "rule_based")
    current_model = st.session_state.get("narrative_model", "rules-v1")
    st.caption(f"Current narrative source: {current_provider} / {current_model}")
    if st.session_state.get("narrative_warning"):
        st.warning(st.session_state["narrative_warning"])

    with st.expander("Calculated metrics sent to AI"):
        st.json(compact_metrics_for_ai(metrics))

with tab_report:
    st.subheader("PPT Builder")
    st.caption("Fixed-template PPTX uses filtered achievement data only and the editable narrative from the AI tab.")

    slide_plan = pd.DataFrame(FIXED_TEMPLATE_SLIDES)
    st.dataframe(slide_plan, use_container_width=True, hide_index=True)

    with st.expander("Traceability metadata for this build"):
        st.json({
            "template_version": TEMPLATE_VERSION,
            "ai_provider": st.session_state.get("narrative_provider", "rule_based"),
            "ai_model": st.session_state.get("narrative_model", "rules-v1"),
            "filters": filter_trace,
        })

    if st.button("Build PPTX from current filters and narrative", type="primary"):
        generated_at = dt.datetime.now(dt.timezone.utc).astimezone().isoformat(timespec="seconds")
        ai_provider = st.session_state.get("narrative_provider", "rule_based")
        ai_model = st.session_state.get("narrative_model", "rules-v1")
        narrative_text = st.session_state.get("narrative_text", "")

        ppt_bytes = generate_achievement_ppt(
            filtered,
            report_title=report_title,
            period_label=period_label,
            narrative_text=narrative_text,
            ai_provider=ai_provider,
            ai_model=ai_model,
            generated_at=generated_at,
            filters=filter_trace,
            metrics=metrics,
        )
        history_record = build_report_history_record(
            generated_at=generated_at,
            report_title=report_title,
            template_version=TEMPLATE_VERSION,
            ai_provider=ai_provider,
            ai_model=ai_model,
            rows_loaded=len(df),
            rows_filtered=len(filtered),
            period_label=period_label,
            filter_trace=filter_trace,
            metrics=metrics,
            narrative_text=narrative_text,
        )
        append_report_history(REPORT_HISTORY_CSV, history_record)

        st.session_state["ppt_bytes"] = ppt_bytes
        st.session_state["ppt_generated_at"] = generated_at
        st.session_state["ppt_file_name"] = "deptflow_achievement_report.pptx"
        st.session_state["ppt_signature"] = narrative_signature
        st.session_state["ppt_narrative_text"] = narrative_text
        st.success(f"PPTX built and history recorded at {generated_at}.")

    ppt_is_current = (
        st.session_state.get("ppt_bytes")
        and st.session_state.get("ppt_signature") == narrative_signature
        and st.session_state.get("ppt_narrative_text") == st.session_state.get("narrative_text", "")
    )
    if ppt_is_current:
        st.download_button(
            "Download PPTX",
            data=st.session_state["ppt_bytes"],
            file_name=st.session_state.get("ppt_file_name", "deptflow_achievement_report.pptx"),
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
    elif st.session_state.get("ppt_bytes"):
        st.warning("The existing PPTX build is stale because filters, title, or narrative changed. Rebuild before downloading.")
    else:
        st.info("Build the PPTX first, then download it.")

    with st.expander("Report history"):
        history = load_report_history(REPORT_HISTORY_CSV)
        st.dataframe(history.tail(10), use_container_width=True, hide_index=True)

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
