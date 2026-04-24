from __future__ import annotations

import datetime as dt
import os
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st
import streamlit.components.v1 as components

from src.ai_utils import generate_narrative, get_provider_status, load_env_file, normalize_provider
from src.data_utils import (
    PRESENTATION_COLUMNS,
    build_people_summary,
    filter_achievements,
    normalize_achievements,
    read_achievement_csv,
)
from src.html_preview_utils import render_deck_preview_html, render_slide_preview_html
from src.metrics_utils import (
    append_report_history,
    build_filter_trace,
    build_report_history_record,
    calculate_report_metrics,
    load_report_history,
    trace_signature,
)
from src.ppt_utils import FIXED_TEMPLATE_SLIDES, TEMPLATE_VERSION, generate_achievement_ppt
from src.slide_chat_utils import propose_slide_patch
from src.slide_spec_utils import (
    DATA_SOURCE_FIELDS,
    STYLE_PRESETS,
    apply_slide_patch,
    build_default_slide_spec,
    get_slide,
    slide_spec_signature,
    sync_slide_spec_context,
)

APP_DIR = Path(__file__).parent
SAMPLE_CSV = APP_DIR / "sample_data" / "2026-CDO China Achievements sample.csv"
ENV_PATH = APP_DIR / ".env"
REPORT_HISTORY_CSV = APP_DIR / "outputs" / "report_history.csv"

load_env_file(ENV_PATH)


def _describe_patch(patch: dict) -> pd.DataFrame:
    rows = []
    if patch.get("style_preset"):
        rows.append({"Scope": "Deck", "Change": "Style preset", "Value": patch["style_preset"]})
    for slide_change in patch.get("slides", []):
        slide_id = slide_change.get("id", "")
        for key, value in slide_change.items():
            if key == "id":
                continue
            rows.append({"Scope": f"Slide {slide_id}", "Change": key, "Value": str(value)})
    return pd.DataFrame(rows, columns=["Scope", "Change", "Value"])

st.set_page_config(
    page_title="DeptFlow Achievement Reporter",
    page_icon="📊",
    layout="wide",
)

st.title("DeptFlow Achievement Reporter")
st.caption("Filter achievements · Preview editable slide draft in HTML · Generate fixed-template PPTX")

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

default_result = generate_narrative(
    metrics=metrics,
    period_label=period_label,
    provider="rule_based",
    report_title=report_title,
    env_path=ENV_PATH,
)
context_signature = f"{trace_signature(filter_trace, metrics)}|report_title={report_title}"

if st.session_state.get("slide_context_signature") != context_signature:
    style_preset = st.session_state.get("slide_spec", {}).get("style_preset", os.getenv("SLIDE_STYLE_PRESET", "atlas"))
    st.session_state["slide_spec"] = build_default_slide_spec(
        metrics,
        filtered,
        report_title=report_title,
        period_label=period_label,
        narrative_text=default_result.text,
        style_preset=style_preset,
    )
    st.session_state["slide_context_signature"] = context_signature
    st.session_state["narrative_provider"] = default_result.provider
    st.session_state["narrative_model"] = default_result.model
    st.session_state["narrative_warning"] = default_result.warning
    for key in ["ppt_bytes", "ppt_generated_at", "ppt_file_name", "ppt_signature", "pending_slide_patch", "pending_slide_patch_warning"]:
        st.session_state.pop(key, None)
else:
    st.session_state["slide_spec"] = sync_slide_spec_context(
        st.session_state["slide_spec"],
        report_title=report_title,
        period_label=period_label,
        default_narrative=default_result.text,
    )

slide_spec = st.session_state["slide_spec"]
if "preview_slide_id" not in st.session_state:
    st.session_state["preview_slide_id"] = 1
st.session_state["preview_slide_id"] = max(1, min(8, int(st.session_state["preview_slide_id"])))
data_context = {
    "filtered_df": filtered,
    "metrics": metrics,
}

kpi1, kpi2, kpi3, kpi4 = st.columns(4)
kpi1.metric("Achievements", summary["total_achievements"])
kpi2.metric("Studies / Workstreams", summary["unique_studies"])
kpi3.metric("Contributors", summary["unique_contributors"])
kpi4.metric("CF-related Count", summary["cf_total"])

st.divider()

tab_overview, tab_detail, tab_people, tab_studio, tab_report = st.tabs([
    "Overview Dashboard",
    "Achievement Review Table",
    "People View",
    "AI Slide Studio",
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

    with st.expander("Data quality checks"):
        checks = pd.DataFrame([
            {"Check": "Rows loaded", "Value": len(df)},
            {"Check": "Rows after filters", "Value": len(filtered)},
            {"Check": "Missing delivery date", "Value": int(df["delivery_date"].isna().sum())},
            {"Check": "Missing Study ID", "Value": int(df["study_id"].eq("").sum())},
            {"Check": "Missing comments", "Value": int(df["comments"].eq("").sum())},
            {"Check": "Missing submitted by", "Value": int(df["submitted_by"].eq("").sum())},
        ])
        st.dataframe(checks, use_container_width=True, hide_index=True)

with tab_detail:
    st.subheader("Achievement Review Table")
    st.caption("This table is the human-review layer before formal PPT generation.")
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
    st.caption("Main activity is the most frequent activity type for each person within the current filtered data. If tied, the first mode returned by the count order is used.")
    people_view = people_df.rename(columns={
        "person_name": "Person",
        "achievement_count": "Achievement Count",
        "main_activity": "Most Frequent Activity",
    })
    st.dataframe(people_view, use_container_width=True, hide_index=True)
    if not people_df.empty:
        st.plotly_chart(px.bar(people_df.head(10), x="achievement_count", y="person_name", orientation="h"), use_container_width=True)

with tab_studio:
    st.subheader("AI Slide Studio")
    st.caption("HTML is a preview layer only. Confirmed PPTX is rebuilt as editable PowerPoint objects from the same SlideSpec.")

    provider_status = get_provider_status(ENV_PATH)
    provider_codes = ["rule_based", "openai", "gemini"]
    default_provider = normalize_provider(os.getenv("AI_PROVIDER", "rule_based"))
    default_index = provider_codes.index(default_provider) if default_provider in provider_codes else 0

    def _provider_label(code: str) -> str:
        status = provider_status[code]
        availability = "available" if status["available"] else "no API key"
        return f"{status['label']} ({availability})"

    control_col, preview_col = st.columns([0.9, 1.3], gap="large")

    with control_col:
        selected_provider = st.selectbox(
            "AI provider",
            provider_codes,
            index=default_index,
            format_func=_provider_label,
        )

        style_options = list(STYLE_PRESETS)
        current_style = slide_spec.get("style_preset", "atlas")
        next_style = st.selectbox(
            "Deck style preset",
            style_options,
            index=style_options.index(current_style) if current_style in style_options else 0,
            format_func=lambda code: STYLE_PRESETS[code]["label"],
        )
        if next_style != slide_spec.get("style_preset"):
            slide_spec["style_preset"] = next_style
            st.session_state["slide_spec"] = slide_spec
            st.session_state.pop("ppt_bytes", None)

        selected_slide_id = int(st.session_state["preview_slide_id"])
        selected_slide = get_slide(slide_spec, selected_slide_id)

        st.markdown(f"**Current Slide Controls: Slide {selected_slide_id}**")
        selected_slide["title"] = st.text_input("Slide title", value=selected_slide.get("title", ""), key=f"title_{selected_slide_id}")
        selected_slide["layout_variant"] = st.selectbox(
            "Layout variant",
            ["hero", "split", "dashboard", "chart_table", "table_full", "table_chart", "cards_table"],
            index=["hero", "split", "dashboard", "chart_table", "table_full", "table_chart", "cards_table"].index(selected_slide.get("layout_variant", "table_full")),
            key=f"layout_{selected_slide_id}",
        )
        selected_slide["chart_type"] = st.selectbox(
            "Chart type",
            ["none", "bar", "column"],
            index=["none", "bar", "column"].index(selected_slide.get("chart_type", "none")),
            key=f"chart_{selected_slide_id}",
        )
        selected_slide["top_n"] = st.number_input("Top N rows/items", min_value=1, max_value=25, value=int(selected_slide.get("top_n", 8)), step=1, key=f"topn_{selected_slide_id}")
        data_source = st.selectbox(
            "Data source",
            list(DATA_SOURCE_FIELDS),
            index=list(DATA_SOURCE_FIELDS).index(selected_slide.get("data_source", "none")),
            key=f"datasource_{selected_slide_id}",
        )
        selected_slide["data_source"] = data_source
        field_options = DATA_SOURCE_FIELDS.get(data_source, [])
        selected_slide["fields"] = st.multiselect(
            "Fields shown",
            field_options,
            default=[field for field in selected_slide.get("fields", []) if field in field_options],
            key=f"fields_{selected_slide_id}",
        )
        if selected_slide_id == 2:
            selected_slide["narrative"] = st.text_area("Editable executive summary", value=selected_slide.get("narrative", ""), height=170)

        st.session_state["slide_spec"] = slide_spec

        st.markdown("**Ask AI to adjust the draft**")
        user_message = st.text_area(
            "Request",
            placeholder="Example: slide 5 only show top 5 Study ID rows and include TA, High Impact, CF Total. Use a cleaner table layout.",
            height=120,
        )
        if st.button("Generate controlled patch", type="primary"):
            result = propose_slide_patch(
                user_message=user_message,
                slide_spec=slide_spec,
                metrics=metrics,
                provider=selected_provider,
                env_path=ENV_PATH,
            )
            st.session_state["pending_slide_patch"] = result.patch
            st.session_state["pending_slide_patch_warning"] = result.warning
            st.session_state["narrative_provider"] = result.provider
            st.session_state["narrative_model"] = result.model

        if st.session_state.get("pending_slide_patch_warning"):
            st.warning(st.session_state["pending_slide_patch_warning"])
        if st.session_state.get("pending_slide_patch"):
            st.markdown("**Proposed changes**")
            st.dataframe(_describe_patch(st.session_state["pending_slide_patch"]), use_container_width=True, hide_index=True)
            apply_col, discard_col = st.columns(2)
            with apply_col:
                if st.button("Apply patch"):
                    try:
                        st.session_state["slide_spec"] = apply_slide_patch(slide_spec, st.session_state["pending_slide_patch"])
                        st.session_state.pop("pending_slide_patch", None)
                        st.session_state.pop("pending_slide_patch_warning", None)
                        st.session_state.pop("ppt_bytes", None)
                        st.rerun()
                    except Exception as exc:
                        st.error(f"Patch rejected: {exc}")
            with discard_col:
                if st.button("Discard patch"):
                    st.session_state.pop("pending_slide_patch", None)
                    st.session_state.pop("pending_slide_patch_warning", None)
                    st.rerun()

        with st.expander("Audit trail summary"):
            st.dataframe(
                pd.DataFrame([
                    {"Item": "AI provider", "Value": selected_provider},
                    {"Item": "Style preset", "Value": slide_spec.get("style_preset", "atlas")},
                    {"Item": "Editing slide", "Value": selected_slide_id},
                    {"Item": "Data source", "Value": selected_slide.get("data_source", "")},
                    {"Item": "Fields", "Value": ", ".join(selected_slide.get("fields", []))},
                ]),
                use_container_width=True,
                hide_index=True,
            )

    with preview_col:
        st.markdown("**HTML Slide Preview**")
        nav_left, nav_mid, nav_right = st.columns([1, 2, 1])
        with nav_left:
            if st.button("‹ Previous", disabled=selected_slide_id <= 1, use_container_width=True):
                st.session_state["preview_slide_id"] = selected_slide_id - 1
                st.rerun()
        with nav_mid:
            st.markdown(f"<div style='text-align:center;font-weight:700;padding:.5rem 0;'>Slide {selected_slide_id} of 8</div>", unsafe_allow_html=True)
        with nav_right:
            if st.button("Next ›", disabled=selected_slide_id >= 8, use_container_width=True):
                st.session_state["preview_slide_id"] = selected_slide_id + 1
                st.rerun()

        view_mode = st.radio("Preview mode", ["Selected slide", "Full deck thumbnails"], horizontal=True, label_visibility="collapsed")
        if view_mode == "Selected slide":
            components.html(render_slide_preview_html(slide_spec, selected_slide_id, data_context), height=650, scrolling=False)
        else:
            components.html(render_deck_preview_html(slide_spec, data_context), height=1100, scrolling=True)

with tab_report:
    st.subheader("PPT Builder")
    st.caption("PPTX uses the current SlideSpec and filtered data. HTML preview is not screenshot into PPT.")

    slide_plan = pd.DataFrame(FIXED_TEMPLATE_SLIDES)
    st.dataframe(slide_plan, use_container_width=True, hide_index=True)

    current_slide_signature = slide_spec_signature(st.session_state["slide_spec"])
    build_signature = f"{context_signature}|slide_spec={current_slide_signature}"

    if st.button("Build editable PPTX from current SlideSpec", type="primary"):
        generated_at = dt.datetime.now(dt.timezone.utc).astimezone().isoformat(timespec="seconds")
        ai_provider = st.session_state.get("narrative_provider", "rule_based")
        ai_model = st.session_state.get("narrative_model", "rules-v1")
        slide_2 = get_slide(st.session_state["slide_spec"], 2)
        narrative_text = slide_2.get("narrative", "")

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
            slide_spec=st.session_state["slide_spec"],
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
            slide_spec=st.session_state["slide_spec"],
            style_preset=st.session_state["slide_spec"].get("style_preset", "atlas"),
            preview_generated=True,
            preview_renderer="html",
        )
        append_report_history(REPORT_HISTORY_CSV, history_record)

        st.session_state["ppt_bytes"] = ppt_bytes
        st.session_state["ppt_generated_at"] = generated_at
        st.session_state["ppt_file_name"] = "deptflow_achievement_report.pptx"
        st.session_state["ppt_signature"] = build_signature
        st.success(f"PPTX built and history recorded at {generated_at}.")

    ppt_is_current = st.session_state.get("ppt_bytes") and st.session_state.get("ppt_signature") == build_signature
    if ppt_is_current:
        st.download_button(
            "Download PPTX",
            data=st.session_state["ppt_bytes"],
            file_name=st.session_state.get("ppt_file_name", "deptflow_achievement_report.pptx"),
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
    elif st.session_state.get("ppt_bytes"):
        st.warning("The existing PPTX build is stale because filters, title, or SlideSpec changed. Rebuild before downloading.")
    else:
        st.info("Build the PPTX first, then download it.")

    with st.expander("Report history"):
        history = load_report_history(REPORT_HISTORY_CSV)
        visible_history_cols = [
            col for col in [
                "generated_at",
                "report_title",
                "template_version",
                "ai_provider",
                "ai_model",
                "style_preset",
                "rows_loaded",
                "rows_filtered",
                "period_label",
                "preview_renderer",
            ] if col in history.columns
        ]
        st.dataframe(history[visible_history_cols].tail(10), use_container_width=True, hide_index=True)
