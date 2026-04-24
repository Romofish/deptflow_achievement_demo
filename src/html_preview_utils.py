from __future__ import annotations

import html
from typing import Any

import pandas as pd

from .slide_spec_utils import STYLE_PRESETS, get_data_frame_for_slide, get_slide, normalize_style_preset


def _esc(value: Any) -> str:
    return html.escape("" if value is None else str(value))


def _truncate(value: Any, max_chars: int = 96) -> str:
    text = "" if value is None or pd.isna(value) else str(value).replace("\n", " ").strip()
    return text if len(text) <= max_chars else text[: max_chars - 1] + "..."


def _metric_cards(metrics: dict[str, Any]) -> str:
    summary = metrics.get("summary", {})
    quality = metrics.get("quality_complexity", {})
    cards = [
        ("Achievements", summary.get("total_achievements", 0)),
        ("Studies / Workstreams", summary.get("unique_studies", 0)),
        ("Contributors", summary.get("unique_contributors", 0)),
        ("CF-related Count", summary.get("cf_total", 0)),
        ("High Impact", quality.get("high_impact_count", 0)),
        ("Quality / Inspection", quality.get("quality_inspection_count", 0)),
    ]
    return "".join(f"<div class='metric'><b>{_esc(value)}</b><span>{_esc(label)}</span></div>" for label, value in cards)


def _table_html(df: pd.DataFrame, compact: bool = False) -> str:
    if df.empty:
        return "<div class='empty'>No records found for selected filters.</div>"

    header = "".join(f"<th>{_esc(col)}</th>" for col in df.columns)
    rows = []
    max_chars = 58 if compact else 96
    for _, row in df.iterrows():
        cells = "".join(f"<td>{_esc(_truncate(row[col], max_chars=max_chars))}</td>" for col in df.columns)
        rows.append(f"<tr>{cells}</tr>")
    return f"<table><thead><tr>{header}</tr></thead><tbody>{''.join(rows)}</tbody></table>"


def _chart_html(df: pd.DataFrame, chart_type: str) -> str:
    if df.empty or df.shape[1] < 2:
        return "<div class='chart empty'>No chart data.</div>"
    label_col, value_col = df.columns[0], df.columns[1]
    values = pd.to_numeric(df[value_col], errors="coerce").fillna(0)
    max_value = max(float(values.max()), 1.0)
    bars = []
    for idx, (_, row) in enumerate(df.iterrows()):
        value = float(pd.to_numeric(row[value_col], errors="coerce") or 0)
        width = max(4, int((value / max_value) * 100))
        if chart_type == "column":
            bars.append(
                f"<div class='colbar' style='height:{width}%' title='{_esc(row[label_col])}: {_esc(value)}'>"
                f"<span>{_esc(int(value))}</span></div>"
            )
        else:
            bars.append(
                f"<div class='barrow'><label>{_esc(_truncate(row[label_col], 34))}</label>"
                f"<div class='bartrack'><div class='barfill' style='width:{width}%'></div></div>"
                f"<strong>{_esc(int(value))}</strong></div>"
            )
    if chart_type == "column":
        labels = "".join(f"<span>{_esc(_truncate(row[label_col], 10))}</span>" for _, row in df.iterrows())
        return f"<div class='columns'>{''.join(bars)}</div><div class='column-labels'>{labels}</div>"
    return f"<div class='bars'>{''.join(bars)}</div>"


def _slide_body(slide: dict[str, Any], df: pd.DataFrame, metrics: dict[str, Any]) -> str:
    kind = slide.get("kind")
    table_df = get_data_frame_for_slide(slide, df, metrics)
    chart_type = slide.get("chart_type", "none")

    if kind == "title":
        return (
            "<section class='hero-grid'>"
            f"<div class='hero-copy'><p>{_esc(slide.get('narrative', ''))}</p>"
            "<p class='muted'>Template locked to 8 editable slides. Data comes from current filters.</p></div>"
            f"<div class='metric-grid'>{_metric_cards(metrics)}</div>"
            "</section>"
        )
    if kind == "narrative":
        return (
            "<section class='split-grid'>"
            f"<div class='narrative'>{_esc(slide.get('narrative', '')).replace(chr(10), '<br><br>')}</div>"
            f"<div class='metric-stack'>{_metric_cards(metrics)}</div>"
            "</section>"
        )
    if kind == "dashboard":
        return (
            "<section class='dashboard-grid'>"
            f"<div class='metric-grid wide'>{_metric_cards(metrics)}</div>"
            f"<div class='panel'>{_chart_html(table_df, chart_type)}</div>"
            f"<div class='panel'>{_chart_html(pd.DataFrame(metrics.get('impact_breakdown', {}).items(), columns=['Impact', 'Count']), 'column')}</div>"
            "</section>"
        )
    if slide.get("layout_variant") in {"chart_table", "table_chart"} and chart_type != "none":
        return (
            "<section class='chart-table-grid'>"
            f"<div class='panel chart-panel'>{_chart_html(table_df, chart_type)}</div>"
            f"<div class='panel table-panel'>{_table_html(table_df, compact=True)}</div>"
            "</section>"
        )
    if kind == "quality":
        return (
            "<section class='quality-grid'>"
            f"<div class='metric-grid'>{_metric_cards(metrics)}</div>"
            f"<div class='panel table-wide'>{_table_html(table_df, compact=True)}</div>"
            "</section>"
        )
    return f"<section class='panel table-wide'>{_table_html(table_df, compact=True)}</section>"


def _style_block(style_preset: str) -> str:
    style = STYLE_PRESETS[normalize_style_preset(style_preset)]
    return f"""
    <style>
      :root {{
        --paper: {style['paper']};
        --surface: {style['surface']};
        --ink: {style['ink']};
        --muted: {style['muted']};
        --accent: {style['accent']};
        --accent-2: {style['accent_2']};
        --accent-3: {style['accent_3']};
      }}
      * {{ box-sizing: border-box; }}
      body {{ margin: 0; background: #d9dee7; font-family: 'Aptos Display', 'Segoe UI', sans-serif; }}
      .deck-preview {{ display: grid; gap: 18px; padding: 10px; }}
      .slide {{
        width: 100%;
        aspect-ratio: 16 / 9;
        background:
          radial-gradient(circle at 8% 15%, color-mix(in srgb, var(--accent) 18%, transparent), transparent 32%),
          linear-gradient(135deg, var(--surface), var(--paper));
        color: var(--ink);
        border-radius: 22px;
        overflow: hidden;
        position: relative;
        box-shadow: 0 22px 70px rgba(17, 24, 39, .18);
        padding: 4.2% 4.6% 3.4%;
      }}
      .slide::before {{
        content: "";
        position: absolute;
        inset: 0 0 auto 0;
        height: 9px;
        background: linear-gradient(90deg, var(--accent), var(--accent-2), var(--accent-3));
      }}
      .slide-header {{ display: flex; justify-content: space-between; gap: 24px; align-items: start; margin-bottom: 2.4%; }}
      h1 {{ margin: 0; font-size: clamp(22px, 3vw, 43px); letter-spacing: -.04em; line-height: .96; }}
      .subtitle {{ color: var(--muted); font-size: clamp(10px, 1vw, 15px); margin-top: 8px; }}
      .chip {{ border: 1px solid color-mix(in srgb, var(--accent) 28%, transparent); border-radius: 999px; padding: 8px 12px; color: var(--accent); font-size: 12px; white-space: nowrap; background: color-mix(in srgb, var(--surface) 78%, transparent); }}
      .hero-grid, .split-grid, .dashboard-grid, .chart-table-grid, .quality-grid {{ display: grid; gap: 22px; height: 78%; }}
      .hero-grid {{ grid-template-columns: 1.05fr 1.2fr; align-items: end; }}
      .split-grid {{ grid-template-columns: 1.35fr .85fr; }}
      .dashboard-grid {{ grid-template-columns: 1fr 1fr; grid-template-rows: auto 1fr; }}
      .dashboard-grid .wide {{ grid-column: 1 / -1; }}
      .chart-table-grid {{ grid-template-columns: 1.2fr .9fr; }}
      .quality-grid {{ grid-template-rows: auto 1fr; }}
      .hero-copy p {{ font-size: clamp(16px, 1.7vw, 27px); line-height: 1.14; margin: 0 0 16px; max-width: 700px; }}
      .muted, .empty {{ color: var(--muted); }}
      .narrative {{ font-size: clamp(13px, 1.45vw, 23px); line-height: 1.38; background: color-mix(in srgb, var(--surface) 78%, transparent); border: 1px solid rgba(100,100,100,.12); border-radius: 22px; padding: 24px; }}
      .panel {{ background: color-mix(in srgb, var(--surface) 86%, transparent); border: 1px solid rgba(100,100,100,.13); border-radius: 22px; padding: 18px; overflow: hidden; min-height: 0; }}
      .metric-grid {{ display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 14px; }}
      .metric-grid.wide {{ grid-template-columns: repeat(6, minmax(0, 1fr)); }}
      .metric-stack {{ display: grid; gap: 13px; }}
      .metric {{ border: 1px solid color-mix(in srgb, var(--accent) 18%, transparent); border-radius: 18px; padding: 15px; background: color-mix(in srgb, var(--surface) 82%, transparent); }}
      .metric b {{ display: block; font-size: clamp(21px, 2.4vw, 34px); color: var(--accent); line-height: .9; }}
      .metric span {{ display: block; margin-top: 9px; font-size: 11px; color: var(--muted); text-transform: uppercase; letter-spacing: .08em; }}
      table {{ width: 100%; border-collapse: collapse; font-size: clamp(8px, .9vw, 13px); }}
      th {{ background: var(--ink); color: var(--surface); text-align: left; padding: 9px 10px; font-size: .9em; }}
      td {{ padding: 8px 10px; border-bottom: 1px solid rgba(90,90,90,.12); vertical-align: top; }}
      tr:nth-child(even) td {{ background: color-mix(in srgb, var(--accent) 5%, transparent); }}
      .bars {{ display: grid; gap: 11px; height: 100%; align-content: center; }}
      .barrow {{ display: grid; grid-template-columns: minmax(90px, 1fr) 2fr 40px; gap: 10px; align-items: center; font-size: 12px; }}
      .bartrack {{ height: 12px; border-radius: 999px; background: color-mix(in srgb, var(--muted) 16%, transparent); overflow: hidden; }}
      .barfill {{ height: 100%; background: linear-gradient(90deg, var(--accent), var(--accent-2)); border-radius: inherit; }}
      .columns {{ display: flex; align-items: end; justify-content: center; gap: 16px; height: 78%; padding: 18px 16px 4px; }}
      .colbar {{ flex: 1; min-height: 8%; border-radius: 14px 14px 4px 4px; background: linear-gradient(180deg, var(--accent-3), var(--accent)); display: flex; align-items: start; justify-content: center; color: white; font-size: 11px; padding-top: 7px; }}
      .column-labels {{ display: flex; gap: 16px; color: var(--muted); font-size: 10px; justify-content: center; }}
      .column-labels span {{ flex: 1; text-align: center; }}
      .slide-footer {{ position: absolute; left: 4.6%; right: 4.6%; bottom: 2.2%; color: var(--muted); font-size: 10px; display: flex; justify-content: space-between; }}
      .thumb-grid {{ display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 16px; }}
      .thumb-grid .slide {{ border-radius: 14px; padding: 4%; }}
    </style>
    """


def render_slide_preview_html(slide_spec: dict[str, Any], slide_id: int, data_context: dict[str, Any]) -> str:
    slide = get_slide(slide_spec, slide_id)
    style_preset = slide_spec.get("style_preset", "atlas")
    df = data_context["filtered_df"]
    metrics = data_context["metrics"]
    body = _slide_body(slide, df, metrics)
    return f"""
    {_style_block(style_preset)}
    <div class="deck-preview">
      <article class="slide">
        <header class="slide-header">
          <div>
            <h1>{_esc(slide.get('title', ''))}</h1>
            <div class="subtitle">{_esc(slide.get('subtitle', ''))}</div>
          </div>
          <div class="chip">Slide {slide_id} / 8</div>
        </header>
        {body}
        <footer class="slide-footer"><span>{_esc(slide_spec.get('version', 'slide-spec-v1'))}</span><span>{_esc(STYLE_PRESETS[normalize_style_preset(style_preset)]['label'])}</span></footer>
      </article>
    </div>
    """


def render_deck_preview_html(slide_spec: dict[str, Any], data_context: dict[str, Any]) -> str:
    slides = []
    for slide in slide_spec.get("slides", []):
        slide_id = int(slide["id"])
        df = data_context["filtered_df"]
        metrics = data_context["metrics"]
        slides.append(
            "<article class='slide'>"
            "<header class='slide-header'>"
            f"<div><h1>{_esc(slide.get('title', ''))}</h1><div class='subtitle'>{_esc(slide.get('subtitle', ''))}</div></div>"
            f"<div class='chip'>Slide {slide_id}</div>"
            "</header>"
            f"{_slide_body(slide, df, metrics)}"
            "</article>"
        )
    return f"{_style_block(slide_spec.get('style_preset', 'atlas'))}<div class='thumb-grid'>{''.join(slides)}</div>"
