"""Streamlit entry point for the Survey Analysis Engine."""

from __future__ import annotations

import os
import sys
import tempfile
import traceback
from pathlib import Path
from typing import Any

try:
    import streamlit as st
except ModuleNotFoundError:  # pragma: no cover - local import smoke test fallback.
    st = None  # type: ignore[assignment]


PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.ui_constants import (
    APP_TAGLINE,
    APP_TITLE,
    EMPTY_NO_CROSS_CUTS,
    EMPTY_NO_RESULTS,
    PIPELINE_STAGES,
    SECTION_CROSS_CUTS,
    SECTION_DOWNLOADS,
    SECTION_GLOBAL_FILTER,
    SECTION_RESULTS,
    SECTION_UPLOAD,
    STATUS_GLOBAL_FILTER_ACTIVE,
    STATUS_GLOBAL_FILTER_INACTIVE,
    TOOLTIP_BREAKDOWN,
    TOOLTIP_CROSS_CUT_SUGGESTIONS,
    TOOLTIP_GLOBAL_FILTER,
    TOOLTIP_PER_QUESTION_FILTER,
    TOOLTIP_THREE_DOWNLOADS,
)


SESSION_DEFAULTS = {
    "decoded_df": None,
    "active_df": None,
    "global_filter_state": None,
    "global_filter_stats": None,
    "global_filter_rows": [],
    "results": [],
    "skips": [],
    "schema": None,
    "quality_report": None,
    "log": None,
    "output_path": None,
    "raw_data_path_label": None,
    "datamap_path_label": None,
    "load_report": None,
    "cross_cut_results": [],
    "cross_cut_skips": [],
    "cross_cut_suggestions": [],
    "cross_cut_only_bytes": None,
    "filtered_results": {},
    "filtered_workbook_bytes": None,
    "run_complete": False,
    "ss_search": "",
    "pending_global_filter": None,
    "pending_per_question_filter": {},
    "global_filter_error": None,
    "per_question_filter_errors": {},
}


# ---------------------------------------------------------------------------
# Infrastructure helpers
# ---------------------------------------------------------------------------


def _require_streamlit() -> Any:
    if st is None:
        raise RuntimeError(
            "Streamlit is not installed in this environment. "
            "Install project requirements before running the app."
        )
    return st


CROSS_CUT_ENGINE_VERSION = "day14.5"


# ---------------------------------------------------------------------------
# Visual theme + UI helpers (Day 16)
# ---------------------------------------------------------------------------


_THEME_CSS = """
<style>
html, body, [class*="css"], .stApp, .stMarkdown,
.stText, .stDataFrame, button, input, select,
textarea, label, p, div, span, h1, h2, h3 {
  font-family: Arial, Helvetica, sans-serif !important;
}
.stApp, .main .block-container { background-color: #FFFFFF !important; }
header[data-testid="stHeader"] {
  background: #FFFFFF !important;
  border-bottom: 3px solid #CC0000 !important;
  height: 52px !important;
}
#MainMenu, footer, .stDeployButton { display: none !important; }
[data-testid="stSidebar"] {
  background: #F8F8F8 !important;
  border-right: 1px solid #E0E0E0 !important;
}
[data-testid="stSidebar"] .stMarkdown p,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] .stTextInput label {
  font-size: 12px !important; color: #333333 !important;
}
[data-testid="stMetric"] {
  background: #FFFFFF !important;
  border: 1px solid #E0E0E0 !important;
  border-top: 3px solid #CC0000 !important;
  padding: 16px 20px !important;
  border-radius: 0 !important;
}
[data-testid="stMetricValue"] {
  font-family: Arial, Helvetica, sans-serif !important;
  font-size: 32px !important; font-weight: 700 !important;
  color: #0A0A0A !important;
}
[data-testid="stMetricLabel"] {
  font-size: 10px !important; font-weight: 700 !important;
  letter-spacing: 0.12em !important; text-transform: uppercase !important;
  color: #888888 !important;
}
.section-header-box {
  display: flex; align-items: center; gap: 12px;
  padding: 10px 0; border-bottom: 2px solid #E0E0E0;
  margin-bottom: 20px;
}
.section-num {
  width: 26px; height: 26px; background: #CC0000; color: white;
  display: flex; align-items: center; justify-content: center;
  font-size: 12px; font-weight: 700; flex-shrink: 0;
  font-family: Arial, Helvetica, sans-serif;
}
.section-name {
  font-size: 14px; font-weight: 700; text-transform: uppercase;
  letter-spacing: 0.08em; color: #0A0A0A;
}
.section-meta {
  margin-left: auto; font-size: 10px; color: #888888;
  font-family: Arial, Helvetica, sans-serif;
}
.stButton > button {
  border-radius: 0 !important;
  font-family: Arial, Helvetica, sans-serif !important;
  font-weight: 700 !important; font-size: 11px !important;
  letter-spacing: 0.1em !important; text-transform: uppercase !important;
  border: none !important; transition: background 0.15s !important;
}
.stButton > button[kind="primary"], .stButton > button:first-child {
  background: #CC0000 !important; color: white !important;
  padding: 10px 28px !important;
}
.stButton > button:hover { background: #990000 !important; color: white !important; }
.stButton > button[kind="secondary"] {
  background: #F0F0F0 !important; color: #333333 !important;
  border: 1px solid #CCCCCC !important;
}
[data-testid="stFileUploader"] {
  background: #FFFFFF !important;
  border: 1px dashed #CC0000 !important;
  border-radius: 0 !important; padding: 24px !important;
}
[data-testid="stFileUploader"]:hover { background: #FFF5F5 !important; }
[data-testid="stFileUploaderDropzoneInstructions"] { color: #CC0000 !important; }
.stAlert { border-radius: 0 !important; font-size: 12px !important; }
[data-testid="stInfo"] {
  background: #FFF5F5 !important;
  border: 1px solid rgba(204,0,0,0.3) !important;
  border-left: 4px solid #CC0000 !important;
  color: #0A0A0A !important;
}
[data-testid="stExpander"] {
  border: 1px solid #E0E0E0 !important; border-radius: 0 !important;
  border-left: 3px solid #E0E0E0 !important;
  margin-bottom: 4px !important; background: #FFFFFF !important;
}
[data-testid="stExpander"]:hover { border-left-color: #CC0000 !important; }
[data-testid="stExpander"] summary {
  font-weight: 600 !important; font-size: 13px !important;
  padding: 12px 16px !important; color: #0A0A0A !important;
}
[data-testid="stExpanderDetails"] {
  border-top: 1px solid #E0E0E0 !important; padding: 16px !important;
}
[data-testid="stSelectbox"] > div, [data-testid="stMultiSelect"] > div {
  border-radius: 0 !important;
}
.stSelectbox [data-baseweb="select"] div,
.stMultiSelect [data-baseweb="select"] div {
  border-radius: 0 !important; border-color: #CCCCCC !important;
  font-size: 12px !important;
}
.stTextInput input {
  border-radius: 0 !important; border-color: #CCCCCC !important;
  font-size: 12px !important;
}
[data-testid="stDataFrame"] {
  border: 1px solid #E0E0E0 !important; border-radius: 0 !important;
}
</style>
"""


_THEME_CSS_DAY18 = """
<style>
/* Day 18 — Fixed red header bar */
header[data-testid="stHeader"] {
  background: #CC0000 !important;
  border-bottom: none !important;
  height: 64px !important;
  z-index: 999 !important;
  position: fixed !important;
  top: 0 !important; left: 0 !important; right: 0 !important;
}
header[data-testid="stHeader"]::before {
  content: ''; position: absolute; inset: 0;
  background: #CC0000; z-index: -1;
}
header[data-testid="stHeader"] [data-testid="stToolbar"] {
  display: none !important;
}
.stApp { margin-top: 0 !important; }
.main .block-container {
  padding-top: 80px !important;
  max-width: 100% !important;
}
[data-testid="stSidebar"] {
  margin-top: 64px !important;
  height: calc(100vh - 64px) !important;
}
.custom-header {
  position: fixed; top: 0; left: 0; right: 0;
  height: 64px; background: #CC0000; z-index: 1000;
  display: flex; align-items: center;
  padding: 0 32px; color: white;
  font-family: Arial, Helvetica, sans-serif;
}
.custom-header-title {
  font-size: 20px; font-weight: 700;
  color: white; letter-spacing: 0.02em; margin-right: 24px;
}
.custom-header-tagline {
  font-size: 12px; color: rgba(255,255,255,0.85);
  font-weight: 400;
  border-left: 1px solid rgba(255,255,255,0.3);
  padding-left: 24px;
}

/* Day 18 — Sidebar question navigation buttons */
[data-testid="stSidebar"] .stButton > button {
  text-align: left !important;
  justify-content: flex-start !important;
  white-space: normal !important;
  height: auto !important;
  min-height: 44px !important;
  padding: 10px 12px !important;
  font-size: 12px !important;
  font-weight: 500 !important;
  line-height: 1.4 !important;
  letter-spacing: 0 !important;
  text-transform: none !important;
  border-radius: 0 !important;
  border-left: 3px solid transparent !important;
  background: white !important;
  color: #333 !important;
  border: 1px solid #E8E8E8 !important;
  margin-bottom: 2px !important;
  word-wrap: break-word !important;
  overflow-wrap: break-word !important;
}
[data-testid="stSidebar"] .stButton > button:hover {
  background: #FFF5F5 !important;
  border-left-color: #CC0000 !important;
  color: #0A0A0A !important;
}
[data-testid="stSidebar"] .stButton > button[kind="primary"] {
  background: #FFF0F0 !important;
  border-left: 3px solid #CC0000 !important;
  color: #CC0000 !important;
  font-weight: 700 !important;
}

/* Day 18 — Plotly chart styling */
.plotly, .plotly-chart { background: white !important; }
.plotly-chart [class*="modebar"] {
  background: rgba(255,255,255,0.8) !important;
  border: 1px solid #E0E0E0 !important;
}
.plotly-chart [class*="modebar-btn"] { color: #666 !important; }
</style>
"""

_CUSTOM_HEADER_HTML = """
<div class="custom-header">
  <div class="custom-header-title">Survey Analysis Engine</div>
  <div class="custom-header-tagline">
    Upload survey data, explore single &amp; cross cuts,
    export consultant-ready workbooks
  </div>
</div>
"""


def _inject_theme_css() -> None:
    app = _require_streamlit()
    app.markdown(_THEME_CSS, unsafe_allow_html=True)
    app.markdown(_THEME_CSS_DAY18, unsafe_allow_html=True)
    app.markdown(_CUSTOM_HEADER_HTML, unsafe_allow_html=True)


def _section_header(
    num: str, title: str, anchor: str | None = None, meta: str = ""
) -> None:
    app = _require_streamlit()
    anchor_html = f"<a name='{anchor}'></a>" if anchor else ""
    meta_html = f"<div class='section-meta'>{meta}</div>" if meta else ""
    app.markdown(
        f"{anchor_html}<div class='section-header-box'>"
        f"<div class='section-num'>{num}</div>"
        f"<div class='section-name'>{title}</div>"
        f"{meta_html}</div>",
        unsafe_allow_html=True,
    )


def _compute_outlier_flags(values: list) -> list:
    """Return 'high' / 'low' / '' per value using a guarded z-score rule.

    Never raises — wrapped in a defensive try/except so a malformed value
    list cannot break a render path.
    """
    try:
        import statistics

        flags = ["" for _ in values]
        numeric = [
            v for v in values
            if v is not None and isinstance(v, (int, float)) and v == v
        ]
        if len(numeric) < 4:
            return flags
        if max(numeric) < 10:
            return flags
        if len(set(numeric)) <= 1:
            return flags
        mean = statistics.mean(numeric)
        std = statistics.stdev(numeric)
        if std == 0:
            return flags
        val_range = max(numeric) - min(numeric)
        min_val = min(numeric)
        range_significant = (
            val_range > mean * 0.5
            and (
                min_val == 0
                or max(numeric) / max(min_val, 0.001) > 3
            )
        )
        for i, v in enumerate(values):
            if v is None or not isinstance(v, (int, float)):
                continue
            z = (v - mean) / std
            if z > 2.0 and v >= mean + val_range * 0.5:
                flags[i] = "high"
            elif z < -1.5 and range_significant and v < mean * 0.3:
                flags[i] = "low"
        return flags
    except Exception:
        return ["" for _ in values]


def _style_outliers(df: Any) -> Any:
    """Per-column outlier styling for st.dataframe Styler.apply(axis=None)."""
    import pandas as pd

    try:
        styled = pd.DataFrame("", index=df.index, columns=df.columns)
        for col in df.columns:
            try:
                vals = pd.to_numeric(df[col], errors="coerce").tolist()
                col_flags = _compute_outlier_flags(vals)
                for i, flag in enumerate(col_flags):
                    if flag == "high":
                        styled.iloc[i, styled.columns.get_loc(col)] = (
                            "background-color:#FFF0F0;color:#CC0000;"
                            "font-weight:bold;border-left:3px solid #CC0000"
                        )
                    elif flag == "low":
                        styled.iloc[i, styled.columns.get_loc(col)] = (
                            "background-color:#FFFBE6;color:#E65100;"
                            "font-weight:600"
                        )
            except Exception:
                continue
        return styled
    except Exception:
        return pd.DataFrame("", index=df.index, columns=df.columns)


def _styled_dataframe(df: Any, **kwargs: Any) -> None:
    """Render a dataframe with outlier styling, falling back gracefully."""
    app = _require_streamlit()
    try:
        app.dataframe(
            df.style.apply(_style_outliers, axis=None), **kwargs
        )
    except Exception:
        app.dataframe(df, **kwargs)


def _render_sc_table_html(
    distribution: dict,
    display_mode: str,
    valid_n: int,
    flags: list | None = None,
    key_suffix: str = "",
) -> None:
    """Branded HTML table for SingleSelect / MultiSelect distributions.

    Labels are HTML-escaped to prevent any XSS via maliciously crafted
    data-map files (defense in depth — labels are analyst-supplied).

    ``flags`` may be precomputed by the caller so that highlighting is
    identical across display-mode toggles. If omitted, flags are derived
    from the distribution counts.
    """
    import html as _html

    app = _require_streamlit()
    if flags is None:
        counts = [p.get("count", 0) for _, p in distribution.items()]
        flags = _compute_outlier_flags(counts)
    rows_html = ""
    for i, (code, payload) in enumerate(distribution.items()):
        label = _html.escape(str(payload.get("label", code)))
        count = payload.get("count", 0)
        rate = payload.get("rate")
        if rate is None:
            rate = (count / valid_n) if valid_n else 0
        pct = f"{rate * 100:.1f}%"
        flag = flags[i] if i < len(flags) else ""
        bar_width = max(0, min(160, int(rate * 160)))
        bar_html = (
            f'<div style="display:flex;align-items:center;gap:8px;">'
            f'<div style="width:160px;height:4px;background:#F0F0F0;'
            f'position:relative;flex-shrink:0;">'
            f'<div style="position:absolute;top:0;left:0;bottom:0;'
            f'width:{bar_width}px;background:#CC0000;"></div>'
            f"</div></div>"
        )
        if flag == "high":
            row_style = "background:#FFF0F0;border-left:3px solid #CC0000;"
            count_style = "font-weight:700;color:#CC0000;"
            flag_cell = "\u2B06"
        elif flag == "low":
            row_style = "background:#FFFBE6;"
            count_style = "font-weight:600;color:#E65100;"
            flag_cell = "\u2193"
        else:
            row_style = ""
            count_style = "color:#333;"
            flag_cell = ""
        td = "padding:7px 10px;border-bottom:1px solid #F5F5F5;"
        if display_mode == "Counts":
            row = (
                f'<tr style="{row_style}">'
                f'<td style="{td}font-size:12px;">{label}</td>'
                f'<td style="{td}{count_style}font-size:12px;text-align:right;">'
                f"{count:,}</td>"
                f'<td style="{td}">{bar_html}</td>'
                f'<td style="{td}font-size:11px;color:#CC0000;">{flag_cell}</td>'
                f"</tr>"
            )
        elif display_mode == "Counts + %":
            row = (
                f'<tr style="{row_style}">'
                f'<td style="{td}font-size:12px;">{label}</td>'
                f'<td style="{td}{count_style}font-size:12px;text-align:right;'
                f'width:60px;">{count:,}</td>'
                f'<td style="{td}font-size:12px;font-weight:700;color:#CC0000;'
                f'text-align:right;width:54px;">{pct}</td>'
                f'<td style="{td}">{bar_html}</td>'
                f'<td style="{td}font-size:11px;color:#CC0000;">{flag_cell}</td>'
                f"</tr>"
            )
        else:  # "% only"
            row = (
                f'<tr style="{row_style}">'
                f'<td style="{td}font-size:12px;">{label}</td>'
                f'<td style="{td}font-size:13px;font-weight:700;color:#CC0000;'
                f'text-align:right;">{pct}</td>'
                f'<td style="{td}">{bar_html}</td>'
                f'<td style="{td}font-size:11px;color:#CC0000;">{flag_cell}</td>'
                f"</tr>"
            )
        rows_html += row

    th = (
        "padding:8px 10px;font-size:10px;text-transform:uppercase;"
        "letter-spacing:0.1em;color:#888;"
    )
    if display_mode == "Counts":
        headers = (
            f"<th style='{th}text-align:left;'>Label</th>"
            f"<th style='{th}text-align:right;'>Count</th>"
            f"<th style='{th}'>Bar</th><th></th>"
        )
    elif display_mode == "Counts + %":
        headers = (
            f"<th style='{th}text-align:left;'>Label</th>"
            f"<th style='{th}text-align:right;'>Count</th>"
            f"<th style='{th}text-align:right;'>%</th>"
            f"<th style='{th}'>Bar</th><th></th>"
        )
    else:
        headers = (
            f"<th style='{th}text-align:left;'>Label</th>"
            f"<th style='{th}text-align:right;'>%</th>"
            f"<th style='{th}'>Bar</th><th></th>"
        )

    app.markdown(
        f'<table style="width:100%;border-collapse:collapse;'
        f'font-family:Arial,Helvetica,sans-serif;">'
        f'<thead><tr style="background:#F8F8F8;'
        f'border-bottom:2px solid #E0E0E0;">{headers}</tr></thead>'
        f"<tbody>{rows_html}</tbody></table>"
        f'<div style="font-size:10px;color:#888;margin-top:6px;'
        f'font-family:Arial,Helvetica,sans-serif;">'
        f"Valid N: {valid_n:,} &nbsp;\u00b7&nbsp; "
        f"\u2B06 high outlier &nbsp;\u00b7&nbsp; \u2193 low outlier"
        f"</div>",
        unsafe_allow_html=True,
    )

    # Secondary, fully-interactive table: sortable / copyable / CSV-downloadable.
    # The HTML table above remains the primary visual (bars cannot live inside
    # st.dataframe). We use a checkbox toggle (NOT st.expander) because this
    # helper is itself rendered inside an st.expander in the single-cut card,
    # and Streamlit forbids nested expanders.
    show_table = app.checkbox(
        "\U0001F4CB Show table view (sortable \u00b7 copyable \u00b7 downloadable)",
        key=f"sc_tableview_{key_suffix}" if key_suffix else None,
        value=False,
    )
    if show_table:
        rows_for_df = []
        for code, payload in distribution.items():
            count = payload.get("count", 0)
            rate = payload.get("rate")
            if rate is None:
                rate = (count / valid_n) if valid_n else 0
            rows_for_df.append(
                {
                    "Label": payload.get("label", str(code)),
                    "Count": count,
                    "%": round(rate * 100, 1),
                }
            )
        df_plain = pd.DataFrame(rows_for_df)
        try:
            app.dataframe(
                df_plain.style.apply(_style_outliers, axis=None),
                use_container_width=True,
                hide_index=True,
            )
        except Exception:
            app.dataframe(
                df_plain, use_container_width=True, hide_index=True
            )


# ---------------------------------------------------------------------------
# Day 18 — Plotly chart helpers
# ---------------------------------------------------------------------------


def _chart_text_label(text: str, limit: int) -> str:
    text = (text or "").strip()
    return text[: limit - 3] + "\u2026" if len(text) > limit else text


def _render_chart_for_distribution(
    distribution: dict,
    spec: Any,
    valid_n: int,
    display_mode: str,
    key_suffix: str = "",
) -> None:
    """Interactive Plotly bar chart for an SS or MS distribution."""
    try:
        import plotly.graph_objects as go
    except Exception:  # noqa: BLE001
        return

    app = _require_streamlit()
    labels: list[str] = []
    counts: list[float] = []
    pcts: list[float] = []
    for code, payload in distribution.items():
        labels.append(str(payload.get("label", code)))
        c = payload.get("count", 0) or 0
        counts.append(c)
        rate = payload.get("rate")
        if rate is None:
            rate = (c / valid_n) if valid_n else 0
        pcts.append(rate * 100)

    if not labels:
        return

    flags = _compute_outlier_flags(counts)
    bar_colors = [
        "#CC0000" if f == "high" else "#E65100" if f == "low" else "#888888"
        for f in flags
    ]
    if display_mode == "% only":
        y_values, y_label = pcts, "%"
    else:
        y_values, y_label = counts, "Count"

    text_vals: list[str] = []
    for c, p in zip(counts, pcts):
        if display_mode == "Counts":
            text_vals.append(f"{int(c):,}")
        elif display_mode == "% only":
            text_vals.append(f"{p:.1f}%")
        else:
            text_vals.append(f"{int(c):,} ({p:.1f}%)")

    fig = go.Figure(
        data=[
            go.Bar(
                x=labels,
                y=y_values,
                marker_color=bar_colors,
                text=text_vals,
                textposition="outside",
                customdata=list(zip(counts, pcts)),
                hovertemplate=(
                    "<b>%{x}</b><br>"
                    "Count: %{customdata[0]:,}<br>"
                    "%: %{customdata[1]:.1f}%<extra></extra>"
                ),
            )
        ]
    )
    fig.update_layout(
        title={
            "text": _chart_text_label(
                spec.question_text or spec.canonical_id, 80
            ),
            "font": {"family": "Arial", "size": 14, "color": "#0A0A0A"},
        },
        xaxis={"title": "", "tickfont": {"family": "Arial", "size": 11}},
        yaxis={
            "title": y_label,
            "tickfont": {"family": "Arial", "size": 10},
            "gridcolor": "#F0F0F0",
        },
        plot_bgcolor="white",
        paper_bgcolor="white",
        showlegend=False,
        height=380,
        margin={"t": 50, "b": 80, "l": 60, "r": 40},
        font={"family": "Arial"},
    )
    try:
        app.plotly_chart(
            fig,
            use_container_width=True,
            key=f"chart_sc_{spec.canonical_id}_{key_suffix}",
            config={
                "displayModeBar": True,
                "displaylogo": False,
                "modeBarButtonsToRemove": ["lasso2d", "select2d"],
                "toImageButtonOptions": {
                    "format": "png",
                    "filename": f"{spec.canonical_id}_chart",
                    "height": 600,
                    "width": 1000,
                    "scale": 2,
                },
            },
        )
    except Exception as exc:  # noqa: BLE001
        app.caption(f"Chart unavailable: {type(exc).__name__}: {exc}")


def _render_chart_for_cross_tab(
    result: Any, schema: Any, display_mode: str = "Counts"
) -> None:
    """Clustered bar chart for a cross-tab result."""
    try:
        import plotly.graph_objects as go
    except Exception:  # noqa: BLE001
        return

    app = _require_streamlit()
    rt = result.result_table
    counts = rt.get("counts", {}) or {}
    row_labels = rt.get("row_label_map", {}) or {}
    col_labels = rt.get("column_label_map", {}) or {}
    row_pct = rt.get("row_pct", {}) or {}
    if not counts:
        return

    row_codes = sorted(counts.keys(), key=lambda v: str(v))
    col_codes = sorted(
        {c for row in counts.values() if isinstance(row, dict) for c in row.keys()},
        key=lambda v: str(v),
    )
    x_categories = [str(row_labels.get(rc, rc)) for rc in row_codes]
    palette = [
        "#CC0000", "#0A0A0A", "#666666", "#990000",
        "#999999", "#330000", "#444444", "#FF6666",
    ]

    fig = go.Figure()
    for i, cc in enumerate(col_codes):
        col_label = str(col_labels.get(cc, cc))
        if display_mode == "Row %":
            y_vals = [
                (row_pct.get(rc, {}) or {}).get(cc, 0) * 100 for rc in row_codes
            ]
            text_vals = [f"{v:.1f}%" for v in y_vals]
        else:
            y_vals = [(counts.get(rc, {}) or {}).get(cc, 0) for rc in row_codes]
            text_vals = [f"{v:,}" for v in y_vals]
        fig.add_trace(
            go.Bar(
                name=col_label,
                x=x_categories,
                y=y_vals,
                text=text_vals,
                textposition="outside",
                marker_color=palette[i % len(palette)],
                hovertemplate=f"<b>{col_label}</b><br>%{{x}}: %{{y}}<extra></extra>",
            )
        )

    a, b = result.source_question_ids
    a_spec = schema.get_question(a) if schema is not None else None
    b_spec = schema.get_question(b) if schema is not None else None
    a_text = a_spec.question_text if (a_spec and a_spec.question_text) else a
    b_text = b_spec.question_text if (b_spec and b_spec.question_text) else b
    title = (
        f"{_chart_text_label(a_text, 50)} \u00d7 {_chart_text_label(b_text, 50)}"
    )

    fig.update_layout(
        title={"text": title, "font": {"family": "Arial", "size": 13}},
        barmode="group",
        xaxis={"title": a, "tickfont": {"family": "Arial", "size": 10}},
        yaxis={
            "title": "Count" if display_mode == "Counts" else "%",
            "gridcolor": "#F0F0F0",
            "tickfont": {"family": "Arial", "size": 10},
        },
        plot_bgcolor="white",
        paper_bgcolor="white",
        legend={"title": b, "font": {"family": "Arial", "size": 10}},
        height=440,
        margin={"t": 60, "b": 80, "l": 60, "r": 40},
        font={"family": "Arial"},
    )
    try:
        app.plotly_chart(
            fig,
            use_container_width=True,
            key=f"chart_ct_{result.cross_cut_id}_{display_mode}",
            config={
                "displayModeBar": True,
                "displaylogo": False,
                "toImageButtonOptions": {
                    "format": "png",
                    "filename": f"{result.cross_cut_id}_chart",
                    "height": 600,
                    "width": 1200,
                    "scale": 2,
                },
            },
        )
    except Exception as exc:  # noqa: BLE001
        app.caption(f"Chart unavailable: {type(exc).__name__}: {exc}")


def _render_chart_for_expected_vs_realized(result: Any) -> None:
    """Side-by-side bars comparing Expected, Realized, and Gap means."""
    try:
        import plotly.graph_objects as go
    except Exception:  # noqa: BLE001
        return

    app = _require_streamlit()
    rt = result.result_table
    expected = rt.get("expected", {}) or {}
    realized = rt.get("realized", {}) or {}
    gap = rt.get("gap", {}) or {}
    series = [
        ("Expected", expected.get("mean")),
        ("Realized", realized.get("mean")),
        ("Gap", gap.get("mean")),
    ]
    labels = [name for name, val in series if val is not None]
    values = [float(val) for name, val in series if val is not None]
    if not values:
        return

    palette = ["#666666", "#0A0A0A", "#CC0000"]
    fig = go.Figure(
        data=[
            go.Bar(
                x=labels,
                y=values,
                marker_color=palette[: len(values)],
                text=[f"{v:.2f}" for v in values],
                textposition="outside",
            )
        ]
    )
    fig.update_layout(
        title="Mean: Expected vs Realized (Gap = Realized \u2212 Expected)",
        xaxis={"tickfont": {"family": "Arial", "size": 11}},
        yaxis={
            "title": "Mean",
            "gridcolor": "#F0F0F0",
            "tickfont": {"family": "Arial", "size": 10},
        },
        plot_bgcolor="white",
        paper_bgcolor="white",
        showlegend=False,
        height=360,
        margin={"t": 50, "b": 60, "l": 60, "r": 40},
        font={"family": "Arial"},
    )
    try:
        app.plotly_chart(
            fig,
            use_container_width=True,
            key=f"chart_evr_{result.cross_cut_id}",
            config={"displaylogo": False},
        )
    except Exception as exc:  # noqa: BLE001
        app.caption(f"Chart unavailable: {type(exc).__name__}: {exc}")


def _render_chart_for_segment_metric(result: Any, schema: Any) -> None:
    """Simple bar chart of mean-by-segment for GROUP_COMPARISON-style results."""
    try:
        import plotly.graph_objects as go
    except Exception:  # noqa: BLE001
        return

    app = _require_streamlit()
    rt = result.result_table
    per_seg = rt.get("per_segment", {}) or {}
    labels: list[str] = []
    values: list[float] = []
    for seg_val, seg_data in per_seg.items():
        if not isinstance(seg_data, dict):
            continue
        mean = seg_data.get("mean")
        if mean is None:
            continue
        labels.append(str(seg_data.get("label", seg_val)))
        values.append(float(mean))
    if not values:
        return

    flags = _compute_outlier_flags(values)
    colors = [
        "#CC0000" if f == "high" else "#E65100" if f == "low" else "#666666"
        for f in flags
    ]
    fig = go.Figure(
        data=[
            go.Bar(
                x=labels,
                y=values,
                marker_color=colors,
                text=[f"{v:.2f}" for v in values],
                textposition="outside",
            )
        ]
    )
    fig.update_layout(
        title="Mean by segment",
        xaxis={"tickfont": {"family": "Arial", "size": 10}},
        yaxis={
            "gridcolor": "#F0F0F0",
            "tickfont": {"family": "Arial", "size": 10},
        },
        plot_bgcolor="white",
        paper_bgcolor="white",
        showlegend=False,
        height=380,
        margin={"t": 50, "b": 80, "l": 60, "r": 40},
        font={"family": "Arial"},
    )
    try:
        app.plotly_chart(
            fig,
            use_container_width=True,
            key=f"chart_seg_{result.cross_cut_id}",
            config={"displaylogo": False},
        )
    except Exception as exc:  # noqa: BLE001
        app.caption(f"Chart unavailable: {type(exc).__name__}: {exc}")


def _drain_pending_actions() -> None:
    """Process queued filter actions before any section renders.

    Implements the "set flag in click handler, compute at top of next render"
    pattern that eliminates the double-click bug: both the global filter and
    every per-question filter become single-click because the new state is
    fully populated before sections 2-5 render.
    """
    app = _require_streamlit()

    pending_gf = app.session_state.get("pending_global_filter")
    if pending_gf is not None:
        app.session_state["pending_global_filter"] = None
        app.session_state["global_filter_error"] = None
        try:
            from src.global_filter import apply_global_filter
            from src.models import GlobalFilterState

            state = GlobalFilterState(
                filters=tuple(pending_gf["filter_specs"])
            )
            filtered_df, stats = apply_global_filter(
                app.session_state["decoded_df"], state
            )
            app.session_state["global_filter_state"] = state
            app.session_state["global_filter_stats"] = stats
            app.session_state["active_df"] = filtered_df
            _rerun_single_cuts_on_active_df()
        except Exception as exc:  # noqa: BLE001
            app.session_state["global_filter_error"] = (
                f"{type(exc).__name__}: {exc}"
            )

    pending_pq = app.session_state.get("pending_per_question_filter") or {}
    if pending_pq:
        from src.filtered_single_cut import compute_filtered_single_cut

        schema = app.session_state.get("schema")
        active_df = app.session_state.get("active_df")
        log = app.session_state.get("log")
        errors: dict[str, str] = {}
        for cid, payload in list(pending_pq.items()):
            try:
                filtered_result = compute_filtered_single_cut(
                    cid, payload["filter_specs"], schema, active_df, log
                )
                app.session_state.setdefault("filtered_results", {})[cid] = (
                    filtered_result
                )
            except Exception as exc:  # noqa: BLE001
                errors[cid] = f"{type(exc).__name__}: {exc}"
        app.session_state["pending_per_question_filter"] = {}
        app.session_state["per_question_filter_errors"] = errors
        app.session_state["filtered_workbook_bytes"] = None


def _initialise_session_state() -> None:
    app = _require_streamlit()
    for key, value in SESSION_DEFAULTS.items():
        app.session_state.setdefault(key, value)
    stamped = app.session_state.get("cross_cut_engine_version")
    if stamped != CROSS_CUT_ENGINE_VERSION:
        app.session_state["cross_cut_results"] = []
        app.session_state["cross_cut_skips"] = []
        app.session_state["cross_cut_only_bytes"] = None
        app.session_state["filtered_results"] = {}
        app.session_state["filtered_workbook_bytes"] = None
        app.session_state["cross_cut_engine_version"] = CROSS_CUT_ENGINE_VERSION


def _upload_status(uploaded_file: Any | None) -> str:
    if uploaded_file is None:
        return "not uploaded"
    size_kb = uploaded_file.size / 1024
    return f"{uploaded_file.name} ({size_kb:.1f} KB)"


def _temp_dir() -> str | None:
    tmp_path = Path("/tmp")
    return str(tmp_path) if tmp_path.exists() else None


def _write_upload_to_temp(uploaded_file: Any) -> str:
    suffix = Path(uploaded_file.name).suffix
    with tempfile.NamedTemporaryFile(
        suffix=suffix,
        delete=False,
        dir=_temp_dir(),
    ) as temp_file:
        temp_file.write(uploaded_file.getbuffer())
        return temp_file.name


def _cleanup_temp_files(*paths: str | None) -> None:
    for path in paths:
        if path and os.path.exists(path):
            os.unlink(path)


# ---------------------------------------------------------------------------
# Pipeline orchestration
# ---------------------------------------------------------------------------


def _run_pipeline(
    data_map: Any,
    dataframe: Any,
    load_report: Any,
    status: Any,
) -> None:
    from src.calculation_log import CalculationLog
    from src.cross_cut_suggestions import suggest_cross_cuts
    from src.excel_exporter import export_single_cuts
    from src.models import DataQualityReport, GlobalFilterState
    from src.question_classifier import classify_questions
    from src.single_cut import compute_single_cuts

    status.update(label=PIPELINE_STAGES[0], state="complete")
    status.update(label=PIPELINE_STAGES[1], state="complete")
    # The unified io layer parses + decodes upstream and discards the
    # decoder's quality report. Reconstruct a minimal one so downstream
    # exporters don't crash while still surfacing parser-side warnings.
    parser_warnings = tuple(
        f"parser: {w}" for w in (load_report.parser_warnings or [])
    )
    quality_report = DataQualityReport(
        total_rows=int(len(dataframe)),
        total_columns=int(len(dataframe.columns)),
        columns_in_datamap=int(len(dataframe.columns)),
        columns_not_in_datamap=tuple(),
        per_column_missing_pct={col: 0.0 for col in dataframe.columns},
        per_column_out_of_range_pct={col: 0.0 for col in dataframe.columns},
        coercion_log=tuple(),
        warnings=parser_warnings,
    )

    status.update(label=PIPELINE_STAGES[2], state="running")
    schema = classify_questions(
        data_map,
        dataframe.columns.tolist(),
        respondent_id_column="record",
        total_respondents=len(dataframe),
        source_rawdata_path=load_report.raw_data_source,
    )

    status.update(label=PIPELINE_STAGES[3], state="running")
    log = CalculationLog()
    results, skips = compute_single_cuts(schema, dataframe, log)

    status.update(label=PIPELINE_STAGES[4], state="running")
    output_path = "/tmp/survey_analysis.xlsx"
    export_single_cuts(
        results=results,
        skips=skips,
        schema=schema,
        quality_report=quality_report,
        log=log,
        output_path=output_path,
    )

    app = _require_streamlit()
    app.session_state["decoded_df"] = dataframe
    app.session_state["active_df"] = dataframe
    app.session_state["global_filter_state"] = GlobalFilterState()
    app.session_state["global_filter_stats"] = None
    app.session_state["global_filter_rows"] = []
    app.session_state["results"] = results
    app.session_state["skips"] = skips
    app.session_state["schema"] = schema
    app.session_state["quality_report"] = quality_report
    app.session_state["log"] = log
    app.session_state["output_path"] = output_path
    app.session_state["cross_cut_results"] = []
    app.session_state["cross_cut_skips"] = []
    app.session_state["cross_cut_suggestions"] = suggest_cross_cuts(schema)
    app.session_state["cross_cut_only_bytes"] = None
    app.session_state["filtered_results"] = {}
    app.session_state["filtered_workbook_bytes"] = None
    app.session_state["run_complete"] = True
    status.update(label="Analysis complete.", state="complete")


def _refresh_full_workbook() -> None:
    from src.excel_exporter import export_single_cuts

    app = _require_streamlit()
    export_single_cuts(
        results=app.session_state["results"],
        skips=app.session_state["skips"],
        schema=app.session_state["schema"],
        quality_report=app.session_state["quality_report"],
        log=app.session_state["log"],
        output_path=app.session_state["output_path"],
        cross_cut_results=app.session_state["cross_cut_results"],
        cross_cut_skips=app.session_state["cross_cut_skips"],
    )


def _run_cross_cut_specs(specs: list[Any]) -> None:
    from src.cross_cut_engine import compute_cross_cuts

    if not specs:
        return

    app = _require_streamlit()
    results, skips = compute_cross_cuts(
        specs,
        app.session_state["schema"],
        app.session_state["active_df"],
        app.session_state["log"],
    )
    existing = {
        result.cross_cut_id: result
        for result in app.session_state["cross_cut_results"]
    }
    for result in results:
        existing[result.cross_cut_id] = result
        app.session_state.setdefault(f"cc_select_{result.cross_cut_id}", True)
    app.session_state["cross_cut_results"] = list(existing.values())
    app.session_state["cross_cut_skips"].extend(skips)
    app.session_state["cross_cut_only_bytes"] = None
    _refresh_full_workbook()


def _rerun_single_cuts_on_active_df() -> None:
    """Recompute single cuts after the active DataFrame changed."""
    from src.calculation_log import CalculationLog
    from src.single_cut import compute_single_cuts

    app = _require_streamlit()
    schema = app.session_state["schema"]
    active_df = app.session_state["active_df"]

    log = CalculationLog()
    results, skips = compute_single_cuts(schema, active_df, log)

    app.session_state["log"] = log
    app.session_state["results"] = results
    app.session_state["skips"] = skips
    app.session_state["cross_cut_results"] = []
    app.session_state["cross_cut_skips"] = []
    app.session_state["cross_cut_only_bytes"] = None
    app.session_state["filtered_results"] = {}
    app.session_state["filtered_workbook_bytes"] = None
    _refresh_full_workbook()


# ---------------------------------------------------------------------------
# Schema helpers
# ---------------------------------------------------------------------------


def _eligible_question_options() -> list[str]:
    schema = _require_streamlit().session_state["schema"]
    if schema is None:
        return []
    return [
        spec.canonical_id
        for spec in schema.questions
        if spec.analysis_eligible
    ]


def _question_label_map() -> dict[str, str]:
    """Full ID + question text, truncated to keep dropdowns readable."""
    schema = _require_streamlit().session_state["schema"]
    out: dict[str, str] = {}
    for spec in schema.questions:
        text = (spec.question_text or "").strip()
        if len(text) > 80:
            text = text[:77] + "\u2026"
        out[spec.canonical_id] = (
            f"{spec.canonical_id} \u2014 {text}" if text else spec.canonical_id
        )
    return out


def _eligible_filter_questions() -> list[Any]:
    """Categorical questions with int-coded option maps usable as filters."""
    schema = _require_streamlit().session_state["schema"]
    if schema is None:
        return []
    eligible = []
    for spec in schema.questions:
        if not spec.option_map:
            continue
        if not all(isinstance(key, int) for key in spec.option_map):
            continue
        eligible.append(spec)
    return eligible


# ---------------------------------------------------------------------------
# Cross-cut preview helpers (preserved from Day 10.6)
# ---------------------------------------------------------------------------


def _preview_cross_tab(result: Any) -> None:
    import pandas as pd
    app = _require_streamlit()
    ct = result.result_table
    a, b = result.source_question_ids
    counts = ct.get("counts", {})
    row_pct = ct.get("row_pct", {})
    column_pct = ct.get("column_pct", {})
    row_label_map = ct.get("row_label_map", {})
    col_label_map = ct.get("column_label_map", {})
    row_codes = sorted(counts.keys(), key=lambda v: str(v))
    col_codes = sorted(
        {c for row in counts.values() if isinstance(row, dict) for c in row.keys()},
        key=lambda v: str(v),
    )

    display_mode = app.radio(
        "Display",
        options=["Counts", "Row %", "Column %"],
        horizontal=True,
        key=f"preview_mode_{result.cross_cut_id}",
    )
    if display_mode == "Row %":
        source = row_pct
    elif display_mode == "Column %":
        source = column_pct
    else:
        source = counts

    df = pd.DataFrame(
        index=[row_label_map.get(rc, str(rc)) for rc in row_codes],
        columns=[col_label_map.get(cc, str(cc)) for cc in col_codes],
        data=[
            [source.get(rc, {}).get(cc, 0) for cc in col_codes]
            for rc in row_codes
        ],
    )
    df.index.name = f"\u2193 {a}"
    df.columns.name = f"\u2192 {b}"
    app.caption(f"Rows: {a}   Columns: {b}")
    schema = app.session_state.get("schema")
    chart_mode = "Row %" if display_mode == "Row %" else "Counts"
    _render_chart_for_cross_tab(result, schema, chart_mode)
    if display_mode == "Counts":
        _styled_dataframe(df, use_container_width=True)
    else:
        try:
            app.dataframe(
                df.style.apply(_style_outliers, axis=None).format("{:.1%}"),
                use_container_width=True,
            )
        except Exception:
            app.dataframe(df.style.format("{:.1%}"), use_container_width=True)
    app.caption(f"Grand total: {ct.get('grand_total', 0):,} responses")
    app.caption(
        "Tip: hover the table to use the built-in toolbar "
        "(search, fullscreen, download as CSV)."
    )


def _preview_segment_profile(result: Any) -> None:
    import pandas as pd
    app = _require_streamlit()
    rt = result.result_table
    app.caption(
        f"Filter: {rt.get('filter_expr', '<no filter>')}  \u00b7  "
        f"Filter N: {rt.get('filter_n', 0):,}"
    )
    tr = rt.get("target_result", {}) or {}
    if "distribution" in tr:
        rows = [
            {
                "Code": code,
                "Label": payload.get("label", ""),
                "Count": payload.get("count", 0),
            }
            for code, payload in sorted(
                tr["distribution"].items(), key=lambda x: str(x[0])
            )
        ]
        _styled_dataframe(
            pd.DataFrame(rows), use_container_width=True, hide_index=True
        )
    elif "selections" in tr:
        rows = []
        for sub_id, payload in tr["selections"].items():
            label = payload.get("label", "") or ""
            label_lower = label.lower()
            if "unchecked" in label_lower or "not selected" in label_lower:
                continue
            rows.append(
                {
                    "Sub-column": sub_id,
                    "Label": label,
                    "Selected count": payload.get("count", 0),
                }
            )
        rows.sort(key=lambda r: r["Selected count"], reverse=True)
        _styled_dataframe(
            pd.DataFrame(rows), use_container_width=True, hide_index=True
        )
        app.caption(
            "Selected counts only. Unchecked counts remain in the audit trail "
            "of the downloaded workbook."
        )
    elif "mean" in tr:
        df = pd.DataFrame(
            [
                {"Statistic": "Valid N", "Count": tr.get("valid_n", 0)},
                {"Statistic": "Missing N", "Count": tr.get("missing_n", 0)},
            ]
        )
        app.dataframe(df, use_container_width=True, hide_index=True)
        app.caption(
            "Numeric statistics (mean, median, std) in the downloaded workbook."
        )
    elif "rows" in tr:
        app.caption(f"Grid with {len(tr['rows'])} rows. Per-row counts:")
        grid_rows = []
        for sub_id, row_result in tr["rows"].items():
            dist = row_result.get("distribution", {}) if isinstance(row_result, dict) else {}
            row_dict: dict[str, Any] = {"Row": sub_id}
            for code, payload in dist.items():
                row_dict[f"{code}: {payload.get('label', '')}"] = payload.get("count", 0)
            grid_rows.append(row_dict)
        _styled_dataframe(
            pd.DataFrame(grid_rows), use_container_width=True, hide_index=True
        )
    else:
        app.info("Preview not available for this target type.")


def _preview_group_comparison(result: Any) -> None:
    import pandas as pd
    app = _require_streamlit()
    rt = result.result_table
    seg_q = rt.get("segment_question_id", "")
    met_q = rt.get("metric_question_id", "")
    app.caption(f"Metric: {met_q}   Segments: {seg_q}")
    _render_chart_for_segment_metric(result, app.session_state.get("schema"))
    rows = []
    for seg_val, seg_data in (rt.get("per_segment", {}) or {}).items():
        rows.append(
            {
                "Segment": seg_data.get("label", str(seg_val)) if isinstance(seg_data, dict) else str(seg_val),
                "N": seg_data.get("n", 0) if isinstance(seg_data, dict) else 0,
            }
        )
    overall = rt.get("overall", {}) or {}
    rows.append(
        {
            "Segment": "Overall",
            "N": overall.get("valid_n", overall.get("n", 0)),
        }
    )
    _styled_dataframe(
        pd.DataFrame(rows), use_container_width=True, hide_index=True
    )
    app.caption("Group means in the downloaded workbook.")


def _preview_expected_vs_realized(result: Any) -> None:
    import pandas as pd
    app = _require_streamlit()
    rt = result.result_table
    exp_q = rt.get("expected_question_id", "")
    real_q = rt.get("realized_question_id", "")
    app.caption(f"Expected: {exp_q}   Realized: {real_q}")
    _render_chart_for_expected_vs_realized(result)
    df = pd.DataFrame(
        [
            {"Metric": "Paired N", "Count": rt.get("paired_n", 0)},
            {"Metric": "Expected valid N", "Count": (rt.get("expected", {}) or {}).get("valid_n", 0)},
            {"Metric": "Realized valid N", "Count": (rt.get("realized", {}) or {}).get("valid_n", 0)},
        ]
    )
    _styled_dataframe(df, use_container_width=True, hide_index=True)
    app.caption(
        "Mean expected, mean realized, gap statistics in the downloaded workbook."
    )


def _render_cross_cut_preview(result: Any) -> None:
    app = _require_streamlit()
    from src.models import AnalysisType
    if app.checkbox(
        "Show preview (counts only)",
        value=True,
        key=f"cc_preview_{result.cross_cut_id}",
    ):
        try:
            at = result.analysis_type
            if at == AnalysisType.CROSS_TAB:
                _preview_cross_tab(result)
            elif at == AnalysisType.SEGMENT_PROFILE:
                _preview_segment_profile(result)
            elif at == AnalysisType.GROUP_COMPARISON:
                _preview_group_comparison(result)
            elif at == AnalysisType.EXPECTED_VS_REALIZED:
                _preview_expected_vs_realized(result)
            else:
                app.info(f"Preview not implemented for {at.value}.")
        except Exception as exc:  # noqa: BLE001
            app.error(f"Could not render preview: {type(exc).__name__}: {exc}")


# ---------------------------------------------------------------------------
# Single-cut display helpers (preserved from Day 12)
# ---------------------------------------------------------------------------


def _format_filter(filter_spec: Any) -> str:
    values = filter_spec.get_effective_values()
    if values is None:
        return f"{filter_spec.filter_question_id} (breakdown)"
    if len(values) == 1:
        return f"{filter_spec.filter_question_id} = {values[0]}"
    joined = ", ".join(str(v) for v in values)
    return f"{filter_spec.filter_question_id} \u2208 {{{joined}}}"


def _resolve_filter_value(schema: Any, df: Any, q_id: str, raw_value: Any) -> Any:
    """Reconcile UI option-map codes against actual raw-data column values.

    Some Word-derived data maps keep option labels ("India") in the raw data
    while the option_map exposes ``code -> label`` pairs. The UI hands back
    the code; if that code does not appear in the column we try str(code) and
    finally the option label, returning whichever value actually filters rows.
    """
    if df is None or q_id not in df.columns:
        return raw_value
    col_values = set(df[q_id].dropna().unique())
    if raw_value in col_values:
        return raw_value
    if str(raw_value) in col_values:
        return str(raw_value)
    q_spec = schema.get_question(q_id) if schema is not None else None
    if q_spec is not None and getattr(q_spec, "option_map", None):
        for key in (raw_value, str(raw_value)):
            label = q_spec.option_map.get(key)
            if label and label in col_values:
                return label
    return raw_value


def _normalize_value_list(val: Any) -> list:
    """Coerce a row's value-state (legacy scalar/None or new list) into a list."""
    if val is None:
        return []
    if isinstance(val, (list, tuple)):
        return [v for v in val if v is not None]
    return [val]


def _build_filter_spec(
    schema: Any, df: Any, q_id: str, vals: list
) -> Any:
    """Build a FilterSpec, resolving label/code mismatches per value."""
    from src.models import FilterSpec

    if not vals:
        return FilterSpec(filter_question_id=q_id)
    resolved = [_resolve_filter_value(schema, df, q_id, v) for v in vals]
    if len(resolved) == 1:
        return FilterSpec(filter_question_id=q_id, filter_value=resolved[0])
    return FilterSpec(filter_question_id=q_id, filter_values=tuple(resolved))


def _render_single_cut_result(result: Any, spec: Any) -> None:
    """Display a SingleCutResult with branded HTML for SS/MS distributions."""
    import pandas as pd
    from src.models import (
        GridSingleSelectResult,
        MultiSelectResult,
        NumericResult,
        SingleSelectResult,
    )

    app = _require_streamlit()
    app.caption(
        f"Valid N: {result.valid_n:,}  \u00b7  Missing N: {result.missing_n:,}"
    )
    if isinstance(result, GridSingleSelectResult):
        grid_rows: list[dict[str, Any]] = []
        for sub_id, row_result in result.rows.items():
            row_label = (
                spec.grid_row_labels.get(sub_id, sub_id)
                if spec.grid_row_labels
                else sub_id
            )
            row_dict: dict[str, Any] = {"Row": row_label}
            for code, payload in row_result.distribution.items():
                row_dict[f"{code}: {payload.get('label', '')}"] = payload.get(
                    "count", 0
                )
            grid_rows.append(row_dict)
        _styled_dataframe(
            pd.DataFrame(grid_rows), use_container_width=True, hide_index=True
        )
    elif isinstance(result, SingleSelectResult):
        sorted_dist = dict(
            sorted(result.distribution.items(), key=lambda kv: str(kv[0]))
        )
        # Compute flags ONCE from counts so highlighting stays identical
        # across Counts / Counts+% / % only toggles.
        ss_counts = [p.get("count", 0) for _, p in sorted_dist.items()]
        ss_flags = _compute_outlier_flags(ss_counts)
        display_mode = app.radio(
            "Display",
            options=["Counts", "Counts + %", "% only"],
            index=1,
            horizontal=True,
            key=f"sc_display_{result.question_id}",
            label_visibility="collapsed",
        )
        _render_chart_for_distribution(
            sorted_dist,
            spec,
            result.valid_n,
            display_mode,
            key_suffix=f"ss_{result.question_id}",
        )
        _render_sc_table_html(
            sorted_dist,
            display_mode,
            result.valid_n,
            flags=ss_flags,
            key_suffix=str(result.question_id),
        )
    elif isinstance(result, MultiSelectResult):
        ms_dist: dict[Any, dict[str, Any]] = {}
        for sub_id, payload in result.selections.items():
            label = payload.get("label", "") or ""
            label_lower = label.lower()
            if "unchecked" in label_lower or "not selected" in label_lower:
                continue
            # MultiSelectResult payloads use "selection_rate"; tolerate "rate"
            # for forward-compat. Either way, _render_sc_table_html will fall
            # back to count/valid_n if the rate is missing.
            ms_rate = payload.get("selection_rate")
            if ms_rate is None:
                ms_rate = payload.get("rate")
            ms_dist[sub_id] = {
                "label": label or sub_id,
                "count": payload.get("count", 0),
                "rate": ms_rate,
            }
        ms_dist = dict(
            sorted(ms_dist.items(), key=lambda kv: kv[1]["count"], reverse=True)
        )
        # Compute flags ONCE from selected counts.
        ms_counts = [p.get("count", 0) for _, p in ms_dist.items()]
        ms_flags = _compute_outlier_flags(ms_counts)
        display_mode = app.radio(
            "Display",
            options=["Counts", "Counts + %", "% only"],
            index=1,
            horizontal=True,
            key=f"sc_display_{result.question_id}",
            label_visibility="collapsed",
        )
        _render_chart_for_distribution(
            ms_dist,
            spec,
            result.valid_n,
            display_mode,
            key_suffix=f"ms_{result.question_id}",
        )
        _render_sc_table_html(
            ms_dist,
            display_mode,
            result.valid_n,
            flags=ms_flags,
            key_suffix=f"ms_{result.question_id}",
        )
        app.caption(
            "Selected counts only. Unchecked counts remain in the audit trail "
            "of the downloaded workbook."
        )
    elif isinstance(result, NumericResult):
        rows = [
            {"Statistic": "Valid N", "Value": result.valid_n},
            {"Statistic": "Missing N", "Value": result.missing_n},
        ]
        app.dataframe(
            pd.DataFrame(rows), use_container_width=True, hide_index=True
        )
        app.caption("Mean, median, std and percentiles in the downloaded workbook.")
    else:
        app.info("Preview not available for this result type.")


# ---------------------------------------------------------------------------
# Per-question multi-row filter panel (Day 14)
# ---------------------------------------------------------------------------


def _find_q_index(q_id: str | None, options: list[tuple[str, Any]]) -> int:
    for i, (_label, value) in enumerate(options):
        if value == q_id:
            return i
    return 0


def _purge_widget_keys(*prefixes: str) -> None:
    """Delete widget-state keys with the given prefixes.

    Streamlit's selectbox honors the ``index=`` arg only on the first render
    of a given key; once that key exists in ``session_state`` the stored value
    overrides ``index=``. When the list of rows changes (add/delete), the
    index-based keys point at the wrong rows, so we purge them and let the
    row-state seed the next render.
    """
    app = _require_streamlit()
    for key in list(app.session_state.keys()):
        if any(key.startswith(prefix) for prefix in prefixes):
            del app.session_state[key]


def _render_single_cut_card(
    result: Any, spec: Any, expanded: bool = False
) -> None:
    from src.models import FilterSpec  # noqa: F401  (used by transitive helpers)

    app = _require_streamlit()
    schema = app.session_state["schema"]
    short_text = (spec.question_text or "")[:80]

    expander_label = f"{spec.canonical_id} \u2014 {short_text}"
    with app.expander(expander_label, expanded=expanded):
        app.caption(
            f"Type: {spec.question_type.value}  \u00b7  "
            f"Valid N: {result.valid_n:,}  \u00b7  "
            f"Missing: {result.missing_n:,}"
        )

        app.markdown("**Filters for this question**")
        app.caption(TOOLTIP_PER_QUESTION_FILTER)

        filter_key = f"filters_{spec.canonical_id}"
        if filter_key not in app.session_state:
            app.session_state[filter_key] = []
        filter_rows: list[tuple[str | None, Any]] = app.session_state[filter_key]

        eligible = _eligible_filter_questions()
        question_options: list[tuple[str, Any]] = [("None", None)] + [
            (
                f"{q.canonical_id}: {(q.question_text or '')[:50]}",
                q.canonical_id,
            )
            for q in eligible
            if q.canonical_id != spec.canonical_id
        ]

        new_rows: list[tuple[str | None, Any]] = []
        delete_index: int | None = None
        for i, (q_id, val) in enumerate(filter_rows):
            cols = app.columns([4, 4, 1])
            with cols[0]:
                q_pick = app.selectbox(
                    "Filter question",
                    options=question_options,
                    format_func=lambda x: x[0],
                    index=_find_q_index(q_id, question_options),
                    key=f"{filter_key}_q_{i}",
                    label_visibility="visible" if i == 0 else "collapsed",
                )
            picked_q_id = q_pick[1]
            with cols[1]:
                if picked_q_id is not None:
                    q_spec = schema.get_question(picked_q_id)
                    value_codes = list(q_spec.option_map.keys())
                    prior = _normalize_value_list(val)
                    default_codes = [v for v in prior if v in value_codes]
                    v_pick = app.multiselect(
                        "Values (leave empty for breakdown)",
                        options=value_codes,
                        format_func=lambda v: f"{v}: {q_spec.option_map[v]}",
                        default=default_codes,
                        key=f"{filter_key}_v_{i}",
                        label_visibility="visible" if i == 0 else "collapsed",
                        help=TOOLTIP_BREAKDOWN if i == 0 else None,
                        placeholder="All values (breakdown)",
                    )
                    new_val = list(v_pick)
                else:
                    app.multiselect(
                        "Values",
                        options=[],
                        disabled=True,
                        key=f"{filter_key}_v_disabled_{i}",
                        label_visibility="visible" if i == 0 else "collapsed",
                        placeholder="Pick a question first",
                    )
                    new_val = []
            with cols[2]:
                if i == 0:
                    app.markdown("&nbsp;", unsafe_allow_html=True)
                if app.button(
                    "\u2715",
                    key=f"{filter_key}_del_{i}",
                    help="Remove this filter",
                ):
                    delete_index = i
            new_rows.append((picked_q_id, new_val))

        app.session_state[filter_key] = new_rows

        if delete_index is not None:
            new_rows.pop(delete_index)
            app.session_state[filter_key] = new_rows
            _purge_widget_keys(f"{filter_key}_q_", f"{filter_key}_v_")
            app.rerun()

        cols_btn = app.columns([2, 2, 4])
        with cols_btn[0]:
            if app.button("+ Add filter", key=f"{filter_key}_add"):
                app.session_state[filter_key] = new_rows + [(None, None)]
                _purge_widget_keys(f"{filter_key}_q_", f"{filter_key}_v_")
                app.rerun()
        with cols_btn[1]:
            apply_clicked = app.button(
                "Apply filters",
                key=f"{filter_key}_apply",
                type="primary",
                disabled=not any(q is not None for q, _v in new_rows),
            )

        pq_errors = app.session_state.get("per_question_filter_errors") or {}
        if spec.canonical_id in pq_errors:
            app.error(f"Filter failed: {pq_errors[spec.canonical_id]}")

        if apply_clicked:
            active_df = app.session_state["active_df"]
            specs = [
                _build_filter_spec(schema, active_df, q, _normalize_value_list(v))
                for q, v in new_rows
                if q is not None
            ]
            seen: set[str] = set()
            duplicate = next(
                (f.filter_question_id for f in specs if f.filter_question_id in seen
                 or seen.add(f.filter_question_id)),
                None,
            )
            breakdowns = [f for f in specs if f.is_breakdown()]
            if duplicate is not None:
                app.error(
                    f"Duplicate filter on {duplicate}. "
                    "Each filter question can only be used once per card."
                )
            elif len(breakdowns) > 1:
                app.error(
                    "Only one breakdown filter at a time. "
                    "Pick a value for at least all but one of the breakdowns."
                )
            else:
                # Queue for processing in _drain_pending_actions on next rerun.
                # Inline compute would require a second click for state to
                # propagate; deferring fixes the double-click bug.
                app.session_state.setdefault(
                    "pending_per_question_filter", {}
                )[spec.canonical_id] = {"filter_specs": specs}
                app.session_state.setdefault(
                    "per_question_filter_errors", {}
                ).pop(spec.canonical_id, None)
                app.rerun()

        app.divider()

        filtered = app.session_state.get("filtered_results", {}).get(
            spec.canonical_id
        )
        if filtered is not None:
            app.info(
                "Filtered: "
                + ", ".join(_format_filter(f) for f in filtered.filters_applied)
                + f"  \u00b7  N = {filtered.filtered_n:,}"
            )
            for warning in filtered.warnings:
                app.warning(warning)
            if filtered.dispatch_mode == "single_cut_filtered":
                _render_single_cut_result(filtered.single_cut_result, spec)
            elif filtered.dispatch_mode == "cross_cut_breakdown":
                _render_cross_cut_preview(filtered.cross_cut_result)
            app.checkbox(
                "Include in filtered workbook download",
                value=True,
                key=f"fsc_select_{spec.canonical_id}",
            )
            if app.button(
                "Clear filters", key=f"clear_filter_{spec.canonical_id}"
            ):
                del app.session_state["filtered_results"][spec.canonical_id]
                app.session_state[filter_key] = []
                app.session_state["filtered_workbook_bytes"] = None
                app.rerun()
        else:
            _render_single_cut_result(result, spec)


# ---------------------------------------------------------------------------
# Cross-cut helpers (preserved)
# ---------------------------------------------------------------------------


def _render_suggested_cross_cuts() -> None:
    app = _require_streamlit()
    suggestions = app.session_state["cross_cut_suggestions"]
    if not suggestions:
        app.write("No rule-based suggestions available for this schema.")
        return

    import html as _html

    for index, (spec, reason) in enumerate(suggestions[:15], start=1):
        col_text, col_button = app.columns([4, 1])
        with col_text:
            app.markdown(
                f'<div style="font-size:13px;font-weight:600;color:#0A0A0A;'
                f'font-family:Arial;">{index}. {_html.escape(spec.title)}</div>'
                f'<div style="font-size:10px;color:#888;margin-top:2px;'
                f'font-family:Arial;">'
                f'{_html.escape(" · ".join(spec.source_question_ids))}'
                f'</div>',
                unsafe_allow_html=True,
            )
            app.caption(reason)
        if col_button.button("Run this", key=f"run_suggestion_{spec.cross_cut_id}"):
            _run_cross_cut_specs([spec])
            app.success(f"Ran {spec.cross_cut_id}")


def _render_manual_cross_cut() -> None:
    from src.models import AnalysisType, CrossCutSpec

    app = _require_streamlit()
    options = _eligible_question_options()
    labels = _question_label_map()
    if len(options) < 2:
        app.write("At least two analysis-eligible questions are required.")
        return

    with app.form("manual_cross_cut_form"):
        analysis_type_name = app.selectbox(
            "Analysis type",
            [
                AnalysisType.CROSS_TAB.value,
                AnalysisType.SEGMENT_PROFILE.value,
                AnalysisType.GROUP_COMPARISON.value,
                AnalysisType.EXPECTED_VS_REALIZED.value,
            ],
        )
        first = app.selectbox(
            "First source question",
            options,
            format_func=lambda value: labels.get(value, value),
        )
        second = app.selectbox(
            "Second source question",
            options,
            index=1,
            format_func=lambda value: labels.get(value, value),
        )
        filter_expr = app.text_input(
            "Filter expression for segment profile",
            value=f"{first} == 1",
            help="Required only for SEGMENT_PROFILE. Supports equality, e.g. Q3 == 1.",
        )
        submitted = app.form_submit_button("Run manual cross cut", type="primary")

    if not submitted:
        return

    analysis_type = AnalysisType(analysis_type_name)
    try:
        spec = CrossCutSpec(
            cross_cut_id=f"MANUAL_{analysis_type.value}_{first}_{second}",
            title=f"{analysis_type.value}: {first} x {second}",
            analysis_type=analysis_type,
            source_question_ids=(first, second),
            filter_expr=filter_expr if analysis_type is AnalysisType.SEGMENT_PROFILE else None,
            filter_mask_description=filter_expr if analysis_type is AnalysisType.SEGMENT_PROFILE else None,
        )
        _run_cross_cut_specs([spec])
        app.success(f"Ran {spec.cross_cut_id}")
    except Exception as exc:  # noqa: BLE001
        app.error(f"{type(exc).__name__}: {exc}")


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------


def _render_sidebar() -> None:
    """Branded, structured sidebar showing only what is useful mid-session."""
    import html as _html

    app = _require_streamlit()
    with app.sidebar:
        # ---- LOGO / TITLE ---------------------------------------------------
        app.markdown(
            """
            <div style="padding:16px 0 8px 0;border-bottom:3px solid #CC0000;
              margin-bottom:16px;">
              <div style="font-size:16px;font-weight:700;color:#0A0A0A;
                font-family:Arial;letter-spacing:0.03em;">
                Survey Analysis Engine
              </div>
              <div style="font-size:10px;color:#888;font-family:Arial;
                margin-top:2px;">Bain &amp; Company</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if app.session_state.get("run_complete"):
            lr = app.session_state.get("load_report")
            gf_stats = app.session_state.get("global_filter_stats")
            results = app.session_state.get("results", [])

            app.markdown("**Session**")

            if lr is not None:
                app.markdown(
                    f"""
                    <div style="font-size:10px;color:#666;font-family:Arial;
                      line-height:1.9;background:#F8F8F8;padding:10px;
                      border-left:3px solid #E0E0E0;margin-bottom:12px;">
                      <b>File</b><br>{_html.escape(str(lr.raw_data_source))}<br>
                      <b>Input type</b><br>
                      {_html.escape(lr.scenario.replace('_', ' '))}<br>
                      <b>Respondents</b><br>{lr.raw_rows:,}<br>
                      <b>Questions</b><br>{lr.questions_parsed}
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            gf_state = app.session_state.get("global_filter_state")
            if (
                gf_state is not None
                and gf_state.is_active()
                and gf_stats
            ):
                app.markdown(
                    f"""
                    <div style="background:#FFF5F5;border-left:3px solid #CC0000;
                      padding:10px;font-size:10px;font-family:Arial;
                      line-height:1.9;margin-bottom:12px;">
                      <b style="color:#CC0000;">● Global Filter Active</b><br>
                      {_html.escape(gf_state.description())}<br>
                      <b>{gf_stats.get('rows_after', 0):,}</b>&nbsp;of&nbsp;
                      {gf_stats.get('rows_before', 0):,}&nbsp;respondents
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            else:
                app.markdown(
                    """
                    <div style="font-size:10px;color:#888;font-family:Arial;
                      padding:8px 0;margin-bottom:8px;">
                      ○ No global filter · full dataset
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            skips = app.session_state.get("skips", [])
            errors = [s for s in skips if s.skip_reason == "calculation_error"]
            log = app.session_state.get("log")
            log_len = len(log) if log else 0
            cc_results = app.session_state.get("cross_cut_results", [])
            err_color = "#CC0000" if errors else "#2E7D32"
            app.markdown(
                f"""
                <div style="font-size:10px;color:#666;font-family:Arial;
                  line-height:2.0;border-top:1px solid #E0E0E0;
                  padding-top:10px;margin-bottom:12px;">
                  <b>Single cuts</b>&nbsp;{len(results)}<br>
                  <b>Skipped</b>&nbsp;{len(skips)}<br>
                  <b>Errors</b>&nbsp;
                  <span style="color:{err_color};">{len(errors)}</span><br>
                  <b>Audit records</b>&nbsp;{log_len}<br>
                  <b>Cross cuts</b>&nbsp;{len(cc_results)}
                </div>
                """,
                unsafe_allow_html=True,
            )

            # ---- QUESTIONS (master-detail nav) ----------------------------
            app.markdown("**Questions**")
            selected_qid = app.session_state.get("selected_question_id")
            if selected_qid is None and results:
                selected_qid = results[0].question_id
                app.session_state["selected_question_id"] = selected_qid

            app.text_input(
                "Search",
                placeholder="Filter by ID or text",
                key="sidebar_q_search",
                label_visibility="collapsed",
            )
            needle = (
                app.session_state.get("sidebar_q_search") or ""
            ).strip().lower()

            schema_obj = app.session_state.get("schema")
            shown = 0
            for result in results:
                if schema_obj is None:
                    continue
                spec = schema_obj.get_question(result.question_id)
                if spec is None:
                    continue
                if needle:
                    hay = f"{spec.canonical_id} {spec.question_text or ''}".lower()
                    if needle not in hay:
                        continue
                shown += 1
                label = (
                    f"{spec.canonical_id} \u2014 "
                    f"{spec.question_text or ''}".strip()
                )
                btn_type = (
                    "primary"
                    if result.question_id == selected_qid
                    else "secondary"
                )
                if app.button(
                    label,
                    key=f"sidebar_qbtn_{result.question_id}",
                    type=btn_type,
                    use_container_width=True,
                ):
                    app.session_state["selected_question_id"] = (
                        result.question_id
                    )
                    app.rerun()
            if needle and shown == 0:
                app.caption(f"No questions match '{needle}'.")
            else:
                app.caption(f"{shown} of {len(results)} shown")

            app.markdown("**Navigate**")
            app.markdown(
                """
                <div style="font-size:11px;font-family:Arial;line-height:2.2;
                  color:#333;">
                  <a href="#section-1" style="color:#CC0000;
                    text-decoration:none;">↑ 1. Upload</a><br>
                  <a href="#section-2" style="color:#CC0000;
                    text-decoration:none;">⊕ 2. Global Filter</a><br>
                  <a href="#section-3" style="color:#CC0000;
                    text-decoration:none;">— 3. Single Cuts</a><br>
                  <a href="#section-4" style="color:#CC0000;
                    text-decoration:none;">╫ 4. Cross Cuts</a><br>
                  <a href="#section-5" style="color:#CC0000;
                    text-decoration:none;">↓ 5. Downloads</a>
                </div>
                """,
                unsafe_allow_html=True,
            )

        else:
            app.markdown(
                """
                <div style="font-size:11px;color:#666;font-family:Arial;
                  line-height:1.9;padding:8px 0;">
                  <b>Getting started:</b><br>
                  1. Upload your survey files<br>
                  2. Click Run Analysis<br>
                  3. Apply filters if needed<br>
                  4. Explore and download
                </div>
                """,
                unsafe_allow_html=True,
            )

        app.markdown("---")
        with app.expander("Help", expanded=False):
            app.markdown(
                """
                <div style="font-size:10px;font-family:Arial;color:#444;
                  line-height:1.9;">
                  <b>Supported file types</b><br>CSV, XLSX, DOCX<br><br>
                  <b>Three input scenarios</b><br>
                  A: Raw data + data map<br>
                  B: Combined XLSX (2 sheets)<br>
                  C: Raw data + Word survey doc<br><br>
                  <b>Filters</b><br>
                  Global filter restricts everything.<br>
                  Per-question filter applies only<br>
                  to that single question.<br><br>
                  <b>Outlier flags</b><br>
                  ⬆ Red = high outlier (&gt;2σ above mean)<br>
                  ↓ Amber = low outlier (&gt;1.5σ below)
                </div>
                """,
                unsafe_allow_html=True,
            )


# ---------------------------------------------------------------------------
# Section 1 — Upload
# ---------------------------------------------------------------------------


def _section_upload() -> None:
    app = _require_streamlit()
    _section_header("1", SECTION_UPLOAD, anchor="section-1", meta="CSV \u00b7 XLSX \u00b7 DOCX")
    app.markdown(
        "Drop **any combination** of files \u2014 the tool detects what "
        "you've uploaded automatically."
    )

    uploaded_files = app.file_uploader(
        "Upload your survey files",
        type=["csv", "xlsx", "docx"],
        accept_multiple_files=True,
        key="unified_upload",
        help=(
            "**Three scenarios are supported:**\n\n"
            "1. Raw data (.csv or .xlsx) + data map (.xlsx) \u2014 two "
            "separate files\n\n"
            "2. Combined .xlsx with raw data and data map on separate "
            "sheets\n\n"
            "3. Raw data (.csv or .xlsx) + Word survey document (.docx) "
            "\u2014 tool auto-builds the data map from the Word doc"
        ),
    )

    detected_scenario: str | None = None
    docx_only = False
    if uploaded_files:
        names = [f.name for f in uploaded_files]
        app.caption(f"Uploaded: {', '.join(names)}")
        try:
            from src.io import _detect_scenario
            detected_scenario = _detect_scenario(uploaded_files)
        except Exception:
            detected_scenario = None
        scenario_labels = {
            "A_separate_files":
                "\u2713 Two-file input detected (raw data + data map)",
            "B_combined_xlsx":
                "\u2713 Combined Excel detected (data + map on separate sheets)",
            "C_word_datamap":
                "\u2713 Word survey document detected (data map will be parsed "
                "from Word file)",
        }
        if detected_scenario in scenario_labels:
            app.info(scenario_labels[detected_scenario])

        # Scenario C without raw data: graceful handling, no exception.
        docx_files = [f for f in uploaded_files if f.name.lower().endswith(".docx")]
        non_docx = [f for f in uploaded_files if not f.name.lower().endswith(".docx")]
        docx_only = bool(docx_files) and not non_docx

    # Update legacy status labels for the sidebar "files uploaded" indicator.
    if uploaded_files:
        app.session_state["raw_data_path_label"] = ", ".join(
            f.name for f in uploaded_files
        )
        app.session_state["datamap_path_label"] = ", ".join(
            f.name for f in uploaded_files
        )

    if docx_only:
        app.warning(
            "Data map will be parsed from the Word document, but no raw "
            "data file was uploaded. Please also upload your raw data "
            "(.csv or .xlsx) to run the analysis."
        )

    ready = bool(uploaded_files) and not docx_only
    centre_left, centre_mid, centre_right = app.columns([2, 3, 2])
    with centre_mid:
        run_clicked = app.button(
            "Run analysis",
            type="primary",
            disabled=not ready,
            use_container_width=True,
        )

    if run_clicked:
        from src.io import load_survey_inputs

        app.session_state["run_complete"] = False
        try:
            with app.status("Starting analysis...", expanded=True) as status:
                status.update(label="Loading uploaded files...", state="running")
                data_map, raw_df, load_report = load_survey_inputs(uploaded_files)
                app.session_state["data_map"] = data_map
                app.session_state["load_report"] = load_report
                for note in load_report.detection_notes:
                    app.caption(f"\u2139\ufe0f {note}")
                _run_pipeline(data_map, raw_df, load_report, status)
        except Exception as exc:  # noqa: BLE001
            app.session_state["run_complete"] = False
            app.error(f"{type(exc).__name__}: {exc}")
            with app.expander("Show full traceback"):
                app.code(traceback.format_exc())

    if app.session_state["run_complete"]:
        schema = app.session_state["schema"]
        results = app.session_state["results"]
        skips = app.session_state["skips"]
        log = app.session_state["log"]
        m1, m2, m3, m4 = app.columns(4)
        m1.metric("Total questions", len(schema.questions))
        m2.metric("Single cuts produced", len(results))
        m3.metric("Skipped", len(skips))
        m4.metric("Audit records", len(log))


# ---------------------------------------------------------------------------
# Section 2 — Global filter
# ---------------------------------------------------------------------------


def _apply_global_filter_action(rows: list[tuple[str | None, Any]]) -> None:
    """Queue the global filter for processing on the next rerun.

    The actual ``apply_global_filter`` + single-cut recompute runs in
    ``_drain_pending_actions()`` at the top of ``main()``. Storing only the
    spec list here (rather than computing inline) eliminates the double-click
    bug: the click handler returns immediately, ``st.rerun()`` fires, and the
    recompute happens before any section renders so sections 2-5 all see the
    filtered state in a single round trip.
    """
    app = _require_streamlit()
    schema = app.session_state.get("schema")
    decoded_df = app.session_state.get("decoded_df")
    spec_list = []
    for q, v in rows:
        vals = _normalize_value_list(v)
        if q is None or not vals:
            continue
        spec_list.append(_build_filter_spec(schema, decoded_df, q, vals))
    app.session_state["pending_global_filter"] = {"filter_specs": spec_list}
    app.session_state["global_filter_error"] = None
    app.rerun()


def _clear_global_filter_action() -> None:
    from src.models import GlobalFilterState

    app = _require_streamlit()
    app.session_state["global_filter_state"] = GlobalFilterState()
    app.session_state["global_filter_stats"] = None
    app.session_state["global_filter_rows"] = []
    app.session_state["active_df"] = app.session_state["decoded_df"]
    _rerun_single_cuts_on_active_df()
    app.rerun()


def _section_global_filter() -> None:
    app = _require_streamlit()
    _section_header("2", SECTION_GLOBAL_FILTER, anchor="section-2", meta="Subset every analysis")

    if not app.session_state["run_complete"]:
        app.caption(
            "Run an analysis first \u2014 the global filter becomes available "
            "once the data is loaded."
        )
        return

    app.markdown(TOOLTIP_GLOBAL_FILTER)

    schema = app.session_state["schema"]
    gf_state = app.session_state.get("global_filter_state")
    stats = app.session_state.get("global_filter_stats") or {}

    if gf_state is not None and gf_state.is_active():
        rows_before = stats.get("rows_before", 0)
        rows_after = stats.get("rows_after", 0)
        app.info(
            f"\U0001F535 Global filter active: {gf_state.description()}. "
            f"All analyses below restricted to {rows_after:,} of "
            f"{rows_before:,} respondents."
        )
    else:
        app.caption(
            f"No global filter. All analyses run on the full "
            f"{schema.total_respondents:,} respondents."
        )

    eligible = _eligible_filter_questions()
    if not eligible:
        app.warning("No categorical questions available to use as global filters.")
        return

    question_options: list[tuple[str, Any]] = [("None", None)] + [
        (
            f"{q.canonical_id}: {(q.question_text or '')[:50]}",
            q.canonical_id,
        )
        for q in eligible
    ]

    rows: list[tuple[str | None, Any]] = list(
        app.session_state.get("global_filter_rows", [])
    )
    if not rows:
        rows = [(None, None)]

    new_rows: list[tuple[str | None, Any]] = []
    delete_index: int | None = None
    for i, (q_id, val) in enumerate(rows):
        cols = app.columns([4, 4, 1])
        with cols[0]:
            q_pick = app.selectbox(
                "Filter question",
                options=question_options,
                format_func=lambda x: x[0],
                index=_find_q_index(q_id, question_options),
                key=f"gf_q_{i}",
                label_visibility="visible" if i == 0 else "collapsed",
            )
        picked_q_id = q_pick[1]
        with cols[1]:
            if picked_q_id is not None:
                q_spec = schema.get_question(picked_q_id)
                value_codes = list(q_spec.option_map.keys())
                prior = _normalize_value_list(val)
                default_codes = [v for v in prior if v in value_codes]
                v_pick = app.multiselect(
                    "Values (select one or more)",
                    options=value_codes,
                    format_func=lambda v: f"{v}: {q_spec.option_map[v]}",
                    default=default_codes,
                    key=f"gf_v_{i}",
                    label_visibility="visible" if i == 0 else "collapsed",
                    placeholder="Pick value(s)",
                )
                new_val = list(v_pick)
            else:
                app.multiselect(
                    "Values",
                    options=[],
                    disabled=True,
                    key=f"gf_v_disabled_{i}",
                    label_visibility="visible" if i == 0 else "collapsed",
                    placeholder="Pick a question first",
                )
                new_val = []
        with cols[2]:
            if i == 0:
                app.markdown("&nbsp;", unsafe_allow_html=True)
            if app.button("\u2715", key=f"gf_del_{i}", help="Remove this filter"):
                delete_index = i
        new_rows.append((picked_q_id, new_val))

    app.session_state["global_filter_rows"] = new_rows

    if delete_index is not None:
        new_rows.pop(delete_index)
        app.session_state["global_filter_rows"] = new_rows
        _purge_widget_keys("gf_q_", "gf_v_")
        app.rerun()

    btn_cols = app.columns([2, 2, 2, 4])
    with btn_cols[0]:
        if app.button("+ Add another filter", key="gf_add"):
            app.session_state["global_filter_rows"] = new_rows + [(None, [])]
            _purge_widget_keys("gf_q_", "gf_v_")
            app.rerun()
    with btn_cols[1]:
        complete_rows = [
            (q, v) for q, v in new_rows
            if q is not None and _normalize_value_list(v)
        ]
        if app.button(
            "Apply global filter",
            key="gf_apply",
            type="primary",
            disabled=not complete_rows,
        ):
            _apply_global_filter_action(new_rows)
    with btn_cols[2]:
        if gf_state is not None and gf_state.is_active():
            if app.button("Clear global filter", key="gf_clear"):
                _clear_global_filter_action()


# ---------------------------------------------------------------------------
# Section 3 — Single cuts
# ---------------------------------------------------------------------------


def _section_single_cuts() -> None:
    """Day 18 master-detail: show only the selected question's analysis."""
    import html as _html

    app = _require_streamlit()
    _section_header("3", SECTION_RESULTS, anchor="section-3", meta="Single cuts")

    if not app.session_state["run_complete"]:
        app.info(EMPTY_NO_RESULTS)
        return

    results = app.session_state["results"]
    schema = app.session_state["schema"]
    if not results:
        app.info(EMPTY_NO_RESULTS)
        return

    selected = app.session_state.get("selected_question_id")
    valid_ids = {r.question_id for r in results}
    if selected not in valid_ids:
        selected = results[0].question_id
        app.session_state["selected_question_id"] = selected

    result = next((r for r in results if r.question_id == selected), results[0])
    spec = schema.get_question(result.question_id)
    if spec is None:
        app.info("Selected question is not available.")
        return

    title = _html.escape(spec.question_text or spec.canonical_id)
    qtype = spec.question_type.value if hasattr(spec.question_type, "value") else str(spec.question_type)
    app.markdown(
        f"""
        <div style="background:white;border:1px solid #E0E0E0;
          border-left:4px solid #CC0000;padding:18px 22px;margin-bottom:16px;
          font-family:Arial;">
          <div style="font-size:11px;font-weight:700;letter-spacing:0.1em;
            color:#CC0000;text-transform:uppercase;">
            {_html.escape(spec.canonical_id)} &middot; {_html.escape(qtype)}
            &middot; Valid N {result.valid_n:,}
          </div>
          <div style="font-size:18px;font-weight:600;color:#0A0A0A;
            margin-top:6px;line-height:1.4;">{title}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    _render_single_cut_card(result, spec, expanded=True)
    app.caption(
        f"Question {results.index(result) + 1} of {len(results)}. "
        "Use the sidebar to switch questions."
    )


# ---------------------------------------------------------------------------
# Section 4 — Cross cuts
# ---------------------------------------------------------------------------


def _section_cross_cuts() -> None:
    app = _require_streamlit()
    _section_header("4", SECTION_CROSS_CUTS, anchor="section-4", meta="Two-question analyses")

    if not app.session_state["run_complete"]:
        app.info(EMPTY_NO_CROSS_CUTS)
        return

    if app.checkbox(
        "Show suggested cross cuts",
        key="show_suggested_cross_cuts",
        help=TOOLTIP_CROSS_CUT_SUGGESTIONS,
    ):
        with app.expander("Suggested cross cuts", expanded=True):
            _render_suggested_cross_cuts()

    with app.expander("Run a manual cross cut"):
        _render_manual_cross_cut()

    results = app.session_state["cross_cut_results"]
    if not results:
        app.caption("No cross cuts run yet. Build one above to populate this list.")
        return

    import html as _html

    schema = app.session_state.get("schema")
    app.markdown("**Cross-cut results**")
    for result in results:
        with app.expander(
            f"{result.cross_cut_id} \u2014 {result.synthetic_question_title}",
            expanded=False,
        ):
            # Show full source-question text as a header so analysts know
            # exactly which questions feed this cross cut.
            source_lines = []
            for qid in result.source_question_ids:
                q = schema.get_question(qid) if schema is not None else None
                if q and q.question_text:
                    short = q.question_text.strip()
                    if len(short) > 60:
                        short = short[:57] + "\u2026"
                    source_lines.append(
                        f"{_html.escape(qid)}: {_html.escape(short)}"
                    )
                else:
                    source_lines.append(_html.escape(qid))
            app.markdown(
                f'<div style="font-size:13px;font-weight:700;color:#0A0A0A;'
                f'font-family:Arial;margin-bottom:4px;">'
                f'{_html.escape(result.synthetic_question_title)}</div>'
                f'<div style="font-size:10px;color:#888;font-family:Arial;'
                f'line-height:1.8;">{"<br>".join(source_lines)}</div>',
                unsafe_allow_html=True,
            )
            app.caption(
                f"Type: {result.analysis_type.value}  \u00b7  "
                f"Display: {result.display_mode}  \u00b7  "
                f"{len(result.audit_records)} audit records"
            )
            _render_cross_cut_preview(result)
            if result.warnings:
                if app.checkbox(
                    f"Show {len(result.warnings)} warning(s)",
                    key=f"cc_warn_{result.cross_cut_id}",
                ):
                    for warning in result.warnings:
                        app.write(f"\u2022 {warning}")
            cb_col, rm_col = app.columns([3, 1])
            with cb_col:
                app.checkbox(
                    "Include in cross-cut workbook",
                    value=app.session_state.get(
                        f"cc_select_{result.cross_cut_id}", True
                    ),
                    key=f"cc_select_{result.cross_cut_id}",
                )
            with rm_col:
                if app.button(
                    "Remove",
                    key=f"cc_remove_{result.cross_cut_id}",
                    help="Remove this cross cut from the session",
                ):
                    app.session_state["cross_cut_results"] = [
                        r
                        for r in app.session_state["cross_cut_results"]
                        if r.cross_cut_id != result.cross_cut_id
                    ]
                    app.session_state["cross_cut_only_bytes"] = None
                    _refresh_full_workbook()
                    app.rerun()


# ---------------------------------------------------------------------------
# Section 5 — Downloads
# ---------------------------------------------------------------------------


def _section_downloads() -> None:
    app = _require_streamlit()
    _section_header("5", SECTION_DOWNLOADS, anchor="section-5", meta="Workbooks")

    if not app.session_state["run_complete"]:
        app.info("Run an analysis to generate downloadable workbooks.")
        return

    output_path = app.session_state.get("output_path")
    cross_cut_results = app.session_state.get("cross_cut_results", [])
    selected_cross_cuts = [
        r
        for r in cross_cut_results
        if app.session_state.get(f"cc_select_{r.cross_cut_id}", True)
    ]
    filtered_results = app.session_state.get("filtered_results", {})
    selected_filtered = [
        r
        for cid, r in filtered_results.items()
        if app.session_state.get(f"fsc_select_{cid}", True)
    ]

    # Invalidate cached workbook bytes if the selection signature changed
    # (otherwise the user could click Download and get an outdated workbook
    # that was generated against a previous selection).
    cc_signature = tuple(sorted(r.cross_cut_id for r in selected_cross_cuts))
    if app.session_state.get("cross_cut_only_signature") != cc_signature:
        app.session_state["cross_cut_only_bytes"] = None
    fsc_signature = tuple(
        sorted(
            cid
            for cid in filtered_results
            if app.session_state.get(f"fsc_select_{cid}", True)
        )
    )
    if app.session_state.get("filtered_workbook_signature") != fsc_signature:
        app.session_state["filtered_workbook_bytes"] = None

    col_full, col_cc, col_fsc = app.columns(3)

    with col_full:
        app.markdown("**Single-cut workbook**")
        app.caption("All single cuts + cross cuts inline.")
        if output_path and os.path.exists(output_path):
            with open(output_path, "rb") as workbook_file:
                workbook_bytes = workbook_file.read()
            app.download_button(
                label="Download single-cut workbook",
                data=workbook_bytes,
                file_name="survey_analysis.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True,
            )
        else:
            app.button(
                "Download single-cut workbook",
                disabled=True,
                use_container_width=True,
            )

    with col_cc:
        app.markdown("**Cross-cut-only workbook**")
        app.caption("Just the cross cuts you've ticked above.")
        if app.button(
            "Generate cross-cut workbook",
            disabled=(len(selected_cross_cuts) == 0),
            use_container_width=True,
            key="gen_cc_workbook",
            help=(
                f"{len(selected_cross_cuts)} cross cuts selected"
                if selected_cross_cuts
                else "Tick at least one cross cut to enable"
            ),
        ):
            from src.excel_exporter import export_cross_cuts_only

            cc_path = "/tmp/cross_cuts.xlsx"
            try:
                export_cross_cuts_only(
                    cross_cut_results=selected_cross_cuts,
                    schema=app.session_state["schema"],
                    log=app.session_state["log"],
                    output_path=cc_path,
                )
                with open(cc_path, "rb") as f:
                    app.session_state["cross_cut_only_bytes"] = f.read()
                app.session_state["cross_cut_only_signature"] = cc_signature
            except Exception as exc:  # noqa: BLE001
                app.error(f"Cross-cut export failed: {type(exc).__name__}: {exc}")
                with app.expander("Show traceback"):
                    app.code(traceback.format_exc())
        if app.session_state.get("cross_cut_only_bytes"):
            app.download_button(
                label="Download cross-cut workbook",
                data=app.session_state["cross_cut_only_bytes"],
                file_name="cross_cut_analysis.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True,
            )

    with col_fsc:
        app.markdown("**Filtered workbook**")
        app.caption("Per-question filtered analyses you've ticked.")
        if app.button(
            "Generate filtered workbook",
            disabled=(len(selected_filtered) == 0),
            use_container_width=True,
            key="gen_fsc_workbook",
            help=(
                f"{len(selected_filtered)} filtered analyses selected"
                if selected_filtered
                else "Apply at least one per-question filter to enable"
            ),
        ):
            from src.excel_exporter import export_filtered_single_cuts

            fsc_path = "/tmp/filtered_single_cuts.xlsx"
            try:
                export_filtered_single_cuts(
                    filtered_results=selected_filtered,
                    schema=app.session_state["schema"],
                    log=app.session_state["log"],
                    output_path=fsc_path,
                )
                with open(fsc_path, "rb") as f:
                    app.session_state["filtered_workbook_bytes"] = f.read()
                app.session_state["filtered_workbook_signature"] = fsc_signature
            except Exception as exc:  # noqa: BLE001
                app.error(f"Filtered export failed: {type(exc).__name__}: {exc}")
                with app.expander("Show traceback"):
                    app.code(traceback.format_exc())
        if app.session_state.get("filtered_workbook_bytes"):
            app.download_button(
                label="Download filtered workbook",
                data=app.session_state["filtered_workbook_bytes"],
                file_name="filtered_single_cuts.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True,
            )

    app.caption(TOOLTIP_THREE_DOWNLOADS)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main() -> None:
    app = _require_streamlit()
    app.set_page_config(
        page_title=APP_TITLE,
        page_icon="\U0001F4CA",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    _initialise_session_state()
    _inject_theme_css()
    _drain_pending_actions()
    _render_sidebar()

    # Inline title/caption/divider removed in Day 18 — the fixed red header
    # banner injected by _inject_theme_css now serves as the page header.

    gf_error = app.session_state.get("global_filter_error")
    if gf_error:
        app.error(f"Global filter failed: {gf_error}")

    _section_upload()
    app.divider()
    _section_global_filter()
    app.divider()
    _section_single_cuts()
    app.divider()
    _section_cross_cuts()
    app.divider()
    _section_downloads()


if __name__ == "__main__":
    main()
