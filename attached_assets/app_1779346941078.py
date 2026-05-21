"""Streamlit entry point for the Survey Analysis Engine."""

from __future__ import annotations

import html
import os
import sys
import tempfile
import traceback
from pathlib import Path
from typing import Any

import pandas as pd

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

import re as _re
from src.ai_insights import (
    generate_insight,
    generate_outlier_insight,
    generate_table_insight,
)
from src.cross_cut_suggestions import score_suggestions_for_outcome
from src.models import InsightResult, OutcomeSegmentationResult


_INSIGHT_CACHE: dict[str, Any] = {}


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
    "wiz_step": 1,
    "wiz_complete": False,
    "wiz_category_assignments": None,
    "wiz_selected_demographics": None,
    "wiz_num_custom_filters": 2,
    "wiz_num_per_question_filters": 1,
    "wiz_crosscut_consolidation": "one_sheet",
    "ss_search": "",
    "pending_global_filter": None,
    "pending_per_question_filter": {},
    "global_filter_error": None,
    "per_question_filter_errors": {},
    "data_map": None,
    "survey_type_result": None,
    "outcome_variable_id": None,
    "segment_definition": None,
    "segmentation_result": None,
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

def _custom_header_html() -> str:
    app = _require_streamlit()
    schema = app.session_state.get("schema")
    if schema is not None and hasattr(schema, "total_respondents"):
        context = (
            f"{schema.total_respondents:,} respondents loaded"
        )
    else:
        context = "No dataset loaded \u2014 upload to begin"
    return (
        '<div class="custom-header">'
        '<span class="custom-header-title">'
        f'<strong>Survey Analysis Engine</strong> &middot; {html.escape(context)}'
        '</span>'
        '<div class="custom-header-nav" id="custom-header-nav"></div>'
        '</div>'
    )


def _inject_global_css() -> None:
    app = _require_streamlit()

    app.markdown(
        """
    <style>
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }

    .insight-card { display: flex; align-items: flex-start; gap: 10px; padding: 12px 14px; background: #FFF; border: 0.5px solid #E5E5E5; border-left: 3px solid #CC0000; border-radius: 6px; margin-bottom: 8px; font-family: Arial, sans-serif; }
    .insight-icon { display: flex; align-items: center; justify-content: center; width: 24px; height: 24px; background: #FCEBEB; color: #CC0000; border-radius: 4px; font-size: 14px; flex-shrink: 0; margin-top: 1px; }
    .insight-body { flex: 1; min-width: 0; }
    .insight-label { font-size: 10px; font-weight: 500; color: #888; text-transform: uppercase; letter-spacing: 0.08em; margin-bottom: 3px; }
    .insight-text { font-size: 14px; font-weight: 500; line-height: 1.45; color: #1A1A1A; }
    .insight-text .num { color: #CC0000; font-weight: 500; }
    .insight-footer { font-size: 10px; color: #999; margin-top: 4px; }
    .insight-template { background: #F5F5F5; border-left-color: #888; }
    .insight-template .insight-icon { background: #E5E5E5; color: #888; }

    .ui-panel { background: #FFF; border: 0.5px solid #E5E5E5; border-radius: 8px; margin-bottom: 12px; overflow: hidden; }
    .ui-panel-head { display: flex; align-items: center; padding: 9px 14px; border-bottom: 0.5px solid #EEE; font-size: 12px; background: #FAFAFA; }
    .ui-panel-title { font-weight: 500; color: #1A1A1A; }
    .ui-panel-meta { margin-left: auto; color: #888; font-size: 11px; }

    div[data-testid="stDataFrame"] { font-size: 12px; }

    /* Nav bar buttons */
    div[data-testid="stButton"] button {
        border-radius: 6px !important;
        font-size: 12px !important;
        padding: 6px 8px !important;
    }
    div[data-testid="stButton"] button[kind="primary"] {
        background: #CC0000 !important;
        border-color: #CC0000 !important;
        color: #FFF !important;
    }
    div[data-testid="stButton"] button[kind="primary"]:hover {
        background: #B30000 !important;
        border-color: #B30000 !important;
    }

    /* CHANGE B — merged red header with embedded nav tabs */
    .custom-header {
        position: fixed !important;
        top: 0 !important;
        left: 0 !important;
        right: 0 !important;
        z-index: 999999 !important;
        background: #CC0000 !important;
        color: #FFFFFF !important;
        padding: 0 24px !important;
        height: 56px !important;
        display: flex !important;
        align-items: center !important;
        gap: 24px !important;
        border-bottom: 1px solid #A30000 !important;
        font-family: Arial, sans-serif !important;
        box-sizing: border-box !important;
    }
    .custom-header-title {
        font-size: 14px !important;
        font-weight: 500 !important;
        color: #FFF !important;
        flex-shrink: 0 !important;
        white-space: nowrap !important;
    }
    .custom-header-title strong { font-weight: 600 !important; }
    .custom-header-nav {
        display: flex !important;
        align-items: center !important;
        gap: 0 !important;
        margin-left: auto !important;
        height: 100% !important;
    }
    .nav-bar { display: none !important; }
    .nav-tab {
        display: inline-flex !important;
        align-items: center !important;
        gap: 6px !important;
        padding: 0 14px !important;
        height: 56px !important;
        font-size: 13px !important;
        font-family: Arial, sans-serif !important;
        font-weight: 400 !important;
        color: rgba(255,255,255,0.75) !important;
        text-decoration: none !important;
        border: none !important;
        border-bottom: 3px solid transparent !important;
        border-radius: 0 !important;
        background: transparent !important;
        white-space: nowrap !important;
        cursor: pointer !important;
        transition: color 0.15s ease, border-color 0.15s ease, background 0.15s ease !important;
        box-sizing: border-box !important;
        outline: none !important;
    }
    .nav-tab:hover {
        color: #FFFFFF !important;
        background: rgba(255,255,255,0.08) !important;
        border-bottom: 3px solid rgba(255,255,255,0.4) !important;
        text-decoration: none !important;
    }
    .nav-tab.active {
        color: #FFFFFF !important;
        font-weight: 500 !important;
        border-bottom: 3px solid #FFFFFF !important;
        background: rgba(255,255,255,0.08) !important;
    }
    .nav-badge {
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        background: rgba(255,255,255,0.2) !important;
        color: #FFFFFF !important;
        font-size: 10px !important;
        font-weight: 500 !important;
        padding: 1px 6px !important;
        border-radius: 10px !important;
        min-width: 18px !important;
        line-height: 1.4 !important;
    }
    .nav-tab.active .nav-badge {
        background: #FFFFFF !important;
        color: #CC0000 !important;
    }
    .main .block-container {
        padding-top: 80px !important;
        max-width: 100% !important;
    }
    [id^="section-"] {
        scroll-margin-top: 72px !important;
        display: block !important;
    }
    header[data-testid="stHeader"] { display: none !important; }
    </style>
    """,
        unsafe_allow_html=True,
    )
    app.markdown(_THEME_CSS, unsafe_allow_html=True)
    app.markdown(_THEME_CSS_DAY18, unsafe_allow_html=True)
    app.markdown(_custom_header_html(), unsafe_allow_html=True)

def _inject_theme_css() -> None:
    """Backwards-compatible alias retained for any internal callers."""
    _inject_global_css()


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


_ICONS = {
    "analysis": '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3v18h18"/><path d="M7 12l3-3 3 3 5-5"/></svg>',
    "insight": '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 18h6"/><path d="M10 22h4"/><path d="M15.09 14c.18-.98.65-1.74 1.41-2.5A4.65 4.65 0 0 0 18 8 6 6 0 0 0 6 8c0 1 .23 2.23 1.5 3.5A4.61 4.61 0 0 1 8.91 14"/></svg>',
    "outlier": '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>',
    "winner": '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="8" r="6"/><path d="M15.477 12.89 17 22l-5-3-5 3 1.523-9.11"/></svg>',
}


def _icon(name: str, color: str = "currentColor") -> str:
    """Render an inline SVG icon as an HTML string."""
    svg = _ICONS.get(name, "")
    if not svg:
        return ""
    return (
        f'<span style="display:inline-flex;align-items:center;'
        f'color:{color};margin-right:8px;vertical-align:middle;">{svg}</span>'
    )


def _style_outliers(df: Any) -> Any:
    """Per-column outlier styling for st.dataframe Styler.apply(axis=None)."""

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


def _style_outliers_by_row(df: Any) -> Any:
    """Per-row outlier styling — applies _style_outliers logic across each row.

    For row-percent cross-tabs, outliers should be detected within each row
    (since row values sum to 100%). Implemented by transposing, reusing the
    per-column logic, and transposing the result back. Values are scaled to
    0-100 first because _compute_outlier_flags ignores series whose max < 10.
    """
    try:
        scaled = df * 100
        styled_t = _style_outliers(scaled.T)
        return styled_t.T
    except Exception:
        return pd.DataFrame("", index=df.index, columns=df.columns)


def _style_outliers_pct(df: Any) -> Any:
    """Per-column outlier styling for fractional (0-1) percent dataframes.

    Scales values to 0-100 before delegating to _style_outliers so the
    `max < 10` guard in _compute_outlier_flags doesn't silently disable
    highlighting on percent tables.
    """
    try:
        scaled = df * 100
        return _style_outliers(scaled)
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


def _clear_all_cached_insights() -> None:
    """Drop every cached AI insight from session state.

    Called whenever the underlying data context changes (global filter apply,
    pipeline rerun, new upload) so users never see insights stale relative to
    the current data.
    """
    app = _require_streamlit()
    stale = [
        k for k in list(app.session_state.keys())
        if isinstance(k, str) and (
            k.startswith("insight_sc_")
            or k.startswith("insight_cc_")
            or k.startswith("insight_breakdown_")
        )
    ]
    for k in stale:
        app.session_state.pop(k, None)


def _build_insight_payload_single_cut(
    result: Any, spec: Any, filters_applied: list | None = None
) -> dict:
    """Build payload dict for a SingleCutResult (any subtype)."""
    rows: list[dict[str, Any]] = []
    if hasattr(result, "distribution") and getattr(result, "distribution", None):
        for code, payload in result.distribution.items():
            label = (
                spec.option_map.get(code, str(code))
                if getattr(spec, "option_map", None)
                else str(code)
            )
            rows.append(
                {
                    "label": label,
                    "count": payload.get("count", 0) if isinstance(payload, dict) else 0,
                    "rate": payload.get("rate", 0.0) if isinstance(payload, dict) else 0.0,
                }
            )
    elif hasattr(result, "selections") and getattr(result, "selections", None):
        for sub_id, payload in result.selections.items():
            if not isinstance(payload, dict):
                continue
            label = payload.get("label", sub_id) or sub_id
            label_lower = str(label).lower()
            if "unchecked" in label_lower or "not selected" in label_lower:
                continue
            rate = payload.get("selection_rate")
            if rate is None:
                rate = payload.get("rate", 0.0)
            rows.append(
                {
                    "label": label,
                    "count": payload.get("count", 0),
                    "rate": float(rate) if rate is not None else 0.0,
                }
            )
    elif hasattr(result, "rows") and getattr(result, "rows", None):
        for sub_id, row_result in result.rows.items():
            row_label = (
                spec.grid_row_labels.get(sub_id, sub_id)
                if getattr(spec, "grid_row_labels", None)
                else sub_id
            )
            dist = (
                row_result.distribution
                if hasattr(row_result, "distribution")
                else {}
            )
            for code, payload in dist.items():
                if not isinstance(payload, dict):
                    continue
                rows.append(
                    {
                        "label": f"{row_label} \u2014 {payload.get('label', str(code))}",
                        "count": payload.get("count", 0),
                        "rate": payload.get("rate", 0.0),
                    }
                )

    summary: dict[str, Any] = {}
    if hasattr(result, "mean") and result.mean is not None:
        try:
            summary["mean"] = round(float(result.mean), 4)
        except (TypeError, ValueError):
            pass
    if hasattr(result, "median") and result.median is not None:
        try:
            summary["median"] = round(float(result.median), 4)
        except (TypeError, ValueError):
            pass
    if hasattr(result, "std") and result.std is not None:
        try:
            summary["std"] = round(float(result.std), 4)
        except (TypeError, ValueError):
            pass

    # Numeric subtypes have no row-wise distribution; synthesize summary rows
    # so the AI insight pipeline does not fall back to template-only mode.
    if not rows and summary:
        for stat_name, stat_value in summary.items():
            rows.append(
                {"label": stat_name, "count": int(getattr(result, "valid_n", 0) or 0), "rate": stat_value}
            )

    return {
        "table_kind": "single_cut",
        "question_id": spec.canonical_id,
        "question_text": spec.question_text,
        "valid_n": int(getattr(result, "valid_n", 0) or 0),
        "missing_n": int(getattr(result, "missing_n", 0) or 0),
        "filters_applied": filters_applied or [],
        "rows": rows,
        "summary": summary,
    }


def _build_insight_payload_cross_cut(result: Any) -> dict:
    """Build payload dict for a CrossCutResult."""
    rt = result.result_table or {}
    rows: list[dict[str, Any]] = []

    if "counts" in rt:
        row_labels = rt.get("row_label_map", {}) or {}
        col_labels = rt.get("column_label_map", {}) or {}
        row_pcts = rt.get("row_pct", {}) or {}
        for row_code, col_dict in (rt.get("counts", {}) or {}).items():
            if not isinstance(col_dict, dict):
                continue
            for col_code, count in col_dict.items():
                rows.append(
                    {
                        "row_label": row_labels.get(row_code, str(row_code)),
                        "col_label": col_labels.get(col_code, str(col_code)),
                        "count": count,
                        "row_pct": (row_pcts.get(row_code, {}) or {}).get(
                            col_code, 0.0
                        ),
                    }
                )
    elif "per_segment" in rt:
        for seg_val, seg_data in (rt.get("per_segment", {}) or {}).items():
            if not isinstance(seg_data, dict):
                continue
            try:
                mean_v = round(float(seg_data.get("mean", 0) or 0), 4)
            except (TypeError, ValueError):
                mean_v = 0
            try:
                median_v = round(float(seg_data.get("median", 0) or 0), 4)
            except (TypeError, ValueError):
                median_v = 0
            rows.append(
                {
                    "segment_label": seg_data.get("label", str(seg_val)),
                    "n": seg_data.get("n", 0),
                    "mean": mean_v,
                    "median": median_v,
                }
            )
    elif "target_result" in rt:
        tr = rt.get("target_result", {}) or {}
        if "distribution" in tr:
            for code, payload in tr["distribution"].items():
                if not isinstance(payload, dict):
                    continue
                rows.append(
                    {
                        "label": payload.get("label", str(code)),
                        "count": payload.get("count", 0),
                        "rate": payload.get("rate", 0.0),
                    }
                )
        elif "selections" in tr:
            for sub_id, payload in tr["selections"].items():
                if not isinstance(payload, dict):
                    continue
                label = payload.get("label", sub_id) or sub_id
                if "unchecked" in str(label).lower():
                    continue
                rows.append(
                    {
                        "label": label,
                        "count": payload.get("count", 0),
                    }
                )

    title = (
        getattr(result, "synthetic_question_title", None)
        or getattr(result, "title", None)
        or result.cross_cut_id
    )
    filter_expr = getattr(result, "filter_expr", None) or rt.get("filter_expr")
    overall = rt.get("overall") or {}
    valid_n = (
        rt.get("grand_total")
        or rt.get("filter_n")
        or rt.get("paired_n")
        or (overall.get("n") if isinstance(overall, dict) else 0)
        or 0
    )
    return {
        "table_kind": result.analysis_type.value.lower(),
        "question_id": result.cross_cut_id,
        "question_text": title,
        "valid_n": int(valid_n),
        "missing_n": 0,
        "filters_applied": [filter_expr] if filter_expr else [],
        "rows": rows,
        "summary": {},
    }


def _render_insight_section(
    insight_key: str, payload_factory: Any, table_kind: str, title_hint: str
) -> None:
    """Render the AI insight section: button + cached result display."""
    app = _require_streamlit()
    app.divider()
    cols = app.columns([3, 1])
    with cols[0]:
        app.markdown("**AI insight**")
    with cols[1]:
        btn_label = (
            "Generate insight"
            if insight_key not in app.session_state
            else "Regenerate"
        )
        if app.button(
            btn_label, key=f"btn_{insight_key}", use_container_width=True
        ):
            with app.spinner("Generating insight..."):
                from src.ai_insights import generate_insight
                payload = payload_factory()
                ir = generate_insight(
                    payload, table_kind=table_kind, title_hint=title_hint
                )
                app.session_state[insight_key] = ir
            app.rerun()

    if insight_key in app.session_state:
        ir = app.session_state[insight_key]
        import html as _html_mod
        title_html = _html_mod.escape(str(ir.title))
        insight_html = _html_mod.escape(str(ir.insight))
        app.markdown(
            f"<div style='font-size: 14px; font-weight: 500; "
            f"color: #1a1a1a; margin-bottom: 4px;'>{title_html}</div>",
            unsafe_allow_html=True,
        )
        if ir.was_template:
            app.caption(
                f"{ir.insight}  *(template \u2014 AI unavailable)*"
            )
            if ir.error_message:
                with app.expander("Why was AI unavailable?"):
                    app.code(ir.error_message)
        else:
            app.markdown(
                f"<div style='font-size: 13px; line-height: 1.7; "
                f"color: #333; background: #f8f9fa; "
                f"border-left: 3px solid #185FA5; padding: 10px 14px; "
                f"border-radius: 0;'>{insight_html}</div>",
                unsafe_allow_html=True,
            )
            app.caption(
                f"Generated by {ir.model_used} \u00b7 "
                f"{ir.tokens_used} tokens"
            )


def _render_insight_card(
    insight: InsightResult,
    label: str = "Key insight",
    icon: str = "ti-bulb",
) -> None:
    app = _require_streamlit()
    if not insight or not insight.insight:
        return

    headline_raw = insight.insight
    number_pattern = _re.compile(r"(\d+(?:\.\d+)?\s*(?:%|x|X|×)?)")
    escaped = html.escape(headline_raw)
    headline_highlighted = number_pattern.sub(
        r'<span class="num">\1</span>', escaped
    )
    label_html = html.escape(label)

    if insight.was_template:
        app.markdown(
            f"""
        <div class="insight-card insight-template">
            <div class="insight-icon"><i class="ti ti-info-circle"></i></div>
            <div class="insight-body">
                <div class="insight-label">{label_html}</div>
                <div class="insight-text">{html.escape(headline_raw)}</div>
                <div class="insight-footer">Template fallback · AI unavailable</div>
            </div>
        </div>
        """,
            unsafe_allow_html=True,
        )
        return

    model = html.escape(insight.model_used or "")
    footer = f"AI insight · {model}" if model else "AI insight"
    app.markdown(
        f"""
    <div class="insight-card">
        <div class="insight-icon"><i class="ti {icon}"></i></div>
        <div class="insight-body">
            <div class="insight-label">{label_html}</div>
            <div class="insight-text">{headline_highlighted}</div>
            <div class="insight-footer">{footer}</div>
        </div>
    </div>
    """,
        unsafe_allow_html=True,
    )


def _copy_button(df: Any, key: str) -> None:
    """Render a Copy Table button that copies the DataFrame as TSV.

    Tab-separated values paste cleanly into Excel preserving columns
    and rows. Uses navigator.clipboard for one-click copy.
    """
    app = _require_streamlit()
    try:
        if df is None or len(df) == 0:
            return
        lines = []
        headers = list(df.columns)
        if df.index.name:
            headers = [df.index.name] + headers
        lines.append("\t".join(str(h) for h in headers))
        for idx, row in df.iterrows():
            if df.index.name:
                row_vals = [str(idx)] + [str(v) for v in row.values]
            else:
                row_vals = [str(v) for v in row.values]
            lines.append("\t".join(row_vals))
        tsv = "\n".join(lines)
        tsv_escaped = (
            tsv.replace("\\", "\\\\")
            .replace("`", "\\`")
            .replace("$", "\\$")
            .replace("</", "<\\/")
        )
        safe_key = "".join(
            c if (c.isalnum() or c in "_-") else "_" for c in str(key)
        )
        button_id = f"copy_btn_{safe_key}"
        app.markdown(
            f"""
<div style="display:flex;justify-content:flex-end;margin-bottom:4px;">
  <button id="{button_id}"
    onclick="navigator.clipboard.writeText(`{tsv_escaped}`).then(()=>{{
      var b=document.getElementById('{button_id}');
      b.textContent='\u2713 Copied!';b.style.background='#2E7D32';
      setTimeout(()=>{{b.textContent='\u2398 Copy Table';b.style.background='#0A0A0A';}},2000);
    }}).catch(()=>{{document.getElementById('{button_id}').textContent='Copy failed';}});"
    style="background:#0A0A0A;color:white;border:none;padding:6px 14px;
      font-family:Arial,Helvetica,sans-serif;font-size:11px;font-weight:600;
      letter-spacing:0.05em;cursor:pointer;display:flex;align-items:center;
      gap:6px;transition:background 0.15s;">\u2398 Copy Table</button>
</div>
""",
            unsafe_allow_html=True,
        )
    except Exception:
        return


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
        "Show table view (sortable \u00b7 copyable \u00b7 downloadable)",
        key=f"sc_tableview_{key_suffix}" if key_suffix else None,
        value=False,
    )
    if show_table:
        rows_for_df = []
        for i, (code, payload) in enumerate(distribution.items()):
            label = payload.get("label", str(code))
            count = payload.get("count", 0)
            rate = payload.get("rate")
            if rate is None:
                rate = (count / valid_n) if valid_n else 0
            pct = round(rate * 100, 1)
            flag = flags[i] if i < len(flags) else ""
            flag_label = (
                "High outlier" if flag == "high"
                else "Low outlier" if flag == "low"
                else ""
            )
            if display_mode == "Counts":
                rows_for_df.append(
                    {"Label": label, "Count": count, "Flag": flag_label}
                )
            elif display_mode == "Counts + %":
                rows_for_df.append(
                    {"Label": label, "Count": count, "%": pct, "Flag": flag_label}
                )
            else:
                rows_for_df.append(
                    {"Label": label, "%": pct, "Flag": flag_label}
                )
        df_plain = pd.DataFrame(rows_for_df)
        _copy_button(df_plain, f"sc_{key_suffix or id(distribution)}_{display_mode}")
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
            _invalidate_stage_c_state()
            _rerun_single_cuts_on_active_df()
            _clear_all_cached_insights()
            _INSIGHT_CACHE.clear()
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
    from src.ai_insights import (
        categorize_demographic_questions,
        categorize_questions_into_themes,
        generate_short_labels,
    )

    questions_for_themes = [
        {
            "question_id": q.canonical_id,
            "question_text": q.question_text,
            "question_type": q.question_type.value,
            "is_demographic": getattr(q, "is_demographic", False),
        }
        for q in schema.questions
    ]
    themes = categorize_questions_into_themes(
        questions_for_themes,
        cache=_INSIGHT_CACHE,
    )
    schema, themes = _wizard_pipeline_overrides(schema, themes)
    demographic_questions_for_priority = [
        {
            "question_id": q.canonical_id,
            "question_text": q.question_text,
            "question_type": q.question_type.value,
        }
        for q in schema.demographic_questions()
    ]
    demo_priority = categorize_demographic_questions(
        demographic_questions_for_priority,
        cache=_INSIGHT_CACHE,
    )
    short_labels = generate_short_labels(
        [
            {"question_id": q.canonical_id, "question_text": q.question_text}
            for q in schema.questions
        ],
        cache=_INSIGHT_CACHE,
    )
    export_single_cuts(
        results=results,
        skips=skips,
        schema=schema,
        quality_report=quality_report,
        log=log,
        output_path=output_path,
        themes=themes,
        decoded_df=dataframe,
        demo_priority=demo_priority,
        short_labels=short_labels,
    )

    app = _require_streamlit()
    _INSIGHT_CACHE.clear()
    app.session_state["decoded_df"] = dataframe
    app.session_state["active_df"] = dataframe
    app.session_state["data_map"] = data_map
    app.session_state["survey_type_result"] = None
    app.session_state["outcome_variable_id"] = None
    app.session_state.pop("outcome_variable_selector", None)
    _invalidate_stage_c_state()
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
    _clear_all_cached_insights()
    app.session_state["run_complete"] = True
    status.update(label="Analysis complete.", state="complete")


def _refresh_full_workbook() -> None:
    from src.ai_insights import (
        categorize_demographic_questions,
        categorize_questions_into_themes,
        generate_short_labels,
    )
    from src.excel_exporter import export_single_cuts

    app = _require_streamlit()
    schema = app.session_state["schema"]
    questions_for_themes = [
        {
            "question_id": q.canonical_id,
            "question_text": q.question_text,
            "question_type": q.question_type.value,
            "is_demographic": getattr(q, "is_demographic", False),
        }
        for q in schema.questions
    ]
    themes = categorize_questions_into_themes(
        questions_for_themes,
        cache=_INSIGHT_CACHE,
    )
    demographic_questions_for_priority = [
        {
            "question_id": q.canonical_id,
            "question_text": q.question_text,
            "question_type": q.question_type.value,
        }
        for q in schema.demographic_questions()
    ]
    demo_priority = categorize_demographic_questions(
        demographic_questions_for_priority,
        cache=_INSIGHT_CACHE,
    )
    short_labels = generate_short_labels(
        [
            {"question_id": q.canonical_id, "question_text": q.question_text}
            for q in schema.questions
        ],
        cache=_INSIGHT_CACHE,
    )
    export_single_cuts(
        results=app.session_state["results"],
        skips=app.session_state["skips"],
        schema=schema,
        quality_report=app.session_state["quality_report"],
        log=app.session_state["log"],
        output_path=app.session_state["output_path"],
        cross_cut_results=app.session_state["cross_cut_results"],
        cross_cut_skips=app.session_state["cross_cut_skips"],
        themes=themes,
        decoded_df=app.session_state.get("decoded_df"),
        demo_priority=demo_priority,
        short_labels=short_labels,
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


_STAGE_C_WIDGET_KEYS = (
    "seg_mode_radio",
    "winner_values_multiselect",
    "winner_label_input",
    "loser_label_input",
    "numeric_threshold_input",
    "threshold_direction_radio",
    "winner_label_num_input",
    "loser_label_num_input",
)


def _invalidate_stage_c_state() -> None:
    """Clear Stage C segment definition + result + widget keys.

    Called whenever the underlying data context changes (new upload, global
    filter applied/cleared, outcome variable swapped) so the displayed
    segmentation result always matches the current context.
    """
    app = _require_streamlit()
    app.session_state["segment_definition"] = None
    app.session_state["segmentation_result"] = None
    for key in _STAGE_C_WIDGET_KEYS:
        app.session_state.pop(key, None)


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
    _clear_all_cached_insights()
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
    _copy_button(df, f"ct_{result.cross_cut_id}_{display_mode}")
    if display_mode == "Counts":
        _styled_dataframe(df, use_container_width=True)
    else:
        # Row %: outliers per row (rows sum to 100%); Column %: per column.
        # Both use scaled (0-100) helpers so the outlier detector's
        # `max < 10` guard doesn't suppress highlights on fractional values.
        styler_fn = (
            _style_outliers_by_row if display_mode == "Row %" else _style_outliers_pct
        )
        try:
            app.dataframe(
                df.style.apply(styler_fn, axis=None).format("{:.1%}"),
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
        _df = pd.DataFrame(rows)
        _copy_button(_df, f"sp_dist_{result.cross_cut_id}")
        _styled_dataframe(_df, use_container_width=True, hide_index=True)
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
        _df = pd.DataFrame(rows)
        _copy_button(_df, f"sp_sel_{result.cross_cut_id}")
        _styled_dataframe(_df, use_container_width=True, hide_index=True)
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
        _copy_button(df, f"sp_num_{result.cross_cut_id}")
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
        _df = pd.DataFrame(grid_rows)
        _copy_button(_df, f"sp_grid_{result.cross_cut_id}")
        _styled_dataframe(_df, use_container_width=True, hide_index=True)
    else:
        app.info("Preview not available for this target type.")


def _preview_group_comparison(result: Any) -> None:
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
    _df = pd.DataFrame(rows)
    _copy_button(_df, f"gc_{result.cross_cut_id}")
    _styled_dataframe(_df, use_container_width=True, hide_index=True)
    app.caption("Group means in the downloaded workbook.")


def _preview_expected_vs_realized(result: Any) -> None:
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
    _copy_button(df, f"evr_{result.cross_cut_id}")
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


def _build_grid_display_rows(
    result: Any, schema: Any
) -> list[dict[str, Any]]:
    """Build UI display rows for a binary grid result.

    Returns one ``{"Option": label, "Count": count}`` dict per sub-question,
    using the schema's ``grid_row_labels`` for display labels and the
    highest-coded distribution entry (typically value=1, "Checked") for
    counts. Sub-questions with a zero count are omitted.
    """
    spec = next(
        (q for q in schema.questions if q.canonical_id == result.question_id),
        None,
    )
    grid_row_labels = getattr(spec, "grid_row_labels", None) or {}

    rows: list[dict[str, Any]] = []
    for sub_id, row_result in result.rows.items():
        label = grid_row_labels.get(sub_id, sub_id)
        dist = getattr(row_result, "distribution", {}) or {}
        if 1 in dist:
            count = dist[1].get("count", 0)
        elif dist:
            max_code = max(dist.keys(), key=lambda k: str(k))
            count = dist[max_code].get("count", 0)
        else:
            count = 0
        if not count:
            continue
        rows.append({"Option": label, "Count": count})
    return rows


def _render_single_cut_result(result: Any, spec: Any) -> None:
    """Display a SingleCutResult with branded HTML for SS/MS distributions."""
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
        _grid_df = pd.DataFrame(grid_rows)
        _copy_button(_grid_df, f"grid_{result.question_id}")
        _styled_dataframe(_grid_df, use_container_width=True, hide_index=True)
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
        _num_df = pd.DataFrame(rows)
        _copy_button(_num_df, f"num_{result.question_id}")
        app.dataframe(_num_df, use_container_width=True, hide_index=True)
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
    pending_insight: dict[str, Any] | None = None
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
                    widget_key = f"{filter_key}_v_{i}_{picked_q_id}"
                    if widget_key not in app.session_state:
                        app.session_state[widget_key] = default_codes
                    else:
                        app.session_state[widget_key] = [
                            v for v in app.session_state[widget_key]
                            if v in value_codes
                        ]
                    v_pick = app.multiselect(
                        "Values (leave empty for breakdown)",
                        options=value_codes,
                        format_func=lambda v: f"{v}: {q_spec.option_map[v]}",
                        key=widget_key,
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
                # Filter context changed — drop stale cached insights for this card.
                for k in (
                    f"insight_sc_{spec.canonical_id}",
                    f"insight_breakdown_{spec.canonical_id}",
                ):
                    app.session_state.pop(k, None)
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
            active_filter_strs = [
                _format_filter(f) for f in filtered.filters_applied
            ]
            if filtered.dispatch_mode == "single_cut_filtered":
                _render_single_cut_result(filtered.single_cut_result, spec)
                pending_insight = {
                    "insight_key": f"insight_sc_{spec.canonical_id}",
                    "payload_factory": lambda: _build_insight_payload_single_cut(
                        filtered.single_cut_result, spec, active_filter_strs
                    ),
                    "table_kind": "filtered_single_cut",
                    "title_hint": spec.question_text,
                }
            elif filtered.dispatch_mode == "cross_cut_breakdown":
                _render_cross_cut_preview(filtered.cross_cut_result)
                pending_insight = {
                    "insight_key": f"insight_breakdown_{spec.canonical_id}",
                    "payload_factory": lambda: _build_insight_payload_cross_cut(
                        filtered.cross_cut_result
                    ),
                    "table_kind": filtered.cross_cut_result.analysis_type.value.lower(),
                    "title_hint": spec.question_text,
                }
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
                for k in (
                    f"insight_sc_{spec.canonical_id}",
                    f"insight_breakdown_{spec.canonical_id}",
                ):
                    app.session_state.pop(k, None)
                app.rerun()
        else:
            _render_single_cut_result(result, spec)
            pending_insight = {
                "insight_key": f"insight_sc_{spec.canonical_id}",
                "payload_factory": lambda: _build_insight_payload_single_cut(
                    result, spec, []
                ),
                "table_kind": "single_cut",
                "title_hint": spec.question_text,
            }

    # Render insight section OUTSIDE the expander above. Streamlit forbids
    # nesting expanders, and _render_insight_section uses an expander to show
    # the API error detail when the AI call fails.
    if pending_insight is not None:
        # Auto-generate insight when a per-question filter is active;
        # otherwise keep the manual "Generate insight" button.
        if pending_insight["table_kind"] != "single_cut":
            with app.spinner("Generating insight..."):
                auto_payload = pending_insight["payload_factory"]()
                auto_insight = generate_insight(
                    table_payload=auto_payload,
                    table_kind=pending_insight["table_kind"],
                    title_hint=str(pending_insight["title_hint"]),
                    cache=_INSIGHT_CACHE,
                )
            _render_insight_card(auto_insight)
        else:
            _render_insight_section(**pending_insight)


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


def _section_upload() -> None:
    app = _require_streamlit()
    app.markdown('<div id="section-upload"></div>', unsafe_allow_html=True)
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

    # -- Setup wizard (appears after upload, before Run Analysis) --
    if uploaded_files and not app.session_state.get("wiz_complete") and not app.session_state.get("run_complete"):
        _render_setup_wizard(uploaded_files)

    ready = bool(uploaded_files) and not docx_only
    centre_left, centre_mid, centre_right = app.columns([2, 3, 2])
    with centre_mid:
        run_clicked = app.button(
            "Run analysis",
            type="primary",
            disabled=not ready,
            use_container_width=True,
        )

    if app.session_state.pop("_wizard_run_requested", False):
        run_clicked = True

    if run_clicked:
        import logging
        logging.warning("RUN_CLICKED uploaded=%s docx_only=%s", bool(uploaded_files), docx_only)
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
    _invalidate_stage_c_state()
    _rerun_single_cuts_on_active_df()
    _INSIGHT_CACHE.clear()
    app.rerun()


def _section_global_filter() -> None:
    app = _require_streamlit()
    app.markdown('<div id="section-filter"></div>', unsafe_allow_html=True)
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
                widget_key = f"gf_v_{i}_{picked_q_id}"
                if widget_key not in app.session_state:
                    app.session_state[widget_key] = default_codes
                else:
                    app.session_state[widget_key] = [
                        v for v in app.session_state[widget_key]
                        if v in value_codes
                    ]
                v_pick = app.multiselect(
                    "Values (select one or more)",
                    options=value_codes,
                    format_func=lambda v: f"{v}: {q_spec.option_map[v]}",
                    key=widget_key,
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
# Section 2.5 — Survey type classification + outcome variable selector
# ---------------------------------------------------------------------------


def _section_survey_classification() -> None:
    """Detect survey type after first run and let the user pick an outcome variable.

    Detection is deterministic and cached in ``session_state.survey_type_result``;
    it only runs once per uploaded survey. The outcome variable defaults to the
    detector's top suggestion but can be overridden via the dropdown below.
    """
    app = _require_streamlit()
    if not app.session_state.get("run_complete"):
        return

    data_map = app.session_state.get("data_map")
    if data_map is None:
        return

    if app.session_state.get("survey_type_result") is None:
        with app.spinner("Detecting survey type..."):
            from src.survey_type_detector import detect_survey_type

            try:
                survey_type_result = detect_survey_type(
                    schema=data_map,
                    decoded_df=app.session_state["decoded_df"],
                )
            except Exception as exc:  # noqa: BLE001 - never block the UI on detection.
                app.warning(f"Survey type detection failed: {type(exc).__name__}: {exc}")
                return

            # Penalise imbalanced outcome candidates so they sort lower in
            # the dropdown and don't get auto-selected as the primary
            # outcome. The detector itself is data-distribution-blind.
            from dataclasses import replace

            decoded_df = app.session_state["decoded_df"]

            def _imbalance_penalty(opt) -> float:
                if opt.question_id not in decoded_df.columns:
                    return 0.0
                col = decoded_df[opt.question_id].dropna()
                if len(col) == 0:
                    return 0.0
                top_pct = col.value_counts(normalize=True).iloc[0]
                if top_pct >= 0.95:
                    return 0.4
                if top_pct >= 0.90:
                    return 0.2
                return 0.0

            # Filter out demographic/metadata questions when picking the
            # AUTO-SELECTED default. The full all_eligible_questions list
            # (including demographics) is preserved in the dropdown so the
            # user can still manually choose one if they want.
            classified_schema = app.session_state.get("schema")

            def _is_demographic_or_metadata(opt) -> bool:
                if classified_schema is None:
                    return False
                spec = classified_schema.get_question(opt.question_id)
                if spec is None:
                    return False
                return spec.is_demographic or spec.question_type.value in (
                    "METADATA_OR_ID",
                    "OPEN_TEXT",
                    "UNKNOWN",
                )

            eligible_for_outcome = [
                opt
                for opt in survey_type_result.all_eligible_questions
                if not _is_demographic_or_metadata(opt)
            ]

            rebalanced = sorted(
                eligible_for_outcome,
                key=lambda opt: (
                    -(opt.relevance_score - _imbalance_penalty(opt)),
                    _imbalance_penalty(opt),
                    opt.question_id,
                ),
            )
            best_id = rebalanced[0].question_id if rebalanced else None
            survey_type_result = replace(
                survey_type_result,
                outcome_question_id=best_id,
            )

            app.session_state["survey_type_result"] = survey_type_result
            if app.session_state.get("outcome_variable_id") is None:
                app.session_state["outcome_variable_id"] = (
                    survey_type_result.outcome_question_id
                )

    result = app.session_state["survey_type_result"]
    app.markdown("### Outcome Variable Selection")
    app.info(
        "Select the primary outcome variable for segmentation analysis. "
        "Auto-selected based on survey type; you can override."
    )

    all_options = result.all_eligible_questions
    if not all_options:
        app.warning(
            "No measurable questions detected. Segmentation analysis will be unavailable."
        )
        app.divider()
        return

    decoded_df_for_labels = app.session_state.get("decoded_df")

    def _outcome_label(opt) -> str:
        base = f"{opt.question_id}: {opt.question_text}"
        score = f"score: {opt.relevance_score:.2f}"
        if (
            decoded_df_for_labels is not None
            and opt.question_id in decoded_df_for_labels.columns
        ):
            col = decoded_df_for_labels[opt.question_id].dropna()
            if len(col) > 0:
                top_val_pct = col.value_counts(normalize=True).iloc[0]
                if top_val_pct >= 0.90:
                    return (
                        f"{base} ({score} \u2014 "
                        f"{top_val_pct:.0%} single value, likely too imbalanced)"
                    )
        return f"{base} ({score})"

    option_labels: list[str] = []
    option_ids: list[str | None] = []
    for opt in all_options:
        option_labels.append(_outcome_label(opt))
        option_ids.append(opt.question_id)
    option_labels.append("None \u2014 no outcome variable")
    option_ids.append(None)

    current_outcome = app.session_state.get("outcome_variable_id")
    default_index = (
        option_ids.index(current_outcome) if current_outcome in option_ids else 0
    )

    selected_label = app.selectbox(
        "Outcome Variable",
        options=option_labels,
        index=default_index,
        help="Auto-selected based on survey type. Override if needed.",
        key="outcome_variable_selector",
    )
    selected_id = option_ids[option_labels.index(selected_label)]
    if selected_id != app.session_state.get("outcome_variable_id"):
        app.session_state["outcome_variable_id"] = selected_id
        # Outcome variable changed — invalidate stale Stage C state so the
        # results displayed always match the currently selected outcome.
        _invalidate_stage_c_state()
        app.rerun()

    if selected_id is not None:
        selected_opt = next(
            opt for opt in all_options if opt.question_id == selected_id
        )
        app.success(
            f"**Selected:** {selected_opt.question_id} \u2014 {selected_opt.reason}"
        )
    else:
        app.warning(
            "No outcome variable selected. Segmentation analysis will be "
            "unavailable in later stages."
        )

    if len(result.candidate_outcome_questions) > 1:
        with app.expander(
            f"Top {len(result.candidate_outcome_questions)} "
            "Outcome Candidates (auto-ranked)"
        ):
            for opt in result.candidate_outcome_questions:
                app.markdown(f"**{opt.question_id}** \u2014 {opt.question_text}")
                app.caption(
                    f"Relevance: {opt.relevance_score:.2f} | {opt.reason}"
                )
                app.markdown("")

    _render_segment_definition_ui()
    _render_segmentation_results()

    app.divider()


def _render_segment_definition_ui() -> None:
    """Stage C: define winner/loser segments and run outcome segmentation."""
    app = _require_streamlit()
    outcome_id = app.session_state.get("outcome_variable_id")
    if not outcome_id:
        return

    schema = app.session_state.get("schema")
    if schema is None:
        return
    outcome_spec = schema.get_question(outcome_id)
    if outcome_spec is None:
        app.error(f"Outcome question {outcome_id} not found in schema.")
        return

    from src.models import SegmentDefinition

    app.markdown("---")
    app.markdown("### Segment Definition")
    app.info(
        "Define what 'winner' means for your outcome variable. This splits "
        "respondents into two groups for segmentation analysis."
    )

    _SEG_MODE_LABELS = {
        "categorical": "Categorical (select winner values)",
        "numeric_threshold": "Numeric threshold (above/below a value)",
        "quartile": "Quartile (top 25% = Winners vs bottom 25% = Laggards)",
    }
    seg_mode = app.radio(
        "Segmentation mode",
        options=["categorical", "numeric_threshold", "quartile"],
        format_func=lambda x: _SEG_MODE_LABELS[x],
        key="seg_mode_radio",
        horizontal=True,
    )

    segment_definition = None

    if seg_mode == "categorical":
        options = list(outcome_spec.option_map.items())
        if not options:
            app.warning(
                "No option codes found for this question. Try numeric threshold mode."
            )
        else:
            option_labels = [f"{code}: {label}" for code, label in options]
            selected_labels = app.multiselect(
                "Select winner values",
                options=option_labels,
                help=(
                    "Respondents matching these codes are 'winners'. "
                    "All others are 'losers'."
                ),
                key="winner_values_multiselect",
            )
            winner_codes = tuple(
                code
                for code, label in options
                if f"{code}: {label}" in selected_labels
            )
            if winner_codes:
                winner_label = app.text_input(
                    "Winner label", value="Winner", key="winner_label_input"
                )
                loser_label = app.text_input(
                    "Laggard label", value="Laggard", key="loser_label_input"
                )
                segment_definition = SegmentDefinition(
                    outcome_question_id=outcome_id,
                    segment_mode="categorical",
                    winner_values=winner_codes,
                    winner_label=winner_label or "Winner",
                    loser_label=loser_label or "Laggard",
                )
            else:
                app.warning("Select at least one winner value to continue.")
    elif seg_mode == "numeric_threshold":
        threshold = app.number_input(
            "Threshold value",
            value=50.0,
            key="numeric_threshold_input",
        )
        direction = app.radio(
            "Winners are respondents who scored",
            options=["gte", "lte"],
            format_func=lambda x: (
                f"\u2265 {threshold} (at or above threshold)"
                if x == "gte"
                else f"\u2264 {threshold} (at or below threshold)"
            ),
            key="threshold_direction_radio",
            horizontal=True,
        )
        winner_label = app.text_input(
            "Winner label", value="High", key="winner_label_num_input"
        )
        loser_label = app.text_input(
            "Loser label", value="Low", key="loser_label_num_input"
        )
        segment_definition = SegmentDefinition(
            outcome_question_id=outcome_id,
            segment_mode="numeric_threshold",
            winner_threshold=float(threshold),
            threshold_direction=direction,
            winner_label=winner_label or "High",
            loser_label=loser_label or "Low",
        )
    else:
        quartile_winner = app.radio(
            "Winner quartile",
            options=["top", "bottom"],
            format_func=lambda value: (
                "Top quartile" if value == "top" else "Bottom quartile"
            ),
            key=f"quartile_winner_{outcome_id}",
            horizontal=True,
        )
        segment_definition = SegmentDefinition(
            outcome_question_id=outcome_id,
            segment_mode="quartile",
            quartile_winner=quartile_winner,
        )

    if segment_definition is None:
        return

    app.session_state["segment_definition"] = segment_definition

    if app.button(
        "\u25B6 Run Outcome Segmentation",
        key="run_segmentation_btn",
        type="primary",
    ):
        from src.outcome_segmentation import compute_outcome_segmentation

        # compute_outcome_segmentation expects a list[AuditRecord] it can
        # `.append()` to. session_state["log"] is a CalculationLog (record-only),
        # so we collect into a local list and forward into the log afterwards.
        local_audit: list = []
        with app.spinner("Running outcome segmentation..."):
            try:
                seg_result = compute_outcome_segmentation(
                    decoded_df=app.session_state["active_df"],
                    schema=schema,
                    outcome_question_id=outcome_id,
                    segment_definition=segment_definition,
                    audit_log=local_audit,
                    min_sample_size=30,
                )
            except Exception as exc:  # noqa: BLE001 - surface error to UI
                app.error(f"Segmentation failed: {type(exc).__name__}: {exc}")
                return
            calc_log = app.session_state.get("log")
            if calc_log is not None:
                for audit in local_audit:
                    calc_log.record(audit)
            app.session_state["segmentation_result"] = seg_result
            app.rerun()


def _render_segmentation_results() -> None:
    """Stage C: display winner/loser metrics, differentiators, and winner profile."""
    app = _require_streamlit()
    seg = app.session_state.get("segmentation_result")
    if seg is None:
        return

    app.markdown("---")
    app.markdown("### Segmentation Results")

    _winner_lbl = seg.segment_definition.winner_label
    _loser_lbl = seg.segment_definition.loser_label
    col1, col2, col3 = app.columns(3)
    with col1:
        app.metric(f"{_winner_lbl}s", seg.winner_n)
    with col2:
        app.metric(f"{_loser_lbl}s", seg.loser_n)
    with col3:
        app.metric("Differentiators Found", len(seg.differentiators))

    for warning in seg.warnings:
        app.warning(warning)

    if seg.differentiators:
        app.markdown("#### Top Differentiators (ranked by Cram\u00e9r's V)")
        diff_data = [
            {
                "Question": f"{diff.question_id}: {diff.question_text[:80]}",
                "Top Option": diff.top_option_label,
                "Cram\u00e9r's V": f"{diff.cramers_v:.3f}",
                f"{_winner_lbl} Rate": f"{diff.top_option_winner_rate:.1%}",
                f"{_loser_lbl} Rate": f"{diff.top_option_loser_rate:.1%}",
                "Lift": (
                    f"{diff.top_option_lift:.2f}x"
                    if diff.top_option_lift < 900
                    else "\u221E"
                ),
                "p-value": (
                    f"{diff.p_value:.3f}" if diff.p_value is not None else "N/A"
                ),
                "Notes": " | ".join(diff.warnings) if diff.warnings else "",
            }
            for diff in seg.differentiators[:20]
        ]
        app.dataframe(diff_data, use_container_width=True)

        infinite_lift_diffs = [
            diff
            for diff in seg.differentiators[:20]
            if diff.top_option_lift >= 999.0
        ]
        if infinite_lift_diffs:
            app.caption(
                f"{len(infinite_lift_diffs)} question(s) show "
                "infinite lift (\u221E) \u2014 the loser segment had 0 "
                "respondents select this option. Interpret with caution; "
                "check sample sizes."
            )

    if seg.winner_profile.defining_traits:
        app.markdown(f"#### {seg.winner_profile.winner_label} Profile")
        app.caption(
            "Composite archetype built from top "
            f"{len(seg.winner_profile.defining_traits)} differentiating traits"
        )
        for trait in seg.winner_profile.defining_traits:
            tcol1, tcol2 = app.columns([3, 1])
            with tcol1:
                app.markdown(f"**{trait.question_id}:** {trait.option_label}")
                app.caption(trait.question_text)
            with tcol2:
                app.metric(
                    f"{_winner_lbl} Rate",
                    f"{trait.winner_rate:.1%}",
                    delta=f"+{trait.rate_gap:.1%} vs {_loser_lbl.lower()}s",
                )
            matching_diff = next(
                (
                    d
                    for d in seg.differentiators
                    if d.question_id == trait.question_id
                ),
                None,
            )
            if matching_diff and matching_diff.warnings:
                app.caption(
                    f"{' | '.join(matching_diff.warnings)}"
                )

    if seg.skipped_questions:
        with app.expander(
            f"{len(seg.skipped_questions)} questions skipped"
        ):
            for qid, reason in seg.skipped_questions:
                app.caption(f"\u2022 {qid}: {reason}")


# ---------------------------------------------------------------------------
# Section 3 — Single cuts
# ---------------------------------------------------------------------------


def _section_single_cuts() -> None:
    """Day 18 master-detail: show only the selected question's analysis."""
    import html as _html

    app = _require_streamlit()
    app.markdown('<div id="section-singlecuts"></div>', unsafe_allow_html=True)
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


def _source_question_labels(
    question_ids: tuple[str, ...],
    schema: Any,
    *,
    include_ids: bool = True,
    max_chars: int = 60,
) -> str:
    parts: list[str] = []
    for question_id in question_ids:
        spec = schema.get_question(question_id) if schema is not None else None
        if spec is None:
            parts.append(question_id)
            continue
        text = spec.question_text[:max_chars]
        if len(spec.question_text) > max_chars:
            text += "..."
        parts.append(f"{question_id}: {text}" if include_ids else text)
    return " \u00d7 ".join(parts)


def _suggestion_label(suggestion: Any, schema: Any) -> str:
    if schema is None:
        return suggestion.synthetic_question_title
    return _source_question_labels(suggestion.source_question_ids, schema, max_chars=60)


def _cross_cut_display_title(result: Any, schema: Any) -> str:
    source_question_ids = getattr(result, "source_question_ids", ())
    if schema is None or not source_question_ids:
        return getattr(
            result,
            "synthetic_question_title",
            getattr(result, "cross_cut_id", str(result)),
        )
    return _source_question_labels(
        source_question_ids, schema, include_ids=False, max_chars=50
    )


def _section_cross_cuts() -> None:
    app = _require_streamlit()
    app.markdown('<div id="section-crosscuts"></div>', unsafe_allow_html=True)
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
        cc_pending_insight: dict[str, Any] | None = None
        with app.expander(
            f"{result.cross_cut_id} \u2014 {_cross_cut_display_title(result, schema)}",
            expanded=False,
        ):
            # Show full source-question text as a header so analysts know
            # exactly which questions feed this cross cut.
            source_lines = []
            for qid in result.source_question_ids:
                q = schema.get_question(qid) if schema is not None else None
                if q and q.question_text:
                    source_lines.append(
                        f"{_html.escape(qid)}: {_html.escape(q.question_text.strip())}"
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
            cc_pending_insight = {
                "insight_key": f"insight_cc_{result.cross_cut_id}",
                "payload_factory": (
                    lambda r=result: _build_insight_payload_cross_cut(r)
                ),
                "table_kind": result.analysis_type.value.lower(),
                "title_hint": getattr(result, "synthetic_question_title", "")
                or result.cross_cut_id,
            }
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
                    app.session_state.pop(
                        f"insight_cc_{result.cross_cut_id}", None
                    )
                    app.session_state["cross_cut_only_bytes"] = None
                    _refresh_full_workbook()
                    app.rerun()

        # Render insight section OUTSIDE the per-card expander above.
        # Streamlit forbids nesting expanders, and _render_insight_section
        # uses an expander to show the API error detail when the AI call fails.
        if cc_pending_insight is not None:
            _render_insight_section(**cc_pending_insight)


# ---------------------------------------------------------------------------
# Section 5 — Downloads
# ---------------------------------------------------------------------------


def _section_downloads() -> None:
    app = _require_streamlit()
    app.markdown('<div id="section-downloads"></div>', unsafe_allow_html=True)
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

    # --- TEMP DEBUG (remove after diagnosing BCN download/filter bugs) ---
    app.write({
        "DEBUG_run_complete": app.session_state.get("run_complete"),
        "DEBUG_results_count": len(app.session_state.get("results", []) or []),
        "DEBUG_skips_count": len(app.session_state.get("skips", []) or []),
        "DEBUG_active_df_shape": getattr(app.session_state.get("active_df"), "shape", None),
        "DEBUG_decoded_df_shape": getattr(app.session_state.get("decoded_df"), "shape", None),
        "DEBUG_schema_present": app.session_state.get("schema") is not None,
        "DEBUG_output_path": app.session_state.get("output_path"),
        "DEBUG_output_path_exists": (
            os.path.exists(app.session_state.get("output_path"))
            if app.session_state.get("output_path") else False
        ),
        "DEBUG_output_path_size_mib": (
            round(os.path.getsize(app.session_state.get("output_path")) / (1024 * 1024), 2)
            if app.session_state.get("output_path")
            and os.path.exists(app.session_state.get("output_path"))
            else None
        ),
        "DEBUG_global_filter_state": str(app.session_state.get("global_filter_state")),
    })
    # --- END TEMP DEBUG ---

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
# Stage D: AI Analysis section
# ---------------------------------------------------------------------------


def _truncate(text: str, max_chars: int) -> str:
    if len(text) <= max_chars:
        return text
    return text[: max_chars - 3] + "..."


def _bar_colors(values: list[float]) -> list[str]:
    """Return color list with max-value bar in red, others black."""
    if not values:
        return []
    max_val = max(values)
    return ["#CC0000" if v == max_val else "#1A1A1A" for v in values]


def _insight_rows_from_payload(
    payload: dict[str, Any],
    table_kind: str,
) -> list[dict[str, Any]]:
    if table_kind == "differentiator":
        return [
            {
                "label": payload.get("top_option", ""),
                "winner_rate": payload.get("winner_rate", 0),
                "loser_rate": payload.get("loser_rate", 0),
                "lift": payload.get("lift", 0),
                "cramers_v": payload.get("cramers_v", 0),
            }
        ]
    if table_kind == "winner_profile":
        return list(payload.get("traits", []))
    return []


def _normalise_insight_payload(
    table_payload: dict[str, Any],
    table_kind: str,
    title_hint: str,
) -> dict[str, Any]:
    payload = dict(table_payload)
    payload.setdefault("table_kind", table_kind)
    payload.setdefault("question_id", payload.get("question_id", ""))
    payload.setdefault("question_text", payload.get("question_text", title_hint))
    payload.setdefault("valid_n", payload.get("winner_n", 0))
    payload.setdefault("missing_n", 0)
    payload.setdefault("filters_applied", [])
    payload.setdefault("summary", {})
    payload.setdefault("rows", _insight_rows_from_payload(payload, table_kind))
    return payload


def _section_ai_analysis() -> None:
    app = _require_streamlit()
    _section_survey_classification()
    app.divider()
    app.markdown('<div id="section-ai"></div>', unsafe_allow_html=True)
    if not app.session_state.get("run_complete"):
        return

    app.markdown("---")
    app.markdown(
        f'<h2 style="display:flex;align-items:center;margin:1rem 0;">'
        f'{_icon("analysis", "#CC0000")}AI Analysis</h2>',
        unsafe_allow_html=True,
    )

    seg: OutcomeSegmentationResult | None = app.session_state.get(
        "segmentation_result"
    )
    if seg is None:
        app.info(
            "Complete outcome segmentation in the Survey Classification "
            "section above to unlock AI Analysis."
        )
        return

    _render_outcome_summary_panel(seg)
    _render_smart_cross_cut_suggestions_panel(seg)
    _render_winner_profile_panel(seg)


def _render_outcome_summary_panel(seg: OutcomeSegmentationResult) -> None:
    app = _require_streamlit()

    if not seg.differentiators:
        app.info("No differentiators found. Try a different segment definition.")
        return

    laggard_label = seg.segment_definition.loser_label
    winner_label = seg.segment_definition.winner_label

    table_payload = {
        "outcome_question_id": seg.outcome_question_id,
        "winner_label": winner_label,
        "loser_label": laggard_label,
        "winner_n": seg.winner_n,
        "loser_n": seg.loser_n,
        "differentiators": [
            {
                "question_text": d.question_text,
                "top_option_label": d.top_option_label,
                "winner_rate": d.top_option_winner_rate,
                "loser_rate": d.top_option_loser_rate,
                "lift": d.top_option_lift,
                "cramers_v": d.cramers_v,
            }
            for d in seg.differentiators
        ],
    }

    with app.spinner("Generating key insight..."):
        try:
            table_insight = generate_table_insight(
                table_payload=table_payload,
                table_kind="differentiator_table",
                cache=_INSIGHT_CACHE,
            )
            _render_insight_card(table_insight, label="Key insight", icon="ti-bulb")
        except Exception as exc:  # noqa: BLE001
            app.error(f"Key insight failed: {exc}")

    with app.spinner("Identifying outlier..."):
        try:
            outlier_insight = generate_outlier_insight(
                table_payload=table_payload,
                table_kind="outlier",
                cache=_INSIGHT_CACHE,
            )
            _render_insight_card(
                outlier_insight, label="Outlier", icon="ti-alert-triangle"
            )
        except Exception as exc:  # noqa: BLE001
            app.error(f"Outlier insight failed: {exc}")

    avg_lift_values = [
        d.top_option_lift for d in seg.differentiators[:5] if d.top_option_lift < 900
    ]
    avg_lift = (
        sum(avg_lift_values) / len(avg_lift_values) if avg_lift_values else 0.0
    )

    segment_mode = str(getattr(seg.segment_definition, "segment_mode", "")).replace(
        "_", " "
    ).title()
    app.markdown(
        f"""
    <div class="ui-panel">
        <div class="ui-panel-head">
            <span class="ui-panel-title">Segment summary</span>
            <span class="ui-panel-meta">Outcome: {seg.outcome_question_id} · {segment_mode}</span>
        </div>
    </div>
    """,
        unsafe_allow_html=True,
    )

    col1, col2, col3, col4 = app.columns(4)
    col1.metric("Winners", f"{seg.winner_n:,}")
    col2.metric(f"{laggard_label}s", f"{seg.loser_n:,}")
    col3.metric("Differentiators", len(seg.differentiators))
    col4.metric("Avg lift (top 5)", f"{avg_lift:.1f}x")

    app.markdown(
        """
    <div class="ui-panel">
        <div class="ui-panel-head">
            <span class="ui-panel-title">Top differentiators</span>
            <span class="ui-panel-meta">Ranked by association strength</span>
        </div>
    </div>
    """,
        unsafe_allow_html=True,
    )

    table_rows = [
        {
            "#": rank,
            "Question": _truncate(d.question_text, 50),
            "Top option": d.top_option_label,
            f"{winner_label}": f"{d.top_option_winner_rate:.1%}",
            f"{laggard_label}": f"{d.top_option_loser_rate:.1%}",
            "Lift": "∞" if d.top_option_lift >= 900 else f"{d.top_option_lift:.2f}x",
            "Cramér's V": f"{d.cramers_v:.3f}",
        }
        for rank, d in enumerate(seg.differentiators[:15], start=1)
    ]
    app.dataframe(table_rows, use_container_width=True, hide_index=True)


def _render_smart_cross_cut_suggestions_panel(
    seg: OutcomeSegmentationResult,
) -> None:
    app = _require_streamlit()
    app.markdown("### Smart Cross-cut Suggestions")
    app.caption("Ranked by relevance to your outcome variable")

    if not app.session_state.get("cross_cut_results"):
        app.info(
            "Run cross cuts in Section 4 first to see outcome-ranked "
            "suggestions here."
        )
        return

    suggestions = app.session_state.get("cross_cut_suggestions", [])
    if not suggestions:
        app.info("No rule-based cross-cut suggestions are available.")
        return

    scored_suggestions = score_suggestions_for_outcome(suggestions, seg)
    schema = app.session_state.get("schema")
    for suggestion in scored_suggestions[:10]:
        analysis_type = suggestion.analysis_type
        analysis_type_label = getattr(analysis_type, "value", str(analysis_type))
        with app.expander(
            f"[{suggestion.outcome_relevance_score:.2f}] "
            f"{_suggestion_label(suggestion, schema)}"
        ):
            app.markdown(
                f"**Business question:** {suggestion.business_question}"
            )
            app.markdown(
                "**Source questions:** "
                f"{', '.join(suggestion.source_question_ids)}"
            )
            app.markdown(
                "**Outcome relevance:** "
                f"{suggestion.outcome_relevance_score:.2f}"
            )
            app.markdown(f"**Analysis type:** {analysis_type_label}")


def _render_winner_profile_panel(seg: OutcomeSegmentationResult) -> None:
    app = _require_streamlit()
    app.markdown(
        f'<h3 style="display:flex;align-items:center;margin:1rem 0;">'
        f'{_icon("winner", "#CC0000")}Winner Profile</h3>',
        unsafe_allow_html=True,
    )

    profile = seg.winner_profile
    if not profile.defining_traits:
        app.warning(
            "Not enough strong differentiators to build a winner profile. "
            "Try selecting a different outcome variable or segment definition."
        )
        return

    app.markdown(
        f"**{profile.winner_label} Profile** "
        f"(n={profile.winner_n} vs {profile.loser_label} n={profile.loser_n})"
    )

    for trait in profile.defining_traits:
        app.markdown(f"#### {trait.question_id}")
        app.caption(trait.question_text)

        trait_payload = {
            "question_text": trait.question_text,
            "option_label": trait.option_label,
            "winner_rate": trait.winner_rate,
            "loser_rate": trait.loser_rate,
            "lift": trait.lift,
            "rate_gap": trait.rate_gap,
            "winner_label": profile.winner_label,
            "laggard_label": profile.loser_label,
            "laggard_top_option_label": trait.laggard_top_option_label,
            "laggard_top_option_winner_rate": trait.laggard_top_option_winner_rate,
            "laggard_top_option_loser_rate": trait.laggard_top_option_loser_rate,
        }

        try:
            trait_insight = generate_insight(
                table_payload=trait_payload,
                table_kind="winner_profile_trait",
                title_hint=trait.question_id,
                cache=_INSIGHT_CACHE,
            )
            _render_insight_card(trait_insight, label="")
        except Exception:
            pass

        app.markdown(
            f'<div style="font-size:11px; color:#888; text-transform:uppercase; '
            f'letter-spacing:0.05em; margin-top:8px; margin-bottom:4px;">'
            f'{html.escape(profile.winner_label)}s chose: '
            f'<strong style="color:#1A1A1A;">'
            f'{html.escape(trait.option_label)}</strong></div>',
            unsafe_allow_html=True,
        )
        col_a1, col_a2, col_a3 = app.columns(3)
        col_a1.metric(profile.winner_label, f"{trait.winner_rate:.1%}")
        col_a2.metric(profile.loser_label, f"{trait.loser_rate:.1%}")
        col_a3.metric("Gap", f"{trait.rate_gap:+.1%}")

        if trait.laggard_top_option_label:
            laggard_gap = (
                trait.laggard_top_option_loser_rate
                - trait.laggard_top_option_winner_rate
            )
            app.markdown(
                f'<div style="font-size:11px; color:#888; text-transform:uppercase; '
                f'letter-spacing:0.05em; margin-top:8px; margin-bottom:4px;">'
                f'{html.escape(profile.loser_label)}s chose: '
                f'<strong style="color:#1A1A1A;">'
                f'{html.escape(trait.laggard_top_option_label)}</strong></div>',
                unsafe_allow_html=True,
            )
            col_b1, col_b2, col_b3 = app.columns(3)
            col_b1.metric(
                profile.winner_label,
                f"{trait.laggard_top_option_winner_rate:.1%}",
            )
            col_b2.metric(
                profile.loser_label,
                f"{trait.laggard_top_option_loser_rate:.1%}",
            )
            col_b3.metric("Gap", f"{laggard_gap:+.1%}")

        app.divider()


# ---------------------------------------------------------------------------
# Main
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


def _render_nav_bar() -> None:
    """Inject scrollspy nav tabs into the merged red header (#custom-header-nav)."""
    app = _require_streamlit()

    n_singlecuts = len(app.session_state.get("results", []))
    n_crosscuts = len(app.session_state.get("cross_cut_results", []))
    seg = app.session_state.get("segmentation_result")
    n_diffs = len(seg.differentiators) if seg else 0
    gf_state = app.session_state.get("global_filter_state")
    n_filters = len(gf_state.filters) if gf_state is not None and hasattr(gf_state, "filters") else 0
    has_data = bool(app.session_state.get("run_complete")) or app.session_state.get("schema") is not None

    tabs_meta = [
        ("section-upload",     "Upload",      "\u2713" if has_data else None, "check"),
        ("section-filter",     "Filter",      str(n_filters) if n_filters else None, None),
        ("section-singlecuts", "Single cuts", str(n_singlecuts) if n_singlecuts else None, None),
        ("section-crosscuts",  "Cross cuts",  str(n_crosscuts) if n_crosscuts else None, None),
        ("section-ai",         "AI analysis", str(n_diffs) if n_diffs else None, None),
        ("section-downloads",  "Downloads",   None, None),
    ]

    nav_items = ""
    for anchor_id, label, badge, badge_type in tabs_meta:
        if badge:
            if badge_type == "check":
                badge_html = (
                    '<span class=\\"nav-badge\\" '
                    'style=\\"background:rgba(76,175,80,0.9);color:#FFF;\\">\u2713</span>'
                )
            else:
                badge_html = f'<span class=\\"nav-badge\\">{html.escape(badge)}</span>'
        else:
            badge_html = ""
        nav_items += (
            f'<a href=\\"#{html.escape(anchor_id)}\\" class=\\"nav-tab\\" '
            f'data-target=\\"{html.escape(anchor_id)}\\">'
            f'{html.escape(label)}{badge_html}</a>'
        )

    try:
        from streamlit.components.v1 import html as components_html
    except Exception:
        components_html = None

    script = f"""
<script>
(function() {{
    var navHTML = "{nav_items}";
    var doc = window.parent.document;

    function injectNavTabs() {{
        var navContainer = doc.getElementById('custom-header-nav');
        if (!navContainer) return false;
        if (navContainer.dataset.signature !== navHTML) {{
            navContainer.innerHTML = navHTML;
            navContainer.dataset.signature = navHTML;
            attachScrollspy();
        }}
        return true;
    }}

    function attachScrollspy() {{
        var tabs = Array.prototype.slice.call(
            doc.querySelectorAll('#custom-header-nav .nav-tab')
        );
        if (!tabs.length) return;
        var parentWin = window.parent;

        tabs.forEach(function(tab) {{
            tab.addEventListener('click', function(e) {{
                e.preventDefault();
                var targetId = tab.getAttribute('data-target');
                var target = doc.getElementById(targetId);
                if (target) target.scrollIntoView({{ behavior: 'smooth', block: 'start' }});
                setActive(targetId);
            }});
        }});

        function setActive(activeId) {{
            tabs.forEach(function(tab) {{
                if (tab.getAttribute('data-target') === activeId) {{
                    tab.classList.add('active');
                }} else {{
                    tab.classList.remove('active');
                }}
            }});
        }}

        function updateActiveTab() {{
            var scrollTop = parentWin.scrollY || parentWin.document.documentElement.scrollTop || 0;
            var activeId = tabs[0].getAttribute('data-target');
            var bestOffset = -Infinity;
            tabs.forEach(function(tab) {{
                var section = doc.getElementById(tab.getAttribute('data-target'));
                if (!section) return;
                var sectionTop = section.getBoundingClientRect().top + scrollTop - 80;
                if (sectionTop <= scrollTop && sectionTop > bestOffset) {{
                    bestOffset = sectionTop;
                    activeId = tab.getAttribute('data-target');
                }}
            }});
            setActive(activeId);
        }}

        // Tear down prior listeners/interval to prevent buildup on re-runs
        if (parentWin.__siNavCleanup) {{
            try {{ parentWin.__siNavCleanup(); }} catch (e) {{}}
        }}
        parentWin.addEventListener('scroll', updateActiveTab, {{ passive: true }});
        parentWin.document.addEventListener('scroll', updateActiveTab, {{ passive: true }});
        var intervalId = parentWin.setInterval(updateActiveTab, 200);
        parentWin.__siNavCleanup = function() {{
            parentWin.removeEventListener('scroll', updateActiveTab);
            parentWin.document.removeEventListener('scroll', updateActiveTab);
            parentWin.clearInterval(intervalId);
        }};
        updateActiveTab();
    }}

    injectNavTabs();
    setTimeout(injectNavTabs, 100);
    setTimeout(injectNavTabs, 500);
    setTimeout(injectNavTabs, 1500);
    setInterval(injectNavTabs, 1000);
}})();
</script>
"""
    if components_html is not None:
        components_html(script, height=0)
    else:
        app.markdown(script, unsafe_allow_html=True)
    return

def main() -> None:
    app = _require_streamlit()
    app.set_page_config(
        page_title=APP_TITLE,
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    _initialise_session_state()
    _inject_global_css()
    _drain_pending_actions()
    _render_sidebar()

    _render_nav_bar()

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
    _section_ai_analysis()
    app.divider()
    _section_downloads()


# ---------------------------------------------------------------------------
# Setup wizard
# ---------------------------------------------------------------------------


def _render_setup_wizard(uploaded_files: list[Any]) -> None:
    app = _require_streamlit()
    _render_wizard_css()
    if app.session_state.get("wiz_category_assignments") is None:
        _wizard_prepare(uploaded_files)
    if app.session_state.get("wiz_category_assignments") is None:
        return

    app.markdown("<div class='wiz-shell'>", unsafe_allow_html=True)
    _render_wizard_stepper()
    _render_wizard_nav("top")
    step = int(app.session_state.get("wiz_step", 1))
    if step == 1:
        _render_wizard_step_categories()
    elif step == 2:
        _render_wizard_step_demographics()
    elif step == 3:
        _render_wizard_step_custom_filters()
    elif step == 4:
        _render_wizard_step_per_question_filters()
    else:
        _render_wizard_step_crosscut()
    _render_wizard_nav("bottom")
    app.markdown("</div>", unsafe_allow_html=True)


def _wizard_prepare(uploaded_files: list[Any]) -> None:
    app = _require_streamlit()
    try:
        from src.ai_insights import categorize_questions_into_themes
        from src.io import load_survey_inputs
        from src.question_classifier import classify_questions
        from src.ui.wizard import (
            category_assignments_from_themes,
            selected_demographics_from_schema,
        )

        data_map, decoded_df, load_report = load_survey_inputs(uploaded_files)
        schema = classify_questions(
            data_map,
            decoded_df.columns.tolist(),
            respondent_id_column="record",
            total_respondents=len(decoded_df),
            source_rawdata_path=load_report.raw_data_source,
        )
        questions_for_themes = [
            {
                "question_id": question.canonical_id,
                "question_text": question.question_text,
                "question_type": question.question_type.value,
                "is_demographic": getattr(question, "is_demographic", False),
            }
            for question in schema.questions
        ]
        try:
            auto_themes = categorize_questions_into_themes(
                questions_for_themes,
                cache=_INSIGHT_CACHE,
            )
        except Exception:
            auto_themes = {
                "themes": [
                    {
                        "name": "All Questions",
                        "question_ids": [q.canonical_id for q in schema.questions if q.analysis_eligible],
                    }
                ],
                "was_template": True,
                "error_message": "",
            }
        app.session_state["wiz_schema"] = schema
        app.session_state["wiz_decoded_df"] = decoded_df
        app.session_state["wiz_load_report"] = load_report
        app.session_state["wiz_category_assignments"] = category_assignments_from_themes(schema, auto_themes)
        app.session_state["wiz_selected_demographics"] = selected_demographics_from_schema(schema)
    except Exception as exc:  # noqa: BLE001
        app.session_state["wiz_category_assignments"] = None
        app.error(f"Setup wizard could not read the uploaded files: {type(exc).__name__}: {exc}")


def _render_wizard_css() -> None:
    _require_streamlit().markdown(
        """
<style>
.wiz-shell { background: #FFFFFF; padding: 24px 0; font-family: Arial, sans-serif; }
.wiz-stepper { display: flex; justify-content: space-between; margin: 24px 0; }
.wiz-step-circle { width: 32px; height: 32px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-weight: 500; font-size: 14px; margin: 0 auto 8px; }
.wiz-step-active { border: 2px solid #CC0000; color: #CC0000; }
.wiz-step-complete { background: #CC0000; color: white; }
.wiz-step-upcoming { border: 2px solid #999; color: #999; }
.wiz-step-label { text-align: center; font-size: 12px; font-family: Arial, sans-serif; }
.wiz-category-cell { background: #CC0000; color: white; padding: 12px 16px; font-weight: 700; font-size: 14px; font-family: Arial, sans-serif; border-radius: 4px; }
.wiz-card { border: 1px solid #E5E5E5; border-radius: 6px; padding: 12px; margin: 4px 0; min-height: 92px; }
.wiz-card-id { color: #666; font-size: 11px; font-family: Arial, sans-serif; }
.wiz-card-text { font-size: 13px; font-family: Arial, sans-serif; margin: 4px 0; }
.wiz-card-type { color: #666; font-size: 11px; font-style: italic; font-family: Arial, sans-serif; }
.wiz-divider { border-top: 1px solid #E5E5E5; margin: 16px 0; }
</style>
""",
        unsafe_allow_html=True,
    )


def _render_wizard_stepper() -> None:
    app = _require_streamlit()
    labels = ["Categories", "Demographics", "Custom filters", "Per-Q filters", "Cross-cuts"]
    current = int(app.session_state.get("wiz_step", 1))
    cols = app.columns(5)
    for index, label in enumerate(labels, start=1):
        if index < current:
            css_class = "wiz-step-complete"
            marker = "?"
        elif index == current:
            css_class = "wiz-step-active"
            marker = str(index)
        else:
            css_class = "wiz-step-upcoming"
            marker = str(index)
        weight = "700" if index == current else "400"
        cols[index - 1].markdown(
            f"<div class='wiz-step-circle {css_class}'>{marker}</div>"
            f"<div class='wiz-step-label' style='font-weight:{weight}'>{html.escape(label)}</div>",
            unsafe_allow_html=True,
        )


def _render_wizard_nav(location: str) -> None:
    app = _require_streamlit()
    step = int(app.session_state.get("wiz_step", 1))
    back_col, step_col, next_col = app.columns([1, 2, 1])
    if back_col.button("Back", disabled=(step <= 1), key=f"wiz_back_{location}"):
        app.session_state["wiz_step"] = max(1, step - 1)
        _wizard_rerun(app)
    step_col.markdown(f"<div style='text-align:center;color:#666;font-size:13px;'>Step {step} of 5</div>", unsafe_allow_html=True)
    if step < 5:
        if next_col.button("Next", type="primary", key=f"wiz_next_{location}"):
            app.session_state["wiz_step"] = min(5, step + 1)
            _wizard_rerun(app)
    elif next_col.button("Run Analysis ?", type="primary", key=f"wiz_run_{location}"):
        _wizard_apply_overrides()
        app.session_state["wiz_complete"] = True
        app.session_state["_wizard_run_requested"] = True
        _wizard_rerun(app)


def _wizard_schema():
    return _require_streamlit().session_state.get("wiz_schema")


def _wizard_dataframe():
    return _require_streamlit().session_state.get("wiz_decoded_df")


def _wizard_questions_by_category() -> dict[str, list[Any]]:
    schema = _wizard_schema()
    assignments = dict(_require_streamlit().session_state.get("wiz_category_assignments") or {})
    grouped: dict[str, list[Any]] = {}
    if schema is None:
        return grouped
    questions = {question.canonical_id: question for question in schema.questions}
    for question_id, category in assignments.items():
        question = questions.get(question_id)
        if question is not None:
            grouped.setdefault(category, []).append(question)
    for category in _require_streamlit().session_state.get("wiz_empty_categories", []):
        grouped.setdefault(category, [])
    return grouped


def _render_wizard_step_categories() -> None:
    app = _require_streamlit()
    from src.ui.wizard import question_display_text, question_type_label

    app.markdown("### Review & edit categories")
    pending = app.session_state.get("wiz_pending_remove")
    if pending:
        question = (_wizard_schema() or None).get_question(pending) if _wizard_schema() else None
        label = question.question_text if question else pending
        app.warning(f"Remove {label} from the analysis? It will not appear in the output workbook.")
        yes, no, _ = app.columns([1, 1, 4])
        if yes.button("Confirm remove", key="wiz_confirm_remove", type="primary"):
            assignments = dict(app.session_state.get("wiz_category_assignments") or {})
            assignments.pop(pending, None)
            app.session_state["wiz_category_assignments"] = assignments
            app.session_state["wiz_pending_remove"] = None
            _wizard_rerun(app)
        if no.button("Cancel", key="wiz_cancel_remove"):
            app.session_state["wiz_pending_remove"] = None
            _wizard_rerun(app)

    grouped = _wizard_questions_by_category()
    categories = list(grouped)
    for category_index, category in enumerate(categories):
        left, right = app.columns([1, 4])
        with left:
            app.markdown(f"<div class='wiz-category-cell'>{html.escape(category)}</div>", unsafe_allow_html=True)
            renamed = app.text_input("Rename", value=category, key=f"wiz_rename_{category_index}_{category}", label_visibility="collapsed")
            renamed = renamed.strip()
            if renamed and renamed != category:
                assignments = dict(app.session_state.get("wiz_category_assignments") or {})
                app.session_state["wiz_category_assignments"] = {
                    question_id: (renamed if assigned == category else assigned)
                    for question_id, assigned in assignments.items()
                }
                app.session_state["wiz_empty_categories"] = [
                    renamed if item == category else item
                    for item in app.session_state.get("wiz_empty_categories", [])
                ]
                _wizard_rerun(app)
        with right:
            questions = grouped.get(category, [])
            if not questions:
                app.caption("No questions in this category yet.")
            for row_start in range(0, len(questions), 3):
                cols = app.columns(3)
                for card_col, question in zip(cols, questions[row_start : row_start + 3]):
                    with card_col:
                        app.markdown(
                            "<div class='wiz-card'>"
                            f"<div class='wiz-card-id'>{html.escape(question.canonical_id)}</div>"
                            f"<div class='wiz-card-text' title='{html.escape(question.question_text)}'>{html.escape(question_display_text(question, 40))}</div>"
                            f"<div class='wiz-card-type'>{html.escape(question_type_label(question))}</div>"
                            "</div>",
                            unsafe_allow_html=True,
                        )
                        move_options = ["Stay"] + [item for item in categories if item != category]
                        move_to = app.selectbox("Move to", move_options, key=f"wiz_move_{question.canonical_id}", label_visibility="collapsed")
                        if move_to != "Stay":
                            assignments = dict(app.session_state.get("wiz_category_assignments") or {})
                            assignments[question.canonical_id] = move_to
                            app.session_state["wiz_category_assignments"] = assignments
                            _wizard_rerun(app)
                        if app.button("Remove", key=f"wiz_remove_{question.canonical_id}"):
                            app.session_state["wiz_pending_remove"] = question.canonical_id
                            _wizard_rerun(app)
        app.markdown("<div class='wiz-divider'></div>", unsafe_allow_html=True)

    if app.button("+ Add category", key="wiz_show_add_category"):
        app.session_state["wiz_show_add_category_input"] = True
    if app.session_state.get("wiz_show_add_category_input"):
        new_category = app.text_input("Category name", key="wiz_new_category_name")
        if app.button("Create", key="wiz_create_category", type="primary"):
            name = new_category.strip()
            if name:
                app.session_state.setdefault("wiz_empty_categories", [])
                if name not in app.session_state["wiz_empty_categories"]:
                    app.session_state["wiz_empty_categories"].append(name)
            app.session_state["wiz_show_add_category_input"] = False
            _wizard_rerun(app)


def _render_wizard_step_demographics() -> None:
    app = _require_streamlit()
    from src.ui.wizard import distinct_value_preview, eligible_filter_question_ids, question_display_text

    schema = _wizard_schema()
    dataframe = _wizard_dataframe()
    if schema is None:
        app.info("Upload files to review demographic filters.")
        return
    app.markdown("### Confirm demographic filters")
    selected = set(app.session_state.get("wiz_selected_demographics") or [])
    candidate_ids = [q.canonical_id for q in schema.questions if q.is_demographic or q.canonical_id in selected]
    for question_id in candidate_ids:
        question = schema.get_question(question_id)
        if question is None:
            continue
        checked = app.checkbox(
            f"{question.canonical_id} - {question_display_text(question, 72)}",
            value=question_id in selected,
            key=f"wiz_demo_{question_id}",
        )
        if checked:
            selected.add(question_id)
        else:
            selected.discard(question_id)
        app.caption(distinct_value_preview(dataframe, question))
    app.session_state["wiz_selected_demographics"] = sorted(selected)

    available = [qid for qid in eligible_filter_question_ids(schema) if qid not in selected]
    labels = {"": "Select a question"}
    labels.update({qid: f"{qid} - {question_display_text(schema.get_question(qid), 72)}" for qid in available if schema.get_question(qid)})
    choice = app.selectbox("+ Add another question as a filter", [""] + available, format_func=lambda value: labels.get(value, value), key="wiz_add_demo_choice")
    if app.button("Add filter", key="wiz_add_demo_button", disabled=(choice == "")):
        selected.add(choice)
        app.session_state["wiz_selected_demographics"] = sorted(selected)
        _wizard_rerun(app)


def _render_wizard_step_custom_filters() -> None:
    app = _require_streamlit()
    from src.ui.wizard import normalise_custom_filter_count

    app.markdown("### Custom workbook filters")
    app.write("Custom filters let partners set arbitrary filters on the entire workbook in addition to demographic ones. How many would you like?")
    value = app.number_input("Custom workbook filters", min_value=0, max_value=5, value=normalise_custom_filter_count(app.session_state.get("wiz_num_custom_filters")), step=1, key="wiz_num_custom_filters_input")
    app.session_state["wiz_num_custom_filters"] = int(value)
    app.caption("Setting saved ? exporter wiring coming in the next iteration.")


def _render_wizard_step_per_question_filters() -> None:
    app = _require_streamlit()
    from src.ui.wizard import normalise_per_question_filter_count

    app.markdown("### Per-question filters")
    app.write("How many filter rows should appear under each question block?")
    value = app.number_input("Per-question filters", min_value=0, max_value=3, value=normalise_per_question_filter_count(app.session_state.get("wiz_num_per_question_filters")), step=1, key="wiz_num_per_question_filters_input")
    app.session_state["wiz_num_per_question_filters"] = int(value)
    app.caption("Setting saved ? exporter wiring coming in the next iteration.")


def _render_wizard_step_crosscut() -> None:
    app = _require_streamlit()
    app.markdown("### Cross-cut output format")
    app.write("When you generate cross-cuts from the dashboard, where should they live?")
    options = {"One consolidated sheet": "one_sheet", "Separate sheet per cross-cut": "separate_sheets"}
    current = app.session_state.get("wiz_crosscut_consolidation", "one_sheet")
    labels = list(options)
    index = list(options.values()).index(current) if current in options.values() else 0
    selected = app.radio("Cross-cut output format", labels, index=index, key="wiz_crosscut_choice")
    app.session_state["wiz_crosscut_consolidation"] = options[selected]
    app.caption("Setting saved ? exporter wiring coming in the next iteration.")


def _wizard_apply_overrides() -> None:
    app = _require_streamlit()
    schema = app.session_state.get("wiz_schema")
    if schema is None:
        return
    from src.ui.wizard import apply_wizard_schema_overrides, themes_from_wizard_assignments

    assignments = dict(app.session_state.get("wiz_category_assignments") or {})
    selected_demographics = list(app.session_state.get("wiz_selected_demographics") or [])
    schema_override = apply_wizard_schema_overrides(schema, assignments, selected_demographics)
    app.session_state["wiz_schema_override"] = schema_override
    app.session_state["wiz_theme_override"] = themes_from_wizard_assignments(schema_override, assignments)


def _wizard_pipeline_overrides(schema: Any, themes: dict) -> tuple[Any, dict]:
    app = _require_streamlit()
    return (
        app.session_state.get("wiz_schema_override") or schema,
        app.session_state.get("wiz_theme_override") or themes,
    )


def _wizard_rerun(app: Any) -> None:
    rerun = getattr(app, "rerun", None) or getattr(app, "experimental_rerun", None)
    if callable(rerun):
        rerun()


if __name__ == "__main__":
    main()
