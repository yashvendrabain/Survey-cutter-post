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
    schema = _require_streamlit().session_state["schema"]
    return {
        spec.canonical_id: f"{spec.canonical_id}: {spec.question_text}"
        for spec in schema.questions
    }


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
    if display_mode == "Counts":
        app.dataframe(df, use_container_width=True)
    else:
        app.dataframe(
            df.style.format("{:.1%}"),
            use_container_width=True,
        )
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
        app.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
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
        app.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
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
        app.dataframe(pd.DataFrame(grid_rows), use_container_width=True, hide_index=True)
    else:
        app.info("Preview not available for this target type.")


def _preview_group_comparison(result: Any) -> None:
    import pandas as pd
    app = _require_streamlit()
    rt = result.result_table
    seg_q = rt.get("segment_question_id", "")
    met_q = rt.get("metric_question_id", "")
    app.caption(f"Metric: {met_q}   Segments: {seg_q}")
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
    app.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    app.caption("Group means in the downloaded workbook.")


def _preview_expected_vs_realized(result: Any) -> None:
    import pandas as pd
    app = _require_streamlit()
    rt = result.result_table
    exp_q = rt.get("expected_question_id", "")
    real_q = rt.get("realized_question_id", "")
    app.caption(f"Expected: {exp_q}   Realized: {real_q}")
    df = pd.DataFrame(
        [
            {"Metric": "Paired N", "Count": rt.get("paired_n", 0)},
            {"Metric": "Expected valid N", "Count": (rt.get("expected", {}) or {}).get("valid_n", 0)},
            {"Metric": "Realized valid N", "Count": (rt.get("realized", {}) or {}).get("valid_n", 0)},
        ]
    )
    app.dataframe(df, use_container_width=True, hide_index=True)
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
    """Counts-only display of a SingleCutResult."""
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
        app.dataframe(
            pd.DataFrame(grid_rows), use_container_width=True, hide_index=True
        )
    elif isinstance(result, SingleSelectResult):
        rows = [
            {
                "Code": code,
                "Label": payload.get("label", ""),
                "Count": payload.get("count", 0),
            }
            for code, payload in sorted(
                result.distribution.items(), key=lambda kv: str(kv[0])
            )
        ]
        app.dataframe(
            pd.DataFrame(rows), use_container_width=True, hide_index=True
        )
    elif isinstance(result, MultiSelectResult):
        rows = []
        for sub_id, payload in result.selections.items():
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
        app.dataframe(
            pd.DataFrame(rows), use_container_width=True, hide_index=True
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


def _render_single_cut_card(result: Any, spec: Any) -> None:
    from src.filtered_single_cut import compute_filtered_single_cut
    from src.models import FilterSpec

    app = _require_streamlit()
    schema = app.session_state["schema"]
    short_text = (spec.question_text or "")[:80]

    expander_label = f"{spec.canonical_id} \u2014 {short_text}"
    with app.expander(expander_label, expanded=False):
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
                try:
                    filtered_result = compute_filtered_single_cut(
                        spec.canonical_id,
                        specs,
                        schema,
                        app.session_state["active_df"],
                        app.session_state["log"],
                    )
                    app.session_state.setdefault("filtered_results", {})
                    app.session_state["filtered_results"][
                        spec.canonical_id
                    ] = filtered_result
                    app.session_state["filtered_workbook_bytes"] = None
                    app.rerun()
                except Exception as exc:  # noqa: BLE001
                    app.error(f"Filter failed: {type(exc).__name__}: {exc}")
                    import hashlib
                    exc_digest = hashlib.sha1(
                        str(exc).encode("utf-8")
                    ).hexdigest()[:10]
                    if app.checkbox(
                        "Show technical details",
                        key=f"show_tb_{spec.canonical_id}_{exc_digest}",
                    ):
                        app.code(traceback.format_exc())

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

    for index, (spec, reason) in enumerate(suggestions[:15], start=1):
        col_text, col_button = app.columns([4, 1])
        col_text.write(f"{index}. {spec.title}")
        col_text.caption(reason)
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
    app = _require_streamlit()
    sb = app.sidebar
    sb.title(APP_TITLE)
    sb.caption(APP_TAGLINE)

    sb.markdown("### Sections")
    sb.markdown(
        "\n".join(
            [
                f"- [{SECTION_UPLOAD}](#section-1)",
                f"- [{SECTION_GLOBAL_FILTER}](#section-2)",
                f"- [{SECTION_RESULTS}](#section-3)",
                f"- [{SECTION_CROSS_CUTS}](#section-4)",
                f"- [{SECTION_DOWNLOADS}](#section-5)",
            ]
        )
    )

    sb.markdown("### Status")
    has_load_report = app.session_state.get("load_report") is not None
    files_selected = app.session_state.get("raw_data_path_label") is not None
    if has_load_report:
        sb.write("\u2705 Inputs loaded")
    elif files_selected:
        sb.write("\u23f3 Files selected \u2014 click Run analysis")
    else:
        sb.write("\u23f3 Waiting for upload")

    load_report = app.session_state.get("load_report")
    if load_report is not None:
        sb.caption(f"Input: {load_report.scenario.replace('_', ' ')}")
        sb.caption(f"Data source: {load_report.raw_data_source}")
        sb.caption(f"Map source: {load_report.datamap_source}")

    if app.session_state["run_complete"]:
        results_count = len(app.session_state["results"])
        sb.write(f"\u2705 Analysis complete \u2014 {results_count} results")
    else:
        sb.write("\u23f3 No analysis run yet")

    gf_state = app.session_state.get("global_filter_state")
    if gf_state is not None and gf_state.is_active():
        stats = app.session_state.get("global_filter_stats") or {}
        rows_after = stats.get("rows_after", 0)
        rows_before = stats.get("rows_before", 0)
        sb.write(
            f"{STATUS_GLOBAL_FILTER_ACTIVE} \u2014 {rows_after:,} of "
            f"{rows_before:,} respondents"
        )
    else:
        sb.write(STATUS_GLOBAL_FILTER_INACTIVE)

    with sb.expander("Need help?"):
        sb.markdown(
            "- **Upload** a CSV or XLSX of raw responses and the data map XLSX.\n"
            "- **Global filter** restricts every analysis to a subset of "
            "respondents.\n"
            "- **Per-question filters** layer on top of the global filter for "
            "ad-hoc slices.\n"
            "- **Cross cuts** combine two questions; tick the suggestion box "
            "for ideas.\n"
            "- **Download** the workbook(s) at the bottom."
        )


# ---------------------------------------------------------------------------
# Section 1 — Upload
# ---------------------------------------------------------------------------


def _section_upload() -> None:
    app = _require_streamlit()
    app.header(SECTION_UPLOAD, anchor="section-1")
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
    from src.global_filter import apply_global_filter
    from src.models import GlobalFilterState

    app = _require_streamlit()
    schema = app.session_state.get("schema")
    decoded_df = app.session_state.get("decoded_df")
    spec_list = []
    for q, v in rows:
        vals = _normalize_value_list(v)
        if q is None or not vals:
            continue
        spec_list.append(_build_filter_spec(schema, decoded_df, q, vals))
    specs = tuple(spec_list)

    try:
        state = GlobalFilterState(filters=specs)
    except ValueError as exc:
        app.error(f"Invalid global filter: {exc}")
        return

    try:
        filtered_df, stats = apply_global_filter(
            app.session_state["decoded_df"], state
        )
    except ValueError as exc:
        app.error(f"Could not apply global filter: {exc}")
        return

    app.session_state["global_filter_state"] = state
    app.session_state["global_filter_stats"] = stats
    app.session_state["active_df"] = filtered_df
    _rerun_single_cuts_on_active_df()
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
    app.header(SECTION_GLOBAL_FILTER, anchor="section-2")

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
    app = _require_streamlit()
    app.header(SECTION_RESULTS, anchor="section-3")

    if not app.session_state["run_complete"]:
        app.info(EMPTY_NO_RESULTS)
        return

    results = app.session_state["results"]
    schema = app.session_state["schema"]

    app.text_input(
        "Search questions",
        placeholder="Filter by ID or text",
        key="ss_search",
        help="Case-insensitive substring match on canonical ID or question text.",
    )
    needle = (app.session_state.get("ss_search") or "").strip().lower()

    matched = 0
    for result in results:
        spec = schema.get_question(result.question_id)
        if spec is None:
            continue
        if needle:
            haystack = f"{spec.canonical_id} {spec.question_text or ''}".lower()
            if needle not in haystack:
                continue
        matched += 1
        _render_single_cut_card(result, spec)

    if needle and matched == 0:
        app.caption(f"No questions match '{needle}'.")
    else:
        app.caption(f"Showing {matched} of {len(results)} single cuts.")


# ---------------------------------------------------------------------------
# Section 4 — Cross cuts
# ---------------------------------------------------------------------------


def _section_cross_cuts() -> None:
    app = _require_streamlit()
    app.header(SECTION_CROSS_CUTS, anchor="section-4")

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

    app.markdown("**Cross-cut results**")
    for result in results:
        with app.expander(
            f"{result.cross_cut_id} \u2014 {result.synthetic_question_title}",
            expanded=False,
        ):
            app.caption(
                f"Type: {result.analysis_type.value}  \u00b7  "
                f"Display: {result.display_mode}  \u00b7  "
                f"{len(result.audit_records)} audit records  \u00b7  "
                f"{', '.join(result.source_question_ids)}"
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
    app.header(SECTION_DOWNLOADS, anchor="section-5")

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
    _render_sidebar()

    app.title(APP_TITLE)
    app.caption(APP_TAGLINE)
    app.divider()

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
