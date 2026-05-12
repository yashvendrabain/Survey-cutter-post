"""Streamlit entry point for the Survey Insight Engine."""

from __future__ import annotations

import os
import sys
import tempfile
import traceback
import html
from pathlib import Path
from typing import Any

try:
    import streamlit as st
except ModuleNotFoundError:  # pragma: no cover - local import smoke test fallback.
    st = None  # type: ignore[assignment]


PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

try:
    from config import APP_NAME, VERSION
except (ModuleNotFoundError, ImportError):  # pragma: no cover - target app provides config.
    APP_NAME = "Survey Insight Engine"
    VERSION = "Stage 2"

from src.ai_insights import generate_insight
from src.cross_cut_suggestions import score_suggestions_for_outcome
from src.models import InsightResult, OutcomeSegmentationResult


_INSIGHT_CACHE: dict[str, Any] = {}


SESSION_DEFAULTS = {
    "data_map": None,
    "dataframe": None,
    "results": [],
    "skips": [],
    "schema": None,
    "quality_report": None,
    "log": None,
    "output_path": None,
    "cross_cut_results": [],
    "cross_cut_skips": [],
    "cross_cut_suggestions": [],
    "cross_cut_only_bytes": None,
    "survey_type_result": None,
    "outcome_variable_id": None,
    "segmentation_result": None,
    "run_complete": False,
}


def _require_streamlit() -> Any:
    if st is None:
        raise RuntimeError(
            "Streamlit is not installed in this environment. "
            "Install project requirements before running the app."
        )
    return st


def _initialise_session_state() -> None:
    app = _require_streamlit()
    for key, value in SESSION_DEFAULTS.items():
        app.session_state.setdefault(key, value)


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


def _run_pipeline(raw_data_path: str, datamap_path: str, status: Any) -> None:
    from src.calculation_log import CalculationLog
    from src.cross_cut_suggestions import suggest_cross_cuts
    from src.datamap_parser import parse_datamap
    from src.excel_exporter import export_single_cuts
    from src.question_classifier import classify_questions
    from src.raw_decoder import decode_raw_data
    from src.single_cut import compute_single_cuts

    _INSIGHT_CACHE.clear()

    status.update(label="Parsing data map...", state="running")
    data_map = parse_datamap(datamap_path)

    status.update(label="Decoding raw data...", state="running")
    dataframe, quality_report = decode_raw_data(raw_data_path, data_map)

    status.update(label="Classifying questions...", state="running")
    schema = classify_questions(
        data_map,
        dataframe.columns.tolist(),
        respondent_id_column="record",
        total_respondents=len(dataframe),
        source_rawdata_path=raw_data_path,
    )

    status.update(label="Computing single cuts...", state="running")
    log = CalculationLog()
    results, skips = compute_single_cuts(schema, dataframe, log)

    status.update(label="Exporting workbook...", state="running")
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
    app.session_state["data_map"] = data_map
    app.session_state["dataframe"] = dataframe
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
    app.session_state["survey_type_result"] = None
    app.session_state["outcome_variable_id"] = None
    app.session_state["segmentation_result"] = None
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
        app.session_state["dataframe"],
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


def _render_header() -> None:
    app = _require_streamlit()
    app.set_page_config(page_title=APP_NAME, layout="wide")
    app.title(f"{APP_NAME} {VERSION}")
    app.caption(
        "Upload raw survey data and a data map to produce audited "
        "single-cut and cross-cut analysis in Excel."
    )


def _section_anchor(anchor_id: str) -> None:
    app = _require_streamlit()
    app.markdown(
        f'<div id="{html.escape(anchor_id)}" class="section-anchor"></div>',
        unsafe_allow_html=True,
    )


def _active_filter_count() -> int:
    app = _require_streamlit()
    count = 0
    global_filter_state = app.session_state.get("global_filter_state")
    if global_filter_state is not None and hasattr(global_filter_state, "filters"):
        count += len(global_filter_state.filters)

    for key, value in app.session_state.items():
        if not str(key).startswith("filters_"):
            continue
        if isinstance(value, (dict, list, tuple, set)):
            count += len(value)
            continue
        if value:
            count += 1
    return count


def _render_nav_bar() -> None:
    app = _require_streamlit()
    has_data = bool(app.session_state.get("run_complete"))
    n_filters = _active_filter_count()
    n_singlecuts = len(app.session_state.get("results", []))
    n_crosscuts = len(app.session_state.get("cross_cut_results", []))
    seg: OutcomeSegmentationResult | None = app.session_state.get(
        "segmentation_result"
    )
    n_diffs = len(seg.differentiators) if seg is not None else 0

    tabs = [
        ("section-upload", "Upload", "✓" if has_data else None, "check"),
        ("section-filter", "Filter", str(n_filters) if n_filters else None, None),
        (
            "section-singlecuts",
            "Single cuts",
            str(n_singlecuts) if n_singlecuts else None,
            None,
        ),
        (
            "section-crosscuts",
            "Cross cuts",
            str(n_crosscuts) if n_crosscuts else None,
            None,
        ),
        ("section-ai", "AI analysis", str(n_diffs) if n_diffs else None, None),
        ("section-downloads", "Downloads", None, None),
    ]

    nav_items = ""
    for anchor_id, label, badge, badge_type in tabs:
        if badge:
            if badge_type == "check":
                badge_html = (
                    '<span class="nav-badge" '
                    'style="background:#EAF3DE;color:#27500A;">✓</span>'
                )
            else:
                badge_html = f'<span class="nav-badge">{html.escape(badge)}</span>'
        else:
            badge_html = ""

        nav_items += f"""
        <a href="#{html.escape(anchor_id)}"
           class="nav-tab"
           data-target="{html.escape(anchor_id)}">
            {html.escape(label)}{badge_html}
        </a>"""

    app.markdown(
        """
        <style>
        .nav-bar {
            position: fixed !important;
            top: 50px !important;
            left: 0 !important;
            right: 0 !important;
            z-index: 99998 !important;
            background: #FFFFFF !important;
            border-bottom: 1px solid #E5E5E5 !important;
            display: flex !important;
            gap: 0 !important;
            padding: 0 24px !important;
            height: 44px !important;
            align-items: center !important;
            overflow-x: auto !important;
            box-shadow: 0 1px 4px rgba(0,0,0,0.06) !important;
        }
        .main .block-container {
            padding-top: 110px !important;
        }
        .section-anchor {
            scroll-margin-top: 120px;
        }
        .nav-tab {
            display: inline-flex !important;
            align-items: center !important;
            gap: 6px !important;
            padding: 0 16px !important;
            height: 44px !important;
            font-size: 13px !important;
            font-family: Arial, sans-serif !important;
            font-weight: 400 !important;
            color: #666 !important;
            text-decoration: none !important;
            border-bottom: 2px solid transparent !important;
            border-radius: 0 !important;
            background: transparent !important;
            white-space: nowrap !important;
            cursor: pointer !important;
            transition: color 0.15s ease, border-color 0.15s ease !important;
            box-sizing: border-box !important;
        }
        .nav-tab:hover {
            color: #1A1A1A !important;
            background: transparent !important;
            border-bottom-color: #DDD !important;
        }
        .nav-tab.active {
            color: #CC0000 !important;
            font-weight: 500 !important;
            border-bottom-color: #CC0000 !important;
            background: transparent !important;
        }
        .nav-badge {
            display: inline-flex !important;
            align-items: center !important;
            justify-content: center !important;
            background: #F0F0F0 !important;
            color: #666 !important;
            font-size: 10px !important;
            font-weight: 500 !important;
            padding: 1px 6px !important;
            border-radius: 10px !important;
            min-width: 18px !important;
            line-height: 1.4 !important;
        }
        .nav-tab.active .nav-badge {
            background: #FCEBEB !important;
            color: #CC0000 !important;
        }
        .nav-tab-sep {
            width: 1px !important;
            height: 16px !important;
            background: #E5E5E5 !important;
            flex-shrink: 0 !important;
        }
        </style>
        <div class="nav-bar">
        """
        + nav_items
        + """
        </div>
        <script>
        (function() {
            function positionNavBar() {
                var header = document.querySelector('.brand-header');
                var navBar = document.querySelector('.nav-bar');
                if (header && navBar) {
                    var headerHeight = header.offsetHeight;
                    navBar.style.top = headerHeight + 'px';
                    var mainBlock = document.querySelector('.main .block-container');
                    if (mainBlock) {
                        mainBlock.style.paddingTop = (headerHeight + 44 + 16) + 'px';
                    }
                }
            }

            function updateActiveTab() {
                var tabs = Array.prototype.slice.call(
                    document.querySelectorAll('.nav-tab')
                );
                if (!tabs.length) {
                    return;
                }
                var activeId = tabs[0].getAttribute('data-target');
                tabs.forEach(function(tab) {
                    var section = document.getElementById(
                        tab.getAttribute('data-target')
                    );
                    if (section && section.getBoundingClientRect().top <= 120) {
                        activeId = tab.getAttribute('data-target');
                    }
                });
                tabs.forEach(function(tab) {
                    if (tab.getAttribute('data-target') === activeId) {
                        tab.classList.add('active');
                    } else {
                        tab.classList.remove('active');
                    }
                });
            }

            positionNavBar();
            updateActiveTab();
            window.addEventListener('resize', positionNavBar);
            window.addEventListener('scroll', updateActiveTab, { passive: true });
            setTimeout(positionNavBar, 300);
            setTimeout(positionNavBar, 800);
            setTimeout(updateActiveTab, 300);
            setTimeout(updateActiveTab, 800);
        })();
        </script>
        """,
        unsafe_allow_html=True,
    )


def _render_upload_section() -> tuple[Any | None, Any | None]:
    app = _require_streamlit()
    _section_anchor("section-upload")
    app.subheader("Upload files")
    raw_col, datamap_col = app.columns(2)

    with raw_col:
        raw_upload = app.file_uploader(
            "Raw survey data",
            type=["csv", "xlsx"],
            key="raw_data_upload",
        )
        app.caption(_upload_status(raw_upload))

    with datamap_col:
        datamap_upload = app.file_uploader(
            "Data map",
            type=["xlsx"],
            key="datamap_upload",
        )
        app.caption(_upload_status(datamap_upload))

    return raw_upload, datamap_upload


def _handle_run(raw_upload: Any | None, datamap_upload: Any | None) -> None:
    app = _require_streamlit()
    app.subheader("Run analysis")
    ready_to_run = raw_upload is not None and datamap_upload is not None

    if not app.button(
        "Run analysis",
        type="primary",
        disabled=not ready_to_run,
    ):
        return

    raw_data_path = None
    datamap_path = None
    app.session_state["run_complete"] = False

    try:
        raw_data_path = _write_upload_to_temp(raw_upload)
        datamap_path = _write_upload_to_temp(datamap_upload)
        with app.status("Starting analysis...", expanded=True) as status:
            _run_pipeline(raw_data_path, datamap_path, status)
    except Exception as exc:  # noqa: BLE001 - UI must show any pipeline failure.
        app.session_state["run_complete"] = False
        app.error(f"{type(exc).__name__}: {exc}")
        with app.expander("Show full traceback"):
            app.code(traceback.format_exc())
    finally:
        _cleanup_temp_files(raw_data_path, datamap_path)


def _skip_rows(skips: list[Any]) -> list[dict[str, Any]]:
    return [
        {
            "Canonical ID": skip.canonical_id,
            "Type": skip.question_type.value,
            "Reason": skip.skip_reason,
            "Details": skip.details or "",
        }
        for skip in skips
    ]


def _result_rows(results: list[Any]) -> list[dict[str, Any]]:
    return [
        {
            "Canonical ID": result.question_id,
            "Type": result.question_type.value,
            "Valid N": result.valid_n,
            "Missing N": result.missing_n,
        }
        for result in results
    ]


def _numeric_allocation_rows(result: Any, schema: Any) -> list[dict[str, Any]]:
    spec = schema.get_question(result.question_id) if schema is not None else None
    rows: list[dict[str, Any]] = []
    for option_id, payload in (result.per_option_stats or {}).items():
        label = spec.option_map.get(option_id, option_id) if spec is not None else option_id
        rows.append(
            {
                "Option": label,
                "Mean": round(float(payload["mean"]), 1),
                "Median": round(float(payload["median"]), 1),
                "Valid N": int(payload.get("valid_n", 0)),
                "Missing N": int(payload.get("missing_n", 0)),
            }
        )
    return rows


def _cross_cut_rows(results: list[Any], schema: Any) -> list[dict[str, Any]]:
    return [
        {
            "Cross Cut ID": result.cross_cut_id,
            "Title": _cross_cut_display_title(result, schema, max_chars=80),
            "Type": result.analysis_type.value,
            "Source Questions": _source_question_labels(
                result.source_question_ids,
                schema,
                max_chars=80,
            ),
            "Warnings": " | ".join(result.warnings),
        }
        for result in results
    ]


def _suggestion_label(suggestion: Any, schema: Any) -> str:
    if schema is None:
        return suggestion.synthetic_question_title
    return _source_question_labels(suggestion.source_question_ids, schema)


def _cross_cut_display_title(
    result: Any,
    schema: Any,
    *,
    max_chars: int | None = None,
) -> str:
    source_question_ids = getattr(result, "source_question_ids", ())
    if schema is None or not source_question_ids:
        return getattr(result, "synthetic_question_title", getattr(result, "cross_cut_id", str(result)))
    return _source_question_labels(
        source_question_ids,
        schema,
        include_ids=False,
        max_chars=max_chars,
    )


def _source_question_labels(
    question_ids: tuple[str, ...],
    schema: Any,
    *,
    include_ids: bool = True,
    max_chars: int | None = None,
) -> str:
    parts: list[str] = []
    for question_id in question_ids:
        spec = schema.get_question(question_id) if schema is not None else None
        if spec is None:
            parts.append(question_id)
            continue
        text = spec.question_text
        if max_chars is not None and len(text) > max_chars:
            text = text[: max_chars - 3] + "..."
        parts.append(f"{question_id}: {text}" if include_ids else text)
    return " × ".join(parts)


def _preview_cross_tab(result: Any) -> None:
    app = _require_streamlit()
    import pandas as pd

    ct = result.result_table
    counts = ct.get("counts")
    if not isinstance(counts, dict):
        app.info("Preview not available: counts table missing.")
        return

    row_label_map = ct.get("row_label_map", {})
    col_label_map = ct.get("column_label_map", {})
    source_ids = tuple(result.source_question_ids)
    row_question = source_ids[0] if len(source_ids) > 0 else ct.get("row_question_id", "")
    column_question = (
        source_ids[1] if len(source_ids) > 1 else ct.get("column_question_id", "")
    )
    row_codes = sorted(counts.keys(), key=str)
    col_codes = sorted(
        {
            column_code
            for row_payload in counts.values()
            if isinstance(row_payload, dict)
            for column_code in row_payload.keys()
        },
        key=str,
    )
    if not row_codes or not col_codes:
        app.info("Preview not available: no cross-tab counts to show.")
        return

    dataframe = pd.DataFrame(
        index=[row_label_map.get(row_code, str(row_code)) for row_code in row_codes],
        columns=[
            col_label_map.get(column_code, str(column_code))
            for column_code in col_codes
        ],
        data=[
            [
                counts.get(row_code, {}).get(column_code, 0)
                if isinstance(counts.get(row_code), dict)
                else 0
                for column_code in col_codes
            ]
            for row_code in row_codes
        ],
    )
    dataframe.index.name = f"↓ {row_question}"
    dataframe.columns.name = f"→ {column_question}"
    app.caption(f"Rows: {row_question}   Columns: {column_question}")
    app.dataframe(dataframe, use_container_width=True)
    app.caption(f"Grand total: {ct.get('grand_total', 0):,} responses")


def _preview_segment_profile(result: Any) -> None:
    app = _require_streamlit()
    import pandas as pd

    result_table = result.result_table
    app.caption(
        f"Filter: {result_table.get('filter_expr', '<no filter>')}  ·  "
        f"Filter N: {result_table.get('filter_n', 0):,}"
    )

    target_result = result_table.get("target_result", {})
    if not isinstance(target_result, dict):
        app.info("Preview not available for this target type.")
        return

    if "distribution" in target_result:
        rows = [
            {
                "Code": code,
                "Label": payload.get("label", ""),
                "Count": payload.get("count", 0),
            }
            for code, payload in sorted(
                target_result.get("distribution", {}).items(),
                key=lambda item: str(item[0]),
            )
            if isinstance(payload, dict)
        ]
        app.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
        return

    if "selections" in target_result:
        rows = [
            {
                "Sub-column": sub_column_id,
                "Label": payload.get("label", ""),
                "Count": payload.get("count", 0),
            }
            for sub_column_id, payload in target_result.get("selections", {}).items()
            if isinstance(payload, dict)
        ]
        rows.sort(key=lambda row: row["Count"], reverse=True)
        app.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
        return

    if "mean" in target_result:
        dataframe = pd.DataFrame(
            [
                {"Statistic": "Valid N", "Count": target_result.get("valid_n", 0)},
                {"Statistic": "Missing N", "Count": target_result.get("missing_n", 0)},
            ]
        )
        app.dataframe(dataframe, use_container_width=True, hide_index=True)
        app.caption("Numeric statistics (mean, median, std) in the downloaded workbook.")
        return

    if "rows" in target_result:
        grid_rows = []
        for sub_column_id, row_result in target_result.get("rows", {}).items():
            if not isinstance(row_result, dict):
                continue
            row_payload = {"Row": sub_column_id}
            for code, payload in row_result.get("distribution", {}).items():
                if isinstance(payload, dict):
                    row_payload[f"{code}: {payload.get('label', '')}"] = payload.get(
                        "count", 0
                    )
            grid_rows.append(row_payload)
        app.caption(
            f"Grid with {len(target_result.get('rows', {}))} rows. Per-row counts:"
        )
        app.dataframe(pd.DataFrame(grid_rows), use_container_width=True, hide_index=True)
        return

    app.info("Preview not available for this target type.")


def _preview_group_comparison(result: Any) -> None:
    app = _require_streamlit()
    import pandas as pd

    result_table = result.result_table
    segment_question = result_table.get("segment_question_id", "")
    metric_question = result_table.get("metric_question_id", "")
    app.caption(f"Metric: {metric_question}   Segments: {segment_question}")

    rows = [
        {
            "Segment": segment_data.get("label", str(segment_value)),
            "N": segment_data.get("n", 0),
        }
        for segment_value, segment_data in result_table.get("per_segment", {}).items()
        if isinstance(segment_data, dict)
    ]
    overall = result_table.get("overall", {})
    rows.append(
        {
            "Segment": "Overall",
            "N": overall.get("valid_n", overall.get("n", 0))
            if isinstance(overall, dict)
            else 0,
        }
    )
    app.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    app.caption("Group means in the downloaded workbook.")


def _preview_expected_vs_realized(result: Any) -> None:
    app = _require_streamlit()
    import pandas as pd

    result_table = result.result_table
    expected_question = result_table.get("expected_question_id", "")
    realized_question = result_table.get("realized_question_id", "")
    expected = result_table.get("expected", {})
    realized = result_table.get("realized", {})
    app.caption(f"Expected: {expected_question}   Realized: {realized_question}")
    dataframe = pd.DataFrame(
        [
            {"Metric": "Paired N", "Count": result_table.get("paired_n", 0)},
            {
                "Metric": "Expected valid N",
                "Count": expected.get("valid_n", 0) if isinstance(expected, dict) else 0,
            },
            {
                "Metric": "Realized valid N",
                "Count": realized.get("valid_n", 0) if isinstance(realized, dict) else 0,
            },
        ]
    )
    app.dataframe(dataframe, use_container_width=True, hide_index=True)
    app.caption("Mean expected, mean realized, gap statistics in the downloaded workbook.")


def _render_cross_cut_preview(result: Any) -> None:
    app = _require_streamlit()
    try:
        from src.models import AnalysisType

        analysis_type = result.analysis_type
        if analysis_type == AnalysisType.CROSS_TAB:
            _preview_cross_tab(result)
        elif analysis_type == AnalysisType.SEGMENT_PROFILE:
            _preview_segment_profile(result)
        elif analysis_type == AnalysisType.GROUP_COMPARISON:
            _preview_group_comparison(result)
        elif analysis_type == AnalysisType.EXPECTED_VS_REALIZED:
            _preview_expected_vs_realized(result)
        else:
            app.info(f"Preview not implemented for {analysis_type.value}.")
    except Exception as exc:  # noqa: BLE001 - preview should never break the app.
        app.error(f"Could not render preview: {type(exc).__name__}: {exc}")


def _audit_rows(records: tuple[Any, ...]) -> list[dict[str, Any]]:
    return [
        {
            "Metric": record.metric_name,
            "Question": record.source_question_id,
            "Source Columns": ", ".join(record.source_columns),
            "Formula": record.formula,
            "Value": record.value_raw,
            "Valid N": record.valid_n,
        }
        for record in records[:100]
    ]


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
        if col_button.button("Run", key=f"run_suggestion_{spec.cross_cut_id}"):
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
        submitted = app.form_submit_button("Run manual cross cut")

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
    except Exception as exc:  # noqa: BLE001 - UI must show invalid manual specs.
        app.error(f"{type(exc).__name__}: {exc}")


def _render_cross_cut_results() -> None:
    app = _require_streamlit()
    results = app.session_state["cross_cut_results"]
    schema = app.session_state.get("schema")

    _section_anchor("section-crosscuts")
    app.subheader("Cross cuts")
    if app.checkbox("Show suggested cross cuts", key="show_suggested_cross_cuts"):
        with app.expander("Suggested cross cuts", expanded=True):
            _render_suggested_cross_cuts()

    with app.expander("Run manual cross cut"):
        _render_manual_cross_cut()

    if not results:
        app.write("No cross cuts have been run yet.")
        return

    app.write("Cross-cut results")
    for result in results:
        with app.container():
            col_check, col_title = app.columns([1, 11])
            with col_check:
                app.checkbox(
                    "Include in cross-cut workbook",
                    value=True,
                    key=f"cc_select_{result.cross_cut_id}",
                    label_visibility="collapsed",
            )
            with col_title:
                app.markdown(f"**{_cross_cut_display_title(result, schema)}**")
                app.caption(
                    f"Type: {result.analysis_type.value}  ·  "
                    f"Display: {result.display_mode}  ·  "
                    f"{len(result.audit_records)} audit records"
                )

            with app.expander("Preview (counts only)", expanded=True):
                _render_cross_cut_preview(result)

            if result.warnings:
                with app.expander("Warnings"):
                    for warning in result.warnings:
                        app.write(f"• {warning}")

            app.divider()

    selected_results = [
        result
        for result in app.session_state["cross_cut_results"]
        if app.session_state.get(f"cc_select_{result.cross_cut_id}", True)
    ]

    if app.button(
        "Download selected cross cuts",
        disabled=(len(selected_results) == 0),
        help=(
            f"{len(selected_results)} cross cuts selected"
            if selected_results
            else "Tick at least one cross cut to enable download"
        ),
    ):
        from src.excel_exporter import export_cross_cuts_only

        cc_output_path = "/tmp/cross_cuts.xlsx"
        export_cross_cuts_only(
            cross_cut_results=selected_results,
            schema=app.session_state["schema"],
            log=app.session_state["log"],
            output_path=cc_output_path,
        )
        with open(cc_output_path, "rb") as file:
            app.session_state["cross_cut_only_bytes"] = file.read()

    if app.session_state["cross_cut_only_bytes"]:
        app.download_button(
            label="Download cross-cut workbook",
            data=app.session_state["cross_cut_only_bytes"],
            file_name="cross_cut_analysis.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )

    with app.expander("View cross-cut results table"):
        app.dataframe(_cross_cut_rows(results, schema), use_container_width=True)


def _render_insight_section(
    insight_key: str,
    table_payload: dict[str, Any],
    table_kind: str,
    title_hint: str = "",
) -> None:
    app = _require_streamlit()
    stored = app.session_state.get(insight_key)

    if app.button("Generate insight", key=f"generate_{insight_key}"):
        with app.spinner("Generating insight..."):
            stored = generate_insight(
                _normalise_insight_payload(table_payload, table_kind, title_hint),
                table_kind=table_kind,
                title_hint=title_hint,
                cache=_INSIGHT_CACHE,
            )
            app.session_state[insight_key] = stored

    if not stored:
        return

    if stored.title:
        app.markdown(f"**{stored.title}**")
    app.write(stored.insight)
    if stored.was_template:
        app.caption("Template insight")
        if stored.error_message:
            app.info(f"AI unavailable: {stored.error_message}")
    else:
        app.caption(f"AI insight generated with {stored.model_used}")


def _inject_insight_css() -> None:
    app = _require_streamlit()

    if app.session_state.get("_insight_css_injected"):
        return

    app.markdown(
        """
        <style>
        .insight-card {
            background: linear-gradient(135deg, #CC0000 0%, #8B0000 100%);
            border-radius: 8px;
            padding: 20px 24px;
            margin: 12px 0 8px 0;
            box-shadow: 0 4px 12px rgba(204, 0, 0, 0.25);
            position: relative;
            overflow: hidden;
        }
        .insight-card::before {
            content: '"';
            position: absolute;
            top: -10px;
            left: 16px;
            font-size: 80px;
            color: rgba(255,255,255,0.15);
            font-family: Georgia, serif;
            line-height: 1;
        }
        .insight-headline {
            color: #FFFFFF;
            font-size: 15px;
            font-weight: 600;
            line-height: 1.5;
            margin: 0;
            padding-left: 8px;
            letter-spacing: 0.01em;
        }
        .insight-footer {
            color: rgba(255,255,255,0.6);
            font-size: 11px;
            margin-top: 10px;
            padding-left: 8px;
        }
        .insight-template {
            background: linear-gradient(135deg, #666666 0%, #444444 100%);
            border-radius: 8px;
            padding: 12px 16px;
            margin: 8px 0;
        }
        .insight-template p {
            color: rgba(255,255,255,0.7);
            font-size: 13px;
            margin: 0;
            font-style: italic;
        }
        .insight-label {
            color: rgba(255,255,255,0.85);
            font-size: 11px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            margin-bottom: 6px;
            padding-left: 8px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    app.session_state["_insight_css_injected"] = True


def _render_insight_card(insight: InsightResult) -> None:
    _render_insight_card_v2(insight)


def _render_insight_card_v2(insight: InsightResult, label: str = "") -> None:
    app = _require_streamlit()
    if not insight or not insight.insight:
        return

    _inject_insight_css()

    headline = html.escape(insight.insight)
    label_html = (
        f'<div class="insight-label">{html.escape(label)}</div>'
        if label
        else ""
    )

    if insight.was_template:
        app.markdown(
            f"""
        <div class="insight-template">
            {label_html}
            <p>💡 {headline}</p>
        </div>
        """,
            unsafe_allow_html=True,
        )
        return

    footer = (
        f"AI insight · {html.escape(insight.model_used)}"
        if insight.model_used
        else "AI insight"
    )
    app.markdown(
        f"""
    <div class="insight-card">
        {label_html}
        <p class="insight-headline">{headline}</p>
        <p class="insight-footer">{footer}</p>
    </div>
    """,
        unsafe_allow_html=True,
    )


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


def _section_outcome_segmentation() -> None:
    app = _require_streamlit()
    if not app.session_state.get("run_complete"):
        return

    data_map = app.session_state.get("data_map")
    dataframe = app.session_state.get("dataframe")
    schema = app.session_state.get("schema")
    if data_map is None or dataframe is None or schema is None:
        return

    from src.survey_type_detector import detect_survey_type

    survey_type_result = app.session_state.get("survey_type_result")
    if survey_type_result is None:
        survey_type_result = detect_survey_type(data_map, dataframe)
        app.session_state["survey_type_result"] = survey_type_result

    app.markdown("---")
    app.markdown("## 🎯 Outcome Variable Selection")

    outcome_options = _outcome_variable_options(survey_type_result, schema)
    if not outcome_options:
        app.info("No eligible outcome variables were found for segmentation.")
        return

    current_outcome = app.session_state.get("outcome_variable_id")
    if current_outcome not in outcome_options:
        current_outcome = (
            survey_type_result.outcome_question_id
            if survey_type_result.outcome_question_id in outcome_options
            else outcome_options[0]
        )
        app.session_state["outcome_variable_id"] = current_outcome

    selected_outcome = app.selectbox(
        "Outcome variable",
        outcome_options,
        index=outcome_options.index(current_outcome),
        format_func=lambda question_id: _question_label(schema, question_id),
    )
    app.session_state["outcome_variable_id"] = selected_outcome

    with app.expander("Top N Outcome Candidates"):
        candidate_rows = _outcome_candidate_rows(survey_type_result)
        if candidate_rows:
            app.dataframe(candidate_rows, use_container_width=True, hide_index=True)
        else:
            app.write("No ranked outcome candidates available.")

    segment_definition = _render_segment_definition_ui(
        dataframe,
        schema,
        selected_outcome,
    )
    _render_segmentation_results_ui(dataframe, schema, selected_outcome, segment_definition)


def _outcome_variable_options(survey_type_result: Any, schema: Any) -> list[str]:
    options: list[str] = []
    for candidate in survey_type_result.all_eligible_questions:
        if schema.get_question(candidate.question_id) is not None:
            options.append(candidate.question_id)

    if options:
        return options

    return [
        spec.canonical_id
        for spec in schema.questions
        if spec.analysis_eligible
        and spec.question_type
        in {
            _question_type_enum("SINGLE_SELECT"),
            _question_type_enum("DIRECT_NUMERIC"),
        }
    ]


def _question_type_enum(name: str) -> Any:
    from src.models import QuestionType

    return getattr(QuestionType, name)


def _question_label(schema: Any, question_id: str) -> str:
    question = schema.get_question(question_id)
    if question is None:
        return question_id
    return f"{question.canonical_id}: {question.question_text}"


def _outcome_candidate_rows(survey_type_result: Any) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for rank, candidate in enumerate(
        survey_type_result.all_eligible_questions[:10],
        start=1,
    ):
        rows.append(
            {
                "Rank": rank,
                "Question": candidate.question_id,
                "Score": f"{candidate.relevance_score:.2f}",
                "Reason": candidate.reason,
                "Text": candidate.question_text,
            }
        )
    return rows


def _render_segment_definition_ui(
    dataframe: Any,
    schema: Any,
    outcome_question_id: str,
) -> Any:
    from src.models import SegmentDefinition

    app = _require_streamlit()
    outcome_spec = schema.get_question(outcome_question_id)
    app.markdown("## ⚙️ Segment Definition")
    if outcome_spec is None:
        app.error(f"Outcome question {outcome_question_id!r} is not in the schema.")
        return None

    mode_label = app.radio(
        "Segmentation mode",
        ["Categorical", "Numeric threshold", "Quartile"],
        horizontal=True,
        key=f"segment_mode_{outcome_question_id}",
    )

    if mode_label == "Categorical":
        values = _categorical_values(dataframe, outcome_spec)
        if not values:
            app.warning("This outcome variable has no usable category values.")
            return None
        selected_values = app.multiselect(
            "Winner values",
            values,
            default=(values[:1]),
            format_func=lambda value: _value_label(outcome_spec, value),
            key=f"winner_values_{outcome_question_id}",
        )
        if not selected_values:
            app.warning("Choose at least one winner value before running segmentation.")
            return None
        return SegmentDefinition(
            outcome_question_id=outcome_question_id,
            segment_mode="categorical",
            winner_values=tuple(selected_values),
        )

    if mode_label == "Numeric threshold":
        numeric = _numeric_outcome(dataframe, outcome_spec)
        if numeric.empty:
            app.warning("This outcome variable has no numeric values.")
            return None
        default_threshold = float(numeric.quantile(0.75))
        threshold = app.number_input(
            "Winner threshold",
            value=default_threshold,
            key=f"winner_threshold_{outcome_question_id}",
        )
        threshold_direction = app.radio(
            "Winner direction",
            ["gte", "lte"],
            format_func=lambda value: (
                "At or above threshold" if value == "gte" else "At or below threshold"
            ),
            horizontal=True,
            key=f"threshold_direction_{outcome_question_id}",
        )
        return SegmentDefinition(
            outcome_question_id=outcome_question_id,
            segment_mode="numeric_threshold",
            winner_threshold=float(threshold),
            threshold_direction=threshold_direction,
        )

    quartile_winner = app.radio(
        "Winner quartile",
        ["top", "bottom"],
        format_func=lambda value: (
            "Top quartile" if value == "top" else "Bottom quartile"
        ),
        horizontal=True,
        key=f"quartile_winner_{outcome_question_id}",
    )
    return SegmentDefinition(
        outcome_question_id=outcome_question_id,
        segment_mode="quartile",
        quartile_winner=quartile_winner,
    )


def _render_segmentation_results_ui(
    dataframe: Any,
    schema: Any,
    outcome_question_id: str,
    segment_definition: Any,
) -> None:
    app = _require_streamlit()
    app.markdown("## 📈 Segmentation Results")
    if segment_definition is None:
        app.info("Complete the segment definition to run segmentation.")
        return

    if app.button("Run segmentation", type="primary"):
        from src.outcome_segmentation import compute_outcome_segmentation

        audit_records = []
        try:
            result = compute_outcome_segmentation(
                dataframe,
                schema,
                outcome_question_id,
                segment_definition,
                audit_records,
            )
            for record in audit_records:
                app.session_state["log"].record(record)
            app.session_state["segmentation_result"] = result
            if audit_records:
                _refresh_full_workbook()
            app.success("Segmentation complete.")
        except Exception as exc:  # noqa: BLE001 - UI must show invalid segmentation.
            app.error(f"{type(exc).__name__}: {exc}")
            return

    seg: OutcomeSegmentationResult | None = app.session_state.get(
        "segmentation_result"
    )
    if seg is None or seg.outcome_question_id != outcome_question_id:
        app.info("Run segmentation to see differentiators and the winner profile.")
        return

    laggard_label = seg.segment_definition.loser_label
    col1, col2, col3, col4 = app.columns(4)
    col1.metric("Outcome Variable", seg.outcome_question_id)
    col2.metric(seg.segment_definition.winner_label + "s", seg.winner_n)
    col3.metric(laggard_label + "s", seg.loser_n)
    col4.metric("Differentiators", len(seg.differentiators))

    if seg.warnings:
        with app.expander("Segmentation warnings", expanded=True):
            for warning in seg.warnings:
                app.write(warning)

    if not seg.differentiators:
        app.write("No differentiators met the minimum sample and variation checks.")
        return

    rows = [
        {
            "Question": diff.question_id,
            "Top option": diff.top_option_label,
            "Winner Rate": f"{diff.top_option_winner_rate:.1%}",
            f"{laggard_label} Rate": f"{diff.top_option_loser_rate:.1%}",
            "Lift": "∞" if diff.top_option_lift >= 900 else f"{diff.top_option_lift:.2f}x",
            "Cramér's V": f"{diff.cramers_v:.3f}",
        }
        for diff in seg.differentiators[:10]
    ]
    app.dataframe(rows, use_container_width=True, hide_index=True)


def _primary_column(question: Any) -> str:
    return question.raw_columns[0] if question.raw_columns else question.canonical_id


def _categorical_values(dataframe: Any, question: Any) -> list[Any]:
    if question.option_map:
        return list(question.option_map.keys())

    column = _primary_column(question)
    if column not in dataframe.columns:
        return []
    values = [value for value in dataframe[column].dropna().unique()]
    return sorted(values, key=str)[:50]


def _value_label(question: Any, value: Any) -> str:
    label = question.option_map.get(value)
    if label is None:
        return str(value)
    return f"{value}: {label}"


def _numeric_outcome(dataframe: Any, question: Any) -> Any:
    import pandas as pd

    column = _primary_column(question)
    if column not in dataframe.columns:
        return pd.Series(dtype=float)
    return pd.to_numeric(dataframe[column], errors="coerce").dropna()


def _section_ai_analysis() -> None:
    app = _require_streamlit()
    _section_anchor("section-ai")
    if not app.session_state.get("run_complete"):
        return

    _section_outcome_segmentation()
    app.divider()

    app.markdown("---")
    app.markdown("## 5️⃣ AI Analysis")

    seg: OutcomeSegmentationResult | None = app.session_state.get(
        "segmentation_result"
    )
    if seg is None:
        app.info(
            "Complete outcome segmentation in the Outcome Variable Selection "
            "section above to unlock AI Analysis."
        )
        return

    _render_outcome_summary_panel(seg)
    _render_smart_cross_cut_suggestions_panel(seg)
    _render_winner_profile_panel(seg)


def _render_outcome_summary_panel(seg: OutcomeSegmentationResult) -> None:
    app = _require_streamlit()
    app.markdown("### 🎯 Outcome Analysis")

    laggard_label = seg.segment_definition.loser_label
    col1, col2, col3, col4 = app.columns(4)
    col1.metric("Outcome Variable", seg.outcome_question_id)
    col2.metric("Winners", seg.winner_n)
    col3.metric(laggard_label + "s", seg.loser_n)
    col4.metric("Differentiators", len(seg.differentiators))

    if not seg.differentiators:
        app.info("No differentiators found. Try different segment definition.")
        return

    table_payload = {
        "outcome_question_id": seg.outcome_question_id,
        "winner_label": seg.segment_definition.winner_label,
        "loser_label": seg.segment_definition.loser_label,
        "winner_n": seg.winner_n,
        "loser_n": seg.loser_n,
        "differentiators": [
            {
                "question_text": diff.question_text,
                "top_option_label": diff.top_option_label,
                "winner_rate": diff.top_option_winner_rate,
                "loser_rate": diff.top_option_loser_rate,
                "lift": diff.top_option_lift,
                "cramers_v": diff.cramers_v,
            }
            for diff in seg.differentiators
        ],
    }

    from src.ai_insights import generate_outlier_insight, generate_table_insight

    with app.spinner("Generating key insight..."):
        table_insight = generate_table_insight(
            table_payload=table_payload,
            table_kind="differentiator_table",
            cache=_INSIGHT_CACHE,
        )
    _render_insight_card_v2(table_insight, label="Key Insight")

    with app.spinner("Identifying outlier..."):
        outlier_insight = generate_outlier_insight(
            table_payload=table_payload,
            table_kind="outlier",
            cache=_INSIGHT_CACHE,
        )
    _render_insight_card_v2(outlier_insight, label="Outlier")

    app.markdown("#### Top Differentiators")
    table_rows = [
        {
            "#": rank,
            "Question": _truncate(diff.question_text, 80),
            "Top Option": diff.top_option_label,
            "Cramér's V": f"{diff.cramers_v:.3f}",
            f"{seg.segment_definition.winner_label} Rate": f"{diff.top_option_winner_rate:.1%}",
            f"{seg.segment_definition.loser_label} Rate": f"{diff.top_option_loser_rate:.1%}",
            "Lift": "∞" if diff.top_option_lift >= 900 else f"{diff.top_option_lift:.2f}x",
        }
        for rank, diff in enumerate(seg.differentiators[:15], start=1)
    ]
    app.dataframe(table_rows, use_container_width=True, hide_index=True)


def _render_smart_cross_cut_suggestions_panel(
    seg: OutcomeSegmentationResult,
) -> None:
    app = _require_streamlit()
    app.markdown("### 🔗 Smart Cross-cut Suggestions")
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

    schema = app.session_state.get("schema")
    scored_suggestions = score_suggestions_for_outcome(suggestions, seg)
    for suggestion in scored_suggestions[:10]:
        analysis_type = suggestion.analysis_type
        analysis_type_label = getattr(analysis_type, "value", str(analysis_type))
        label = _suggestion_label(suggestion, schema)
        with app.expander(
            f"[{suggestion.outcome_relevance_score:.2f}] "
            f"{label}"
        ):
            app.markdown(f"**Business question:** {suggestion.business_question}")
            app.markdown(
                "**Source questions:** "
                f"{_source_question_labels(suggestion.source_question_ids, schema)}"
            )
            app.markdown(
                "**Outcome relevance:** "
                f"{suggestion.outcome_relevance_score:.2f}"
            )
            app.markdown(f"**Analysis type:** {analysis_type_label}")


def _render_winner_profile_panel(seg: OutcomeSegmentationResult) -> None:
    app = _require_streamlit()
    app.markdown("### 🏆 Winner Profile")

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
        with app.container():
            col1, col2, col3 = app.columns([4, 1, 1])
            col1.markdown(f"**{trait.question_id}:** {trait.option_label}")
            col1.caption(trait.question_text)
            col2.metric(profile.winner_label, f"{trait.winner_rate:.1%}")
            col3.metric(
                "vs " + profile.loser_label,
                f"{trait.loser_rate:.1%}",
                delta=f"{trait.rate_gap:+.1%}",
            )
            app.divider()

    _render_insight_section(
        insight_key="insight_winner_profile",
        table_payload={
            "winner_label": profile.winner_label,
            "loser_label": profile.loser_label,
            "winner_n": profile.winner_n,
            "loser_n": profile.loser_n,
            "traits": [
                {
                    "question_id": trait.question_id,
                    "option_label": trait.option_label,
                    "winner_rate": trait.winner_rate,
                    "loser_rate": trait.loser_rate,
                    "lift": trait.lift,
                    "rate_gap": trait.rate_gap,
                }
                for trait in profile.defining_traits
            ],
        },
        table_kind="winner_profile",
        title_hint=f"{profile.winner_label} archetype summary",
    )


def _truncate(text: str, max_chars: int) -> str:
    if len(text) <= max_chars:
        return text
    return text[: max_chars - 3] + "..."


def _render_results_section() -> None:
    app = _require_streamlit()
    if not app.session_state["run_complete"]:
        return

    _section_anchor("section-singlecuts")
    results = app.session_state["results"]
    skips = app.session_state["skips"]
    schema = app.session_state["schema"]
    quality_report = app.session_state["quality_report"]
    log = app.session_state["log"]
    output_path = app.session_state["output_path"]

    app.subheader("Results")
    total_col, results_col, skips_col, audit_col = app.columns(4)
    total_col.metric("Total questions analysed", len(schema.questions))
    results_col.metric("Results produced", len(results))
    skips_col.metric("Skipped questions", len(skips))
    audit_col.metric("Audit log records", len(log))

    _section_anchor("section-downloads")
    if output_path and os.path.exists(output_path):
        with open(output_path, "rb") as workbook_file:
            workbook_bytes = workbook_file.read()
        app.download_button(
            label="Download Excel workbook",
            data=workbook_bytes,
            file_name="survey_analysis.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )

    _render_cross_cut_results()

    with app.expander("View quality warnings"):
        if quality_report.warnings:
            for warning in quality_report.warnings:
                app.write(warning)
        else:
            app.write("No warnings")

    with app.expander("View skipped questions"):
        app.dataframe(_skip_rows(skips), use_container_width=True)

    with app.expander("View results preview"):
        app.dataframe(_result_rows(results), use_container_width=True)
        allocation_results = [
            result
            for result in results
            if getattr(result.question_type, "value", "") == "NUMERIC_ALLOCATION"
        ]
        for result in allocation_results:
            app.markdown(f"**{result.question_id} allocation options**")
            app.dataframe(
                _numeric_allocation_rows(result, schema),
                use_container_width=True,
                hide_index=True,
            )

    with app.expander("View audit log"):
        app.dataframe(_audit_rows(log.all_records()), use_container_width=True)


def main() -> None:
    _initialise_session_state()
    _render_header()
    _render_nav_bar()
    raw_upload, datamap_upload = _render_upload_section()
    _handle_run(raw_upload, datamap_upload)
    _section_anchor("section-filter")
    _render_results_section()
    _section_ai_analysis()


if __name__ == "__main__":
    main()
