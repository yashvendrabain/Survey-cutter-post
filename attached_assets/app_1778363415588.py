"""Streamlit entry point for the Survey Insight Engine."""

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

try:
    from config import APP_NAME, VERSION
except (ModuleNotFoundError, ImportError):  # pragma: no cover - target app provides config.
    APP_NAME = "Survey Insight Engine"
    VERSION = "Stage 2"


SESSION_DEFAULTS = {
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


def _render_upload_section() -> tuple[Any | None, Any | None]:
    app = _require_streamlit()
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


def _cross_cut_rows(results: list[Any]) -> list[dict[str, Any]]:
    return [
        {
            "Cross Cut ID": result.cross_cut_id,
            "Title": result.synthetic_question_title,
            "Type": result.analysis_type.value,
            "Source Questions": ", ".join(result.source_question_ids),
            "Warnings": " | ".join(result.warnings),
        }
        for result in results
    ]


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
                app.markdown(f"**{result.synthetic_question_title}**")
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
        app.dataframe(_cross_cut_rows(results), use_container_width=True)


def _render_results_section() -> None:
    app = _require_streamlit()
    if not app.session_state["run_complete"]:
        return

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

    with app.expander("View audit log"):
        app.dataframe(_audit_rows(log.all_records()), use_container_width=True)


def main() -> None:
    _initialise_session_state()
    _render_header()
    raw_upload, datamap_upload = _render_upload_section()
    _handle_run(raw_upload, datamap_upload)
    _render_results_section()


if __name__ == "__main__":
    main()
