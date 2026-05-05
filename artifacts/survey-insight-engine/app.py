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
    VERSION = "Stage 1"


SESSION_DEFAULTS = {
    "results": [],
    "skips": [],
    "schema": None,
    "quality_report": None,
    "log": None,
    "output_path": None,
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
    app.session_state["results"] = results
    app.session_state["skips"] = skips
    app.session_state["schema"] = schema
    app.session_state["quality_report"] = quality_report
    app.session_state["log"] = log
    app.session_state["output_path"] = output_path
    app.session_state["run_complete"] = True
    status.update(label="Analysis complete.", state="complete")


def _render_header() -> None:
    app = _require_streamlit()
    app.set_page_config(page_title=APP_NAME, layout="wide")
    app.title(APP_NAME)
    app.caption(
        "Upload raw survey data and a data map to produce audited "
        "single-cut analysis in Excel."
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
