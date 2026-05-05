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
    "cross_cut_results": [],
    "cross_cut_skips": [],
    "dataframe": None,
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

    # Clear any prior session cross-cut state BEFORE export so a new pipeline
    # run never carries stale cross-cut artifacts from a previous dataset.
    app = _require_streamlit()
    app.session_state["cross_cut_results"] = []
    app.session_state["cross_cut_skips"] = []

    status.update(label="Exporting workbook...", state="running")
    output_path = "/tmp/survey_analysis.xlsx"
    export_single_cuts(
        results=results,
        skips=skips,
        schema=schema,
        quality_report=quality_report,
        log=log,
        output_path=output_path,
        cross_cut_results=[],
        cross_cut_skips=[],
    )

    app.session_state["results"] = results
    app.session_state["skips"] = skips
    app.session_state["schema"] = schema
    app.session_state["dataframe"] = dataframe
    app.session_state["quality_report"] = quality_report
    app.session_state["log"] = log
    app.session_state["output_path"] = output_path
    app.session_state["run_complete"] = True
    status.update(label="Analysis complete.", state="complete")


def _reexport_workbook() -> None:
    """Re-write the Excel workbook with the current single + cross-cut state."""
    from src.excel_exporter import export_single_cuts

    app = _require_streamlit()
    if not app.session_state.get("run_complete"):
        return
    output_path = app.session_state.get("output_path")
    if not output_path:
        return
    try:
        export_single_cuts(
            results=app.session_state["results"],
            skips=app.session_state["skips"],
            schema=app.session_state["schema"],
            quality_report=app.session_state["quality_report"],
            log=app.session_state["log"],
            output_path=output_path,
            cross_cut_results=app.session_state.get("cross_cut_results", []),
            cross_cut_skips=app.session_state.get("cross_cut_skips", []),
        )
    except Exception as exc:  # noqa: BLE001
        app.warning(f"Could not refresh workbook: {type(exc).__name__}: {exc}")


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


_CATEGORICAL_FILTER_TYPES = {"SINGLE_SELECT", "DEMOGRAPHIC_OR_SEGMENT"}


def _question_label(question: Any, max_text: int = 60) -> str:
    text = (question.question_text or "").strip().replace("\n", " ")
    if len(text) > max_text:
        text = text[: max_text - 1] + "\u2026"
    return f"{question.canonical_id} \u2014 {text}" if text else question.canonical_id


def _eligible_for_cross_cut(question: Any) -> bool:
    qt = question.question_type.name
    return qt in {
        "SINGLE_SELECT",
        "DEMOGRAPHIC_OR_SEGMENT",
        "DIRECT_NUMERIC",
        "MULTI_SELECT_BINARY",
        "NUMERIC_ALLOCATION",
        "GRID_SINGLE_SELECT",
    }


def _compatible_analysis_types(q1: Any, q2: Any) -> list[str]:
    """Return list of AnalysisType.name values valid for (q1, q2).

    Day 9 guardrails:
      - MULTI_SELECT_BINARY excluded from CROSS_TAB
      - NUMERIC_ALLOCATION excluded from GROUP_COMPARISON
    """
    if q1 is None or q2 is None or q1.canonical_id == q2.canonical_id:
        return []
    t1 = q1.question_type.name
    t2 = q2.question_type.name
    options: list[str] = []

    # CROSS_TAB: both single-select / demo (multi-select excluded by guardrail)
    if t1 in _CATEGORICAL_FILTER_TYPES and t2 in _CATEGORICAL_FILTER_TYPES:
        options.append("CROSS_TAB")

    # SEGMENT_PROFILE: at least one side must be categorical (filter source);
    # the other side is the target. Builder will normalize so categorical is q1.
    if t1 in _CATEGORICAL_FILTER_TYPES or t2 in _CATEGORICAL_FILTER_TYPES:
        options.append("SEGMENT_PROFILE")

    # GROUP_COMPARISON: one categorical, one DIRECT_NUMERIC (allocation excluded)
    is_cat = lambda t: t in _CATEGORICAL_FILTER_TYPES  # noqa: E731
    if (is_cat(t1) and t2 == "DIRECT_NUMERIC") or (is_cat(t2) and t1 == "DIRECT_NUMERIC"):
        options.append("GROUP_COMPARISON")

    # EXPECTED_VS_REALIZED: both DIRECT_NUMERIC
    if t1 == "DIRECT_NUMERIC" and t2 == "DIRECT_NUMERIC":
        options.append("EXPECTED_VS_REALIZED")

    return options


def _build_cross_cut_spec(
    cross_cut_id: str,
    title: str,
    analysis_type_name: str,
    q1: Any,
    q2: Any,
    filter_value: Any | None = None,
) -> Any:
    from src.models import AnalysisType, CrossCutSpec

    analysis_type = AnalysisType[analysis_type_name]

    # Normalize source order so the engine receives them as expected.
    if analysis_type is AnalysisType.GROUP_COMPARISON:
        if q1.question_type.name == "DIRECT_NUMERIC":
            q1, q2 = q2, q1  # (segment, metric)
    elif analysis_type is AnalysisType.SEGMENT_PROFILE:
        if q1.question_type.name not in _CATEGORICAL_FILTER_TYPES:
            q1, q2 = q2, q1  # categorical filter source must be q1

    # Rebuild title AFTER any swap so what users see matches what runs.
    title = f"{q1.canonical_id} \u00d7 {q2.canonical_id} ({analysis_type.name})"

    filter_expr: str | None = None
    filter_desc: str | None = None
    if analysis_type is AnalysisType.SEGMENT_PROFILE:
        if filter_value is None:
            raise ValueError("SEGMENT_PROFILE requires a filter value")
        filter_expr = f"{q1.canonical_id} == {filter_value}"
        label = q1.option_map.get(filter_value, str(filter_value))
        filter_desc = f"{q1.canonical_id} = {label}"

    return CrossCutSpec(
        cross_cut_id=cross_cut_id,
        title=title,
        analysis_type=analysis_type,
        source_question_ids=(q1.canonical_id, q2.canonical_id),
        filter_expr=filter_expr,
        filter_mask_description=filter_desc,
    )


def _run_one_cross_cut(spec: Any) -> tuple[Any | None, str | None]:
    from src.calculation_log import CalculationLog
    from src.cross_cut_engine import compute_cross_cuts

    app = _require_streamlit()
    schema = app.session_state["schema"]
    dataframe = app.session_state["dataframe"]
    log = app.session_state["log"] or CalculationLog()
    try:
        results, skips = compute_cross_cuts([spec], schema, dataframe, log)
    except Exception as exc:  # noqa: BLE001
        return None, f"{type(exc).__name__}: {exc}"
    # Capture any skips into session state so the exporter records them.
    if skips:
        app.session_state.setdefault("cross_cut_skips", []).extend(skips)
        _reexport_workbook()
        s = skips[0]
        return None, f"{s.skip_reason}: {s.details or ''}".strip(": ")
    if not results:
        return None, "No result produced"
    return results[0], None


def _result_table_to_frame(result: Any) -> Any:
    """Render the result_table dict as a flat DataFrame for st.dataframe."""
    import pandas as pd

    rt = result.result_table
    name = result.analysis_type.name

    if name == "CROSS_TAB":
        counts = rt.get("counts", {})
        row_labels = rt.get("row_label_map", {})
        col_labels = rt.get("column_label_map", {})
        if not counts:
            return pd.DataFrame()
        rows = []
        for row_key, col_dict in counts.items():
            row = {"": row_labels.get(row_key, str(row_key))}
            for col_key, value in col_dict.items():
                row[col_labels.get(col_key, str(col_key))] = value
            rows.append(row)
        df = pd.DataFrame(rows)
        totals = rt.get("row_totals", {})
        if totals:
            df["Total"] = [totals.get(k, 0) for k in counts.keys()]
        return df

    if name == "SEGMENT_PROFILE":
        target = rt.get("target_result", {})
        return pd.DataFrame(
            [
                {"Field": "Filter", "Value": rt.get("filter_expr", "")},
                {"Field": "Filter N", "Value": rt.get("filter_n", 0)},
                {"Field": "Target question", "Value": rt.get("target_question_id", "")},
                {"Field": "Target valid N", "Value": target.get("valid_n", "")},
                {"Field": "Target missing N", "Value": target.get("missing_n", "")},
            ]
        )

    if name == "GROUP_COMPARISON":
        per = rt.get("per_segment", {})
        rows = [
            {
                "Segment": payload.get("label", str(seg_key)),
                "N": payload.get("n"),
                "Mean": payload.get("mean"),
                "Median": payload.get("median"),
                "Std": payload.get("std"),
            }
            for seg_key, payload in per.items()
        ]
        overall = rt.get("overall", {})
        rows.append(
            {
                "Segment": "OVERALL",
                "N": overall.get("n"),
                "Mean": overall.get("mean"),
                "Median": overall.get("median"),
                "Std": overall.get("std"),
            }
        )
        return pd.DataFrame(rows)

    if name == "EXPECTED_VS_REALIZED":
        expected = rt.get("expected", {})
        realized = rt.get("realized", {})
        gap = rt.get("gap", {})
        return pd.DataFrame(
            [
                {"Series": "Expected", **{k: expected.get(k) for k in ("valid_n", "mean", "median", "std")}},
                {"Series": "Realized", **{k: realized.get(k) for k in ("valid_n", "mean", "median", "std")}},
                {"Series": "Gap (R-E)", **{k: gap.get(k) for k in ("valid_n", "mean", "median", "std")}},
            ]
        )

    return pd.DataFrame([{"info": "No display formatter for this analysis type"}])


def _try_import_suggestions() -> Any | None:
    try:
        from src import cross_cut_suggestions  # type: ignore
        return cross_cut_suggestions
    except Exception:
        return None


def _render_manual_builder(eligible: tuple[Any, ...]) -> None:
    app = _require_streamlit()
    app.markdown("**Build a cross cut**")

    if len(eligible) < 2:
        app.info("At least two analysis-eligible questions are required.")
        return

    labels = [_question_label(q) for q in eligible]
    id_to_q = {q.canonical_id: q for q in eligible}
    default_idx2 = 1 if len(eligible) > 1 else 0

    col1, col2, col3 = app.columns(3)
    with col1:
        q1_label = app.selectbox("Question 1", labels, key="cc_q1")
    with col2:
        q2_label = app.selectbox("Question 2", labels, index=default_idx2, key="cc_q2")
    q1 = id_to_q[q1_label.split(" \u2014 ")[0].split(" — ")[0]]
    q2 = id_to_q[q2_label.split(" \u2014 ")[0].split(" — ")[0]]

    compat = _compatible_analysis_types(q1, q2)
    with col3:
        if compat:
            atype = app.selectbox("Analysis type", compat, key="cc_atype")
        else:
            app.selectbox(
                "Analysis type",
                ["(no compatible analysis)"],
                disabled=True,
                key="cc_atype_disabled",
                help=(
                    "Pick two questions with compatible types. "
                    "Multi-select cross-tab and allocation group comparison "
                    "are not yet supported."
                ),
            )
            atype = None

    filter_value: Any | None = None
    if atype == "SEGMENT_PROFILE":
        # Pick whichever side is categorical as the filter source.
        if q1.question_type.name in _CATEGORICAL_FILTER_TYPES:
            filter_q = q1
        elif q2.question_type.name in _CATEGORICAL_FILTER_TYPES:
            filter_q = q2
        else:
            filter_q = None
        if filter_q is None or not filter_q.option_map:
            app.warning(
                "No categorical question with options is available "
                "to filter on. Pick a different pair."
            )
        else:
            options = list(filter_q.option_map.items())
            filter_value = app.selectbox(
                f"Filter on value of {filter_q.canonical_id}",
                options=[k for k, _ in options],
                format_func=lambda k: f"{k} \u2014 {filter_q.option_map.get(k, '')}",
                key="cc_filter_value",
            )

    can_run = atype is not None and (
        atype != "SEGMENT_PROFILE" or filter_value is not None
    )
    if app.button("Run cross cut", type="primary", disabled=not can_run):
        try:
            cc_id = f"CC_MAN_{len(app.session_state['cross_cut_results']) + 1}"
            title = f"{q1.canonical_id} \u00d7 {q2.canonical_id} ({atype})"
            spec = _build_cross_cut_spec(cc_id, title, atype, q1, q2, filter_value)
            result, error = _run_one_cross_cut(spec)
            if error:
                app.error(error)
            elif result is not None:
                app.session_state["cross_cut_results"].append(result)
                _reexport_workbook()
                app.rerun()
        except Exception as exc:  # noqa: BLE001
            app.error(f"{type(exc).__name__}: {exc}")


def _render_suggestions(eligible: tuple[Any, ...]) -> None:
    app = _require_streamlit()
    show = app.checkbox(
        "Show suggested cross cuts",
        value=False,
        key="show_suggestions",
        help="Opt-in: ask the suggestion engine for cross cuts worth running.",
    )
    if not show:
        return

    module = _try_import_suggestions()
    if module is None:
        app.info(
            "Suggestion engine not installed yet "
            "(src/cross_cut_suggestions.py is pending Codex delivery)."
        )
        return

    schema = app.session_state["schema"]
    try:
        suggestions = module.suggest_cross_cuts(schema)
    except Exception as exc:  # noqa: BLE001
        app.error(f"Suggestion engine error: {type(exc).__name__}: {exc}")
        return

    if not suggestions:
        app.info("The suggestion engine returned no suggestions.")
        return

    for idx, suggestion in enumerate(suggestions):
        # suggest_cross_cuts() returns list[tuple[CrossCutSpec, str]]
        if isinstance(suggestion, tuple) and len(suggestion) >= 2:
            spec, reason = suggestion[0], suggestion[1]
        else:
            spec = getattr(suggestion, "spec", suggestion)
            reason = getattr(suggestion, "reason", "") or getattr(
                suggestion, "rationale", ""
            )
        title = getattr(spec, "title", None) or getattr(
            spec, "cross_cut_id", f"Suggestion {idx + 1}"
        )
        atype = getattr(spec, "analysis_type", None)
        atype_name = atype.name if atype is not None else "?"

        with app.container(border=True):
            cols = app.columns([4, 1])
            with cols[0]:
                app.markdown(f"**{title}**  &nbsp; `{atype_name}`")
                if reason:
                    app.caption(reason)
            with cols[1]:
                if app.button("Run this", key=f"sugg_run_{idx}"):
                    result, error = _run_one_cross_cut(spec)
                    if error:
                        app.error(error)
                    elif result is not None:
                        app.session_state["cross_cut_results"].append(result)
                        _reexport_workbook()
                        app.rerun()


def _render_cross_cut_results() -> None:
    app = _require_streamlit()
    results = app.session_state["cross_cut_results"]
    if not results:
        app.caption("No cross cuts run yet.")
        return

    app.markdown(f"**Cross cut results ({len(results)})**")
    for idx, result in enumerate(list(results)):
        with app.expander(
            f"{result.cross_cut_id} \u2014 "
            f"{result.synthetic_question_title} "
            f"[{result.analysis_type.name}]",
            expanded=(idx == len(results) - 1),
        ):
            try:
                frame = _result_table_to_frame(result)
                app.dataframe(frame, use_container_width=True)
            except Exception as exc:  # noqa: BLE001
                app.error(f"Could not render result_table: {exc}")
                app.json(result.result_table)
            if app.button("Remove", key=f"cc_remove_{result.cross_cut_id}_{idx}"):
                app.session_state["cross_cut_results"].pop(idx)
                _reexport_workbook()
                app.rerun()


def _render_cross_cuts_section() -> None:
    app = _require_streamlit()
    if not app.session_state["run_complete"]:
        return
    schema = app.session_state["schema"]
    if schema is None or app.session_state["dataframe"] is None:
        return

    app.subheader("Cross cuts")
    app.caption(
        "Combine two questions to compute deterministic cross-question "
        "analyses on top of the single-cut run."
    )

    eligible_all = schema.analysis_eligible_questions()
    eligible = tuple(q for q in eligible_all if _eligible_for_cross_cut(q))

    _render_suggestions(eligible)
    _render_manual_builder(eligible)
    _render_cross_cut_results()


def main() -> None:
    _initialise_session_state()
    _render_header()
    raw_upload, datamap_upload = _render_upload_section()
    _handle_run(raw_upload, datamap_upload)
    _render_results_section()
    _render_cross_cuts_section()


if __name__ == "__main__":
    main()
