"""Excel workbook exporter for Survey Insight Engine single cuts."""

from __future__ import annotations

from collections import defaultdict
from datetime import datetime, timezone
import math
from pathlib import Path
from typing import Any
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile

try:
    import xlsxwriter
except ModuleNotFoundError:
    xlsxwriter = None

from src.calculation_log import CalculationLog
from src.models import (
    AuditRecord,
    AnalysisType,
    CrossCutResult,
    DataQualityReport,
    FilteredSingleCutResult,
    FilterSpec,
    GridSingleSelectResult,
    MultiSelectResult,
    NumericResult,
    QuestionType,
    SingleCutResult,
    SingleSelectResult,
    SkipRecord,
    SurveySchema,
)


def export_single_cuts(
    results: list[SingleCutResult],
    skips: list[SkipRecord],
    schema: SurveySchema,
    quality_report: DataQualityReport,
    log: CalculationLog,
    output_path: str,
    cross_cut_results: list[CrossCutResult] | None = None,
    cross_cut_skips: list[SkipRecord] | None = None,
) -> None:
    """Write a complete single-cut Excel workbook."""

    cross_cut_results = cross_cut_results or []
    cross_cut_skips = cross_cut_skips or []
    workbook = _create_workbook(output_path)
    formats = _make_formats(workbook)
    used_sheet_names: set[str] = set()
    sc_sheet_names = {
        result.question_id: _unique_sheet_name(
            f"SC_{result.question_id}", used_sheet_names
        )
        for result in results
    }
    cc_sheet_names = {
        result.cross_cut_id: _unique_sheet_name(
            f"CC_{result.cross_cut_id}", used_sheet_names
        )
        for result in cross_cut_results
    }

    _write_run_summary(
        workbook,
        schema,
        quality_report,
        results,
        skips,
        log,
        formats,
        cross_cut_results,
        cross_cut_skips,
    )
    _write_question_metadata(workbook, schema, formats)
    _write_single_cut_index(workbook, results, sc_sheet_names, formats)
    _write_skip_log(workbook, skips, cross_cut_skips, formats)
    for result in results:
        _write_sc_sheet(workbook, result, schema, formats, sc_sheet_names[result.question_id])
    if cross_cut_results:
        _write_cross_cut_index(workbook, cross_cut_results, cc_sheet_names, formats)
        for result in cross_cut_results:
            _write_cc_sheet(
                workbook,
                result,
                schema,
                formats,
                cc_sheet_names[result.cross_cut_id],
            )
    _write_calculation_log(workbook, log, formats)
    _write_filter_log(workbook, cross_cut_results, cc_sheet_names, formats)
    _write_data_quality(workbook, quality_report, formats)
    _write_warnings(
        workbook,
        quality_report,
        results,
        skips,
        formats,
        cross_cut_results,
        cross_cut_skips,
    )
    workbook.close()


def export_cross_cuts_only(
    cross_cut_results: list[CrossCutResult],
    schema: SurveySchema,
    log: CalculationLog,
    output_path: str,
) -> None:
    """Write a workbook containing only selected cross-cut outputs."""

    workbook = _create_workbook(output_path)
    formats = _make_formats(workbook)
    used_sheet_names: set[str] = set()
    cc_sheet_names = {
        result.cross_cut_id: _unique_sheet_name(
            f"CC_{result.cross_cut_id}", used_sheet_names
        )
        for result in cross_cut_results
    }

    _write_cross_cut_index(workbook, cross_cut_results, cc_sheet_names, formats)
    for result in cross_cut_results:
        _write_cc_sheet(
            workbook,
            result,
            schema,
            formats,
            cc_sheet_names[result.cross_cut_id],
        )
    _write_calculation_log(workbook, log, formats)
    _write_filter_log(workbook, cross_cut_results, cc_sheet_names, formats)
    workbook.close()


def export_filtered_single_cuts(
    filtered_results: list[FilteredSingleCutResult],
    schema: SurveySchema,
    log: CalculationLog,
    output_path: str,
) -> None:
    """Write a workbook of filtered single-cut analyses."""

    workbook = _create_workbook(output_path)
    formats = _make_formats(workbook)
    used_sheet_names: set[str] = set()
    fsc_sheet_names = _filtered_sheet_names(filtered_results, used_sheet_names)

    _write_filtered_run_summary(workbook, filtered_results, schema, log, formats)
    _write_filtered_cut_index(workbook, filtered_results, fsc_sheet_names, schema, formats)
    for index, result in enumerate(filtered_results):
        _write_fsc_sheet(workbook, result, schema, formats, fsc_sheet_names[index])
    _write_filtered_calculation_log(workbook, filtered_results, log, formats)
    _write_filtered_filter_log(workbook, filtered_results, fsc_sheet_names, schema, formats)
    workbook.close()


def _create_workbook(output_path: str) -> Any:
    if xlsxwriter is not None:
        return xlsxwriter.Workbook(
            output_path,
            {"nan_inf_to_errors": True, "strings_to_formulas": False},
        )
    return _FallbackWorkbook(output_path)


def _make_formats(workbook: Any) -> dict[str, Any]:
    return {
        "bold": workbook.add_format({"bold": True}),
        "italic": workbook.add_format({"italic": True}),
        "bold_italic": workbook.add_format({"bold": True, "italic": True}),
        "header": workbook.add_format(
            {
                "bold": True,
                "bg_color": "#F2F2F2",
                "border": 1,
            }
        ),
        "pct": workbook.add_format({"num_format": "0.0%"}),
        "count": workbook.add_format({"num_format": "#,##0"}),
        "stat": workbook.add_format({"num_format": "0.00"}),
        "link": workbook.add_format({"font_color": "blue", "underline": 1}),
    }


def _write_run_summary(
    workbook: Any,
    schema: SurveySchema,
    quality_report: DataQualityReport,
    results: list[SingleCutResult],
    skips: list[SkipRecord],
    log: CalculationLog,
    formats: dict[str, Any],
    cross_cut_results: list[CrossCutResult],
    cross_cut_skips: list[SkipRecord],
) -> None:
    ws = workbook.add_worksheet("Run_Summary")
    filter_log_entries = _filter_log_entry_count(cross_cut_results)
    rows = [
        ("Source datamap:", schema.source_datamap_path),
        ("Source raw data:", schema.source_rawdata_path),
        ("Run timestamp:", schema.parsed_at.isoformat()),
        ("Total respondents:", schema.total_respondents),
        ("Total questions:", len(schema.questions)),
        ("Results produced:", len(results)),
        ("Questions skipped:", len(skips)),
        ("Cross cuts produced:", len(cross_cut_results)),
        ("Cross cut skips:", len(cross_cut_skips)),
        ("Filter Log entries:", filter_log_entries),
        ("Audit log records:", len(log)),
        (
            "Calculation errors:",
            sum(1 for skip in skips if skip.skip_reason == "calculation_error"),
        ),
        ("Quality warnings:", len(quality_report.warnings)),
    ]
    for row_index, (label, value) in enumerate(rows):
        _write(ws, row_index, 0, label, formats["bold"])
        _write(ws, row_index, 1, value)
    ws.set_column(0, 0, 25)
    _autofit(ws)


def _write_question_metadata(
    workbook: Any, schema: SurveySchema, formats: dict[str, Any]
) -> None:
    ws = workbook.add_worksheet("Question_Metadata")
    headers = [
        "Question ID",
        "Canonical ID",
        "Question Text",
        "Type",
        "Raw Columns",
        "Options Count",
        "Analysis Eligible",
        "Exclusion Reason",
        "Parent Question",
    ]
    _write_header_row(ws, 0, headers, formats)
    for row_index, spec in enumerate(schema.questions, start=1):
        values = [
            spec.question_id,
            spec.canonical_id,
            spec.question_text,
            spec.question_type.value,
            ", ".join(spec.raw_columns),
            len(spec.option_map),
            "Yes" if spec.analysis_eligible else "No",
            spec.exclusion_reason or "",
            spec.parent_question_id or "",
        ]
        _write_row(ws, row_index, values)
    ws.freeze_panes(1, 0)
    _autofit(ws)


def _write_single_cut_index(
    workbook: Any,
    results: list[SingleCutResult],
    sc_sheet_names: dict[str, str],
    formats: dict[str, Any],
) -> None:
    ws = workbook.add_worksheet("Single_Cut_Index")
    headers = [
        "Question ID",
        "Canonical ID",
        "Type",
        "Sheet Name",
        "Valid N",
        "Missing N",
        "Missing %",
        "Warnings",
    ]
    _write_header_row(ws, 0, headers, formats)
    for row_index, result in enumerate(results, start=1):
        sheet_name = sc_sheet_names[result.question_id]
        denominator = result.valid_n + result.missing_n
        missing_pct = result.missing_n / denominator if denominator > 0 else 0.0
        _write(ws, row_index, 0, result.question_id)
        _write(ws, row_index, 1, result.question_id)
        _write(ws, row_index, 2, result.question_type.value)
        _write_url(
            ws,
            row_index,
            3,
            f"internal:'{_quote_sheet_name(sheet_name)}'!A1",
            formats["link"],
            sheet_name,
        )
        _write(ws, row_index, 4, result.valid_n, formats["count"])
        _write(ws, row_index, 5, result.missing_n, formats["count"])
        _write(ws, row_index, 6, missing_pct, formats["pct"])
        _write(ws, row_index, 7, " | ".join(result.warnings))
    ws.freeze_panes(1, 0)
    _autofit(ws)


def _write_skip_log(
    workbook: Any,
    skips: list[SkipRecord],
    cross_cut_skips: list[SkipRecord],
    formats: dict[str, Any],
) -> None:
    ws = workbook.add_worksheet("Skip_Log")
    headers = ["Source", "Question/Cut ID", "Type", "Skip Reason", "Details"]
    _write_header_row(ws, 0, headers, formats)
    row_index = 1
    for skip in skips:
        _write_row(
            ws,
            row_index,
            [
                "single_cut",
                skip.canonical_id,
                skip.question_type.value,
                skip.skip_reason,
                skip.details or "",
            ],
        )
        row_index += 1
    for skip in cross_cut_skips:
        _write_row(
            ws,
            row_index,
            [
                "cross_cut",
                skip.canonical_id,
                skip.question_type.value,
                skip.skip_reason,
                skip.details or "",
            ],
        )
        row_index += 1
    ws.freeze_panes(1, 0)
    _autofit(ws)


def _write_sc_sheet(
    workbook: Any,
    result: SingleCutResult,
    schema: SurveySchema,
    formats: dict[str, Any],
    sheet_name: str,
) -> None:
    spec = schema.get_question(result.question_id)
    question_text = spec.question_text if spec is not None else result.question_id
    ws = workbook.add_worksheet(sheet_name)

    _write(ws, 0, 0, "Question:", formats["bold"])
    _write(ws, 0, 1, question_text)
    _write(ws, 1, 0, "Type:", formats["bold"])
    _write(ws, 1, 1, result.question_type.value)
    _write(ws, 2, 0, "Denominator:", formats["bold"])
    _write(ws, 2, 1, _denominator_description(result))
    _write(ws, 3, 0, "Valid N:", formats["bold"])
    _write(ws, 3, 1, result.valid_n, formats["count"])
    _write(ws, 4, 0, "Missing N:", formats["bold"])
    _write(ws, 4, 1, result.missing_n, formats["count"])

    if isinstance(result, GridSingleSelectResult):
        _write_grid_body(ws, result, schema, formats, start_row=6)
    elif isinstance(result, SingleSelectResult):
        _write_single_select_body(ws, result, formats, start_row=6)
    elif isinstance(result, MultiSelectResult):
        _write_multi_select_body(ws, result, formats, start_row=6)
    elif isinstance(result, NumericResult):
        _write_numeric_body(ws, result, formats, start_row=6)

    if result.warnings:
        last_row = _find_last_used_row(ws) + 2
        _write(ws, last_row, 0, "Warnings:", formats["bold"])
        for offset, warning in enumerate(result.warnings, start=1):
            _write(ws, last_row + offset, 0, warning)

    ws.freeze_panes(6, 0)
    _autofit(ws)


def _write_single_select_body(
    ws: Any,
    result: SingleSelectResult,
    formats: dict[str, Any],
    start_row: int,
) -> int:
    _write_header_row(ws, start_row, ["Code", "Label", "Count", "%"], formats)
    row_index = start_row + 1
    total_count = 0
    total_rate = 0.0
    for code, payload in sorted(result.distribution.items(), key=lambda item: _sort_key(item[0])):
        total_count += int(payload["count"])
        total_rate += float(payload["rate"])
        _write(ws, row_index, 0, code)
        _write(ws, row_index, 1, payload["label"])
        _write(ws, row_index, 2, payload["count"], formats["count"])
        _write(ws, row_index, 3, payload["rate"], formats["pct"])
        row_index += 1
    total_rate = 1.0 if math.isclose(total_rate, 1.0) else total_rate
    _write(ws, row_index, 1, "Total", formats["bold"])
    _write(ws, row_index, 2, total_count, formats["count"])
    _write(ws, row_index, 3, total_rate, formats["pct"])
    return row_index


def _write_multi_select_body(
    ws: Any,
    result: MultiSelectResult,
    formats: dict[str, Any],
    start_row: int,
) -> int:
    _write_header_row(
        ws, start_row, ["Sub-Column ID", "Label", "Count", "Selection %"], formats
    )
    row_index = start_row + 1
    sorted_items = sorted(
        result.selections.items(),
        key=lambda item: _rate_sort_value(item[1]["selection_rate"]),
        reverse=True,
    )
    for sub_column_id, payload in sorted_items:
        _write(ws, row_index, 0, sub_column_id)
        _write(ws, row_index, 1, payload["label"])
        _write(ws, row_index, 2, payload["count"], formats["count"])
        _write(ws, row_index, 3, payload["selection_rate"], formats["pct"])
        row_index += 1
    row_index += 1
    _write(
        ws,
        row_index,
        0,
        f"Respondents who answered any: {result.respondents_who_answered_any}",
        formats["bold"],
    )
    return row_index


def _write_numeric_body(
    ws: Any,
    result: NumericResult,
    formats: dict[str, Any],
    start_row: int,
) -> int:
    rows = [
        ("Mean", result.mean),
        ("Median", result.median),
        ("Std Dev", result.std),
        ("Min", result.min_val),
        ("Max", result.max_val),
        ("25th percentile", result.percentiles[25]),
        ("50th percentile", result.percentiles[50]),
        ("75th percentile", result.percentiles[75]),
    ]
    row_index = start_row
    for label, value in rows:
        _write(ws, row_index, 0, label)
        _write(ws, row_index, 1, value, formats["stat"])
        row_index += 1

    if result.question_type is QuestionType.NUMERIC_ALLOCATION:
        row_index += 1
        _write(ws, row_index, 0, "Per-option allocation means", formats["bold"])
        row_index += 1
        _write_header_row(ws, row_index, ["Option", "Mean", "Median"], formats)
        row_index += 1
        for option_id, payload in (result.per_option_stats or {}).items():
            _write(ws, row_index, 0, option_id)
            _write(ws, row_index, 1, payload["mean"], formats["stat"])
            _write(ws, row_index, 2, payload["median"], formats["stat"])
            row_index += 1
        row_index += 1
        allocation_rows = [
            ("Allocation target:", result.allocation_target),
            ("Tolerance:", result.allocation_tolerance),
            ("Excluded (out of tolerance):", result.allocation_excluded_n),
        ]
        for label, value in allocation_rows:
            _write(ws, row_index, 0, label)
            _write(ws, row_index, 1, value, formats["stat"])
            row_index += 1
    return row_index


def _write_grid_body(
    ws: Any,
    result: GridSingleSelectResult,
    schema: SurveySchema,
    formats: dict[str, Any],
    start_row: int,
) -> int:
    spec = schema.get_question(result.question_id)
    row_labels = spec.grid_row_labels if spec and spec.grid_row_labels else {}
    row_index = start_row
    _write(ws, row_index, 0, "Per-row distributions", formats["bold"])
    row_index += 2
    for sub_column_id, row_result in result.rows.items():
        row_label = row_labels.get(sub_column_id, sub_column_id)
        _write(
            ws,
            row_index,
            0,
            f"{row_label} (n={row_result.valid_n})",
            formats["bold"],
        )
        row_index += 1
        row_index = _write_single_select_body(ws, row_result, formats, row_index) + 2
    _write(ws, row_index, 0, "Overall valid N:", formats["bold"])
    _write(ws, row_index, 1, result.overall_valid_n, formats["count"])
    return row_index


def _write_cross_cut_index(
    workbook: Any,
    cross_cut_results: list[CrossCutResult],
    cc_sheet_names: dict[str, str],
    formats: dict[str, Any],
) -> None:
    ws = workbook.add_worksheet("Cross_Cut_Index")
    headers = [
        "Cross Cut ID",
        "Title",
        "Analysis Type",
        "Sheet Name",
        "Source Questions",
        "Filter",
        "Audit Records",
        "Warnings",
    ]
    _write_header_row(ws, 0, headers, formats)
    for row_index, result in enumerate(cross_cut_results, start=1):
        sheet_name = cc_sheet_names[result.cross_cut_id]
        filter_expr = (
            result.result_table.get("filter_expr", "")
            if result.analysis_type is AnalysisType.SEGMENT_PROFILE
            else ""
        )
        _write(ws, row_index, 0, result.cross_cut_id)
        _write(ws, row_index, 1, result.synthetic_question_title)
        _write(ws, row_index, 2, result.analysis_type.value)
        _write_url(
            ws,
            row_index,
            3,
            f"internal:'{_quote_sheet_name(sheet_name)}'!A1",
            formats["link"],
            sheet_name,
        )
        _write(ws, row_index, 4, ", ".join(result.source_question_ids))
        _write(ws, row_index, 5, filter_expr)
        _write(ws, row_index, 6, len(result.audit_records), formats["count"])
        _write(ws, row_index, 7, " | ".join(result.warnings))
    ws.freeze_panes(1, 0)
    _autofit(ws)


def _write_filtered_run_summary(
    workbook: Any,
    filtered_results: list[FilteredSingleCutResult],
    schema: SurveySchema,
    log: CalculationLog,
    formats: dict[str, Any],
) -> None:
    ws = workbook.add_worksheet("Run_Summary")
    rows = [
        ("Workbook type:", "Filtered single cuts"),
        ("Source datamap:", schema.source_datamap_path),
        ("Source raw data:", schema.source_rawdata_path),
        ("Run timestamp:", schema.parsed_at.isoformat()),
        ("Filtered analyses included:", len(filtered_results)),
        ("Audit log records:", len(log)),
    ]
    for row_index, (label, value) in enumerate(rows):
        _write(ws, row_index, 0, label, formats["bold"])
        _write(ws, row_index, 1, value)
    ws.set_column(0, 0, 30)
    _autofit(ws)


def _write_filtered_cut_index(
    workbook: Any,
    filtered_results: list[FilteredSingleCutResult],
    fsc_sheet_names: dict[int, str],
    schema: SurveySchema,
    formats: dict[str, Any],
) -> None:
    ws = workbook.add_worksheet("Filtered_Cut_Index")
    headers = [
        "Sheet Name",
        "Target Question",
        "Filters Applied",
        "Dispatch Mode",
        "Filtered N",
        "Warnings",
    ]
    _write_header_row(ws, 0, headers, formats)
    for row_index, result in enumerate(filtered_results, start=1):
        sheet_name = fsc_sheet_names[row_index - 1]
        _write_url(
            ws,
            row_index,
            0,
            f"internal:'{_quote_sheet_name(sheet_name)}'!A1",
            formats["link"],
            sheet_name,
        )
        _write(ws, row_index, 1, result.target_question_id)
        _write(
            ws,
            row_index,
            2,
            " | ".join(
                _filter_description(filter_spec, schema, compact_breakdown=True)
                for filter_spec in result.filters_applied
            ),
        )
        _write(ws, row_index, 3, result.dispatch_mode)
        _write(ws, row_index, 4, result.filtered_n, formats["count"])
        _write(ws, row_index, 5, " | ".join(result.warnings))
    ws.freeze_panes(1, 0)
    _autofit(ws)


def _write_fsc_sheet(
    workbook: Any,
    result: FilteredSingleCutResult,
    schema: SurveySchema,
    formats: dict[str, Any],
    sheet_name: str,
) -> None:
    ws = workbook.add_worksheet(sheet_name)
    target_spec = schema.get_question(result.target_question_id)
    target_text = target_spec.question_text if target_spec else result.target_question_id

    _write(ws, 0, 0, "Target question:", formats["bold"])
    _write(ws, 1, 0, target_text)
    _write(ws, 2, 0, "Target ID:", formats["bold"])
    _write(ws, 3, 0, result.target_question_id)
    _write(ws, 4, 0, "Filters applied:", formats["bold"])

    row_index = 5
    if result.filters_applied:
        for filter_spec in result.filters_applied:
            _write(
                ws,
                row_index,
                0,
                "  " + _filter_description(
                    filter_spec,
                    schema,
                    compact_breakdown=False,
                ),
            )
            row_index += 1
    else:
        _write(ws, row_index, 0, "  <none>")
        row_index += 1

    row_index += 1
    _write(ws, row_index, 0, "Filtered N:", formats["bold"])
    _write(ws, row_index, 1, result.filtered_n, formats["count"])
    row_index += 1
    _write(ws, row_index, 0, "Dispatch mode:", formats["bold"])
    _write(ws, row_index, 1, result.dispatch_mode)
    row_index += 2

    if result.dispatch_mode == "single_cut_filtered":
        if result.single_cut_result is None:
            _write(ws, row_index, 0, "Missing filtered single-cut result", formats["bold"])
            last_row = row_index
        else:
            last_row = _write_single_cut_result_body(
                ws, result.single_cut_result, schema, formats, row_index
            )
    elif result.dispatch_mode == "cross_cut_breakdown":
        if result.cross_cut_result is None:
            _write(ws, row_index, 0, "Missing cross-cut breakdown result", formats["bold"])
            last_row = row_index
        else:
            last_row = _write_cross_cut_result_body(
                ws, result.cross_cut_result, schema, formats, row_index
            )
    else:
        _write(ws, row_index, 0, "Unsupported filtered result mode", formats["bold"])
        last_row = row_index

    if result.warnings:
        last_row += 2
        _write(ws, last_row, 0, "Warnings:", formats["bold"])
        for offset, warning in enumerate(result.warnings, start=1):
            _write(ws, last_row + offset, 0, warning)

    ws.freeze_panes(row_index, 0)
    _autofit(ws)


def _write_single_cut_result_body(
    ws: Any,
    result: SingleCutResult,
    schema: SurveySchema,
    formats: dict[str, Any],
    start_row: int,
) -> int:
    if isinstance(result, GridSingleSelectResult):
        return _write_grid_body(ws, result, schema, formats, start_row)
    if isinstance(result, SingleSelectResult):
        return _write_single_select_body(ws, result, formats, start_row)
    if isinstance(result, MultiSelectResult):
        return _write_multi_select_body(ws, result, formats, start_row)
    if isinstance(result, NumericResult):
        return _write_numeric_body(ws, result, formats, start_row)
    _write(ws, start_row, 0, "Unsupported single-cut result type", formats["bold"])
    return start_row


def _write_cross_cut_result_body(
    ws: Any,
    result: CrossCutResult,
    schema: SurveySchema,
    formats: dict[str, Any],
    start_row: int,
) -> int:
    if result.analysis_type is AnalysisType.CROSS_TAB:
        return _write_cross_tab_body(ws, result, schema, formats, start_row)
    if result.analysis_type is AnalysisType.SEGMENT_PROFILE:
        return _write_segment_profile_body(ws, result, schema, formats, start_row)
    if result.analysis_type is AnalysisType.GROUP_COMPARISON:
        return _write_group_comparison_body(ws, result, schema, formats, start_row)
    if result.analysis_type is AnalysisType.EXPECTED_VS_REALIZED:
        return _write_expected_vs_realized_body(ws, result, schema, formats, start_row)
    _write(ws, start_row, 0, "Unsupported cross-cut result type", formats["bold"])
    return start_row


def _write_cc_sheet(
    workbook: Any,
    result: CrossCutResult,
    schema: SurveySchema,
    formats: dict[str, Any],
    sheet_name: str,
) -> None:
    ws = workbook.add_worksheet(sheet_name)
    filter_expr = result.result_table.get("filter_expr") or "<no filter>"
    header_rows = [
        ("Title:", result.synthetic_question_title),
        ("Analysis type:", result.analysis_type.value),
        ("Source questions:", ", ".join(result.source_question_ids)),
        ("Filter:", filter_expr),
        ("AI insight:", result.ai_insight or "<none>"),
        ("Audit records:", len(result.audit_records)),
    ]
    for row_index, (label, value) in enumerate(header_rows):
        _write(ws, row_index, 0, label, formats["bold"])
        _write(ws, row_index, 1, value)

    if result.analysis_type is AnalysisType.CROSS_TAB:
        last_row = _write_cross_tab_body(ws, result, schema, formats, 7)
    elif result.analysis_type is AnalysisType.SEGMENT_PROFILE:
        last_row = _write_segment_profile_body(ws, result, schema, formats, 7)
    elif result.analysis_type is AnalysisType.GROUP_COMPARISON:
        last_row = _write_group_comparison_body(ws, result, schema, formats, 7)
    elif result.analysis_type is AnalysisType.EXPECTED_VS_REALIZED:
        last_row = _write_expected_vs_realized_body(ws, result, schema, formats, 7)
    else:
        _write(ws, 7, 0, "Unsupported cross-cut result type", formats["bold"])
        last_row = 7

    if result.warnings:
        last_row += 2
        _write(ws, last_row, 0, "Warnings:", formats["bold"])
        for offset, warning in enumerate(result.warnings, start=1):
            _write(ws, last_row + offset, 0, warning)

    ws.freeze_panes(7, 0)
    _autofit(ws)


def _write_cross_tab_body(
    ws: Any,
    result: CrossCutResult,
    schema: SurveySchema,
    formats: dict[str, Any],
    start_row: int,
) -> int:
    result_table = result.result_table
    row_question_id = result_table.get("row_question_id", result.source_question_ids[0])
    column_question_id = result_table.get(
        "column_question_id", result.source_question_ids[1]
    )
    row_question_text = _question_text(schema, row_question_id)
    column_question_text = _question_text(schema, column_question_id)
    _merge_range(
        ws,
        start_row,
        0,
        start_row,
        3,
        f"Rows (vertical): {row_question_id}",
        formats["bold_italic"],
    )
    _merge_range(ws, start_row + 1, 0, start_row + 1, 3, row_question_text)
    _merge_range(
        ws,
        start_row + 2,
        0,
        start_row + 2,
        3,
        f"Columns (horizontal): {column_question_id}",
        formats["bold_italic"],
    )
    _merge_range(ws, start_row + 3, 0, start_row + 3, 3, column_question_text)
    orientation_label = f"↓ {row_question_id}  →  {column_question_id}"
    blocks_to_render = []
    if result.display_mode in ("counts", "both", "all"):
        blocks_to_render.append(("Counts", "counts", formats["count"]))
    if result.display_mode in ("row_pct", "both", "all"):
        blocks_to_render.append(("Row %", "row_pct", formats["pct"]))
    if result.display_mode in ("col_pct", "all"):
        blocks_to_render.append(("Column %", "column_pct", formats["pct"]))

    row_index = start_row + 5
    for block_index, (header, key, cell_format) in enumerate(blocks_to_render):
        if block_index > 0:
            row_index += 2
        row_index = _write_cross_tab_matrix(
            ws,
            header,
            result_table.get(key, {}),
            result_table.get("row_label_map", {}),
            result_table.get("column_label_map", {}),
            formats,
            row_index,
            cell_format,
            orientation_label,
        )
    row_index += 2
    _write(ws, row_index, 0, "Grand total:", formats["bold"])
    _write(ws, row_index, 1, result_table.get("grand_total", ""), formats["count"])
    return row_index


def _write_cross_tab_matrix(
    ws: Any,
    title: str,
    matrix: dict,
    row_label_map: dict,
    column_label_map: dict,
    formats: dict[str, Any],
    start_row: int,
    value_format: Any,
    orientation_label: str,
) -> int:
    row_codes = sorted(matrix.keys(), key=_sort_key)
    column_codes = _cross_tab_column_codes(matrix, column_label_map)
    _write(ws, start_row, 0, title, formats["bold"])
    _write(ws, start_row + 1, 0, orientation_label, formats["italic"])
    _write(ws, start_row + 1, 1, "")
    for offset, column_code in enumerate(column_codes, start=2):
        _write(ws, start_row + 1, offset, column_code, formats["header"])
        _write(
            ws,
            start_row + 2,
            offset,
            _label_for_code(column_label_map, column_code),
            formats["header"],
        )
    row_index = start_row + 3
    for row_code in row_codes:
        _write(ws, row_index, 0, row_code)
        _write(ws, row_index, 1, _label_for_code(row_label_map, row_code))
        for offset, column_code in enumerate(column_codes, start=2):
            _write(
                ws,
                row_index,
                offset,
                _nested_lookup(matrix, row_code, column_code),
                value_format,
            )
        row_index += 1
    return row_index - 1


def _write_segment_profile_body(
    ws: Any,
    result: CrossCutResult,
    schema: SurveySchema,
    formats: dict[str, Any],
    start_row: int,
) -> int:
    result_table = result.result_table
    filter_question_id = result.source_question_ids[0]
    target_question_id = result_table.get(
        "target_question_id", result.source_question_ids[1]
    )
    filter_spec = schema.get_question(filter_question_id)
    _write(ws, start_row, 0, "Filter applied:", formats["bold"])
    _write(ws, start_row + 1, 0, _question_label(schema, filter_question_id))
    _write(
        ws,
        start_row + 2,
        0,
        f"= {_filter_value_label(result_table.get('filter_expr'), filter_spec)}",
    )
    _write(ws, start_row + 3, 0, "Target question:", formats["bold"])
    _write(ws, start_row + 4, 0, _question_label(schema, target_question_id))
    _write(ws, start_row + 6, 0, "Target distribution", formats["bold"])
    return _write_serialized_single_cut_body(
        ws,
        result_table.get("target_result", {}),
        formats,
        start_row + 7,
        _segment_profile_display_mode(result.display_mode),
    )


def _write_serialized_single_cut_body(
    ws: Any,
    payload: dict,
    formats: dict[str, Any],
    start_row: int,
    display_mode: str = "all",
) -> int:
    if "distribution" in payload:
        return _write_distribution_dict(
            ws, payload.get("distribution", {}), formats, start_row, display_mode
        )
    if "selections" in payload:
        return _write_selections_dict(
            ws, payload.get("selections", {}), formats, start_row
        )
    if "rows" in payload:
        return _write_grid_dict(ws, payload.get("rows", {}), formats, start_row)
    if "mean" in payload:
        return _write_numeric_dict(ws, payload, formats, start_row)
    _write(ws, start_row, 0, "Unsupported target result")
    return start_row


def _write_distribution_dict(
    ws: Any,
    distribution: dict,
    formats: dict[str, Any],
    start_row: int,
    display_mode: str = "all",
) -> int:
    headers = ["Code", "Label", "Count"]
    include_rate = display_mode != "counts"
    if include_rate:
        headers.append("%")
    _write_header_row(ws, start_row, headers, formats)
    row_index = start_row + 1
    for code, payload in sorted(distribution.items(), key=lambda item: _sort_key(item[0])):
        _write(ws, row_index, 0, code)
        _write(ws, row_index, 1, payload.get("label", ""))
        _write(ws, row_index, 2, payload.get("count", 0), formats["count"])
        if include_rate:
            _write(ws, row_index, 3, payload.get("rate", 0.0), formats["pct"])
        row_index += 1
    return row_index - 1


def _write_selections_dict(
    ws: Any,
    selections: dict,
    formats: dict[str, Any],
    start_row: int,
) -> int:
    _write_header_row(
        ws, start_row, ["Sub-Column ID", "Label", "Count", "Selection %"], formats
    )
    row_index = start_row + 1
    for sub_column_id, payload in selections.items():
        _write(ws, row_index, 0, sub_column_id)
        _write(ws, row_index, 1, payload.get("label", ""))
        _write(ws, row_index, 2, payload.get("count", 0), formats["count"])
        _write(ws, row_index, 3, payload.get("selection_rate", 0.0), formats["pct"])
        row_index += 1
    return row_index - 1


def _write_numeric_dict(
    ws: Any,
    payload: dict,
    formats: dict[str, Any],
    start_row: int,
) -> int:
    rows = [
        ("Mean", payload.get("mean")),
        ("Median", payload.get("median")),
        ("Std Dev", payload.get("std")),
        ("Min", payload.get("min_val")),
        ("Max", payload.get("max_val")),
        ("25th percentile", payload.get("percentiles", {}).get(25)),
        ("50th percentile", payload.get("percentiles", {}).get(50)),
        ("75th percentile", payload.get("percentiles", {}).get(75)),
    ]
    row_index = start_row
    for label, value in rows:
        _write(ws, row_index, 0, label)
        _write(ws, row_index, 1, value, formats["stat"])
        row_index += 1
    return row_index - 1


def _write_grid_dict(
    ws: Any,
    rows: dict,
    formats: dict[str, Any],
    start_row: int,
) -> int:
    row_index = start_row
    for row_id, row_payload in rows.items():
        _write(
            ws,
            row_index,
            0,
            f"{row_id} (n={row_payload.get('valid_n', 0)})",
            formats["bold"],
        )
        row_index += 1
        row_index = _write_distribution_dict(
            ws,
            row_payload.get("distribution", {}),
            formats,
            row_index,
        ) + 2
    return row_index - 1


def _write_group_comparison_body(
    ws: Any,
    result: CrossCutResult,
    schema: SurveySchema,
    formats: dict[str, Any],
    start_row: int,
) -> int:
    result_table = result.result_table
    segment_question_id = result_table.get(
        "segment_question_id", result.source_question_ids[0]
    )
    metric_question_id = result_table.get(
        "metric_question_id", result.source_question_ids[1]
    )
    _write(ws, start_row, 0, "Segments (rows):", formats["bold"])
    _write(ws, start_row + 1, 0, _question_label(schema, segment_question_id))
    _write(ws, start_row + 2, 0, "Metric (columns):", formats["bold"])
    _write(ws, start_row + 3, 0, _question_label(schema, metric_question_id))
    table_start = start_row + 5
    _write(ws, table_start, 0, "Per-segment comparison", formats["bold"])
    _write_header_row(ws, table_start + 1, ["Segment", "Label", "N", "Mean", "Median", "Std"], formats)
    row_index = table_start + 2
    for segment, payload in result_table.get("per_segment", {}).items():
        _write(ws, row_index, 0, segment)
        _write(ws, row_index, 1, payload.get("label", ""))
        _write(ws, row_index, 2, payload.get("n", 0), formats["count"])
        _write(ws, row_index, 3, payload.get("mean"), formats["stat"])
        _write(ws, row_index, 4, payload.get("median"), formats["stat"])
        _write(ws, row_index, 5, payload.get("std"), formats["stat"])
        row_index += 1
    overall = result_table.get("overall", {})
    _write(ws, row_index, 0, "Overall", formats["bold"])
    _write(ws, row_index, 2, overall.get("n", 0), formats["count"])
    _write(ws, row_index, 3, overall.get("mean"), formats["stat"])
    _write(ws, row_index, 4, overall.get("median"), formats["stat"])
    _write(ws, row_index, 5, overall.get("std"), formats["stat"])
    return row_index


def _write_expected_vs_realized_body(
    ws: Any,
    result: CrossCutResult,
    schema: SurveySchema,
    formats: dict[str, Any],
    start_row: int,
) -> int:
    result_table = result.result_table
    expected = result_table.get("expected", {})
    realized = result_table.get("realized", {})
    gap = result_table.get("gap", {})
    expected_question_id = result_table.get(
        "expected_question_id", result.source_question_ids[0]
    )
    realized_question_id = result_table.get(
        "realized_question_id", result.source_question_ids[1]
    )
    _write(ws, start_row, 0, "Expected:", formats["bold"])
    _write(ws, start_row + 1, 0, _question_label(schema, expected_question_id))
    _write(ws, start_row + 2, 0, "Realized:", formats["bold"])
    _write(ws, start_row + 3, 0, _question_label(schema, realized_question_id))
    table_start = start_row + 5
    _write(ws, table_start, 0, "Expected vs Realized", formats["bold"])
    _write_header_row(ws, table_start + 1, ["Metric", "Expected", "Realized", "Gap"], formats)
    rows = [
        ("Mean", "mean", formats["stat"]),
        ("Median", "median", formats["stat"]),
        ("Std", "std", formats["stat"]),
        ("N", "valid_n", formats["count"]),
    ]
    row_index = table_start + 2
    for label, key, cell_format in rows:
        _write(ws, row_index, 0, label)
        _write(ws, row_index, 1, expected.get(key), cell_format)
        _write(ws, row_index, 2, realized.get(key), cell_format)
        _write(ws, row_index, 3, gap.get(key), cell_format)
        row_index += 1
    row_index += 1
    _write(ws, row_index, 0, f"Paired N: {result_table.get('paired_n', 0)}", formats["bold"])
    return row_index


def _write_calculation_log(
    workbook: Any, log: CalculationLog, formats: dict[str, Any]
) -> None:
    ws = workbook.add_worksheet("Calculation_Log")
    headers = [
        "Output Sheet",
        "Metric Name",
        "Source Question ID",
        "Source Columns",
        "Filter",
        "Numerator",
        "Denominator",
        "Formula",
        "Value Raw",
        "Valid N",
        "Missing N",
        "Timestamp",
    ]
    _write_header_row(ws, 0, headers, formats)
    for row_index, record in enumerate(log.all_records(), start=1):
        _write(ws, row_index, 0, record.output_sheet)
        _write(ws, row_index, 1, record.metric_name)
        _write(ws, row_index, 2, record.source_question_id)
        _write(ws, row_index, 3, ", ".join(record.source_columns))
        _write(ws, row_index, 4, record.filter_expr or "")
        _write(ws, row_index, 5, "" if record.numerator is None else record.numerator, formats["count"])
        _write(ws, row_index, 6, "" if record.denominator is None else record.denominator, formats["count"])
        _write(ws, row_index, 7, record.formula)
        _write(ws, row_index, 8, record.value_raw, formats["stat"])
        _write(ws, row_index, 9, record.valid_n, formats["count"])
        _write(ws, row_index, 10, record.missing_n, formats["count"])
        _write(ws, row_index, 11, record.timestamp.isoformat())
    ws.freeze_panes(1, 0)
    _autofit(ws)


def _write_filter_log(
    workbook: Any,
    cross_cut_results: list[CrossCutResult],
    cc_sheet_names: dict[str, str],
    formats: dict[str, Any],
) -> None:
    ws = workbook.add_worksheet("Filter_Log")
    _write_header_row(
        ws,
        0,
        [
            "Cross Cut ID",
            "Title",
            "Filter Expression",
            "Description",
            "Affects Sheet",
        ],
        formats,
    )
    row_index = 1
    for result in cross_cut_results:
        filter_expr = result.result_table.get("filter_expr")
        if not filter_expr:
            continue
        _write(ws, row_index, 0, result.cross_cut_id)
        _write(ws, row_index, 1, result.synthetic_question_title)
        _write(ws, row_index, 2, filter_expr)
        _write(
            ws,
            row_index,
            3,
            result.result_table.get("filter_mask_description") or filter_expr,
        )
        _write(ws, row_index, 4, cc_sheet_names.get(result.cross_cut_id, ""))
        row_index += 1
    ws.freeze_panes(1, 0)
    _autofit(ws)


def _write_filtered_calculation_log(
    workbook: Any,
    filtered_results: list[FilteredSingleCutResult],
    log: CalculationLog,
    formats: dict[str, Any],
) -> None:
    filtered_log = CalculationLog()
    for record in _filtered_audit_records(filtered_results, log):
        filtered_log.record(record)
    _write_calculation_log(workbook, filtered_log, formats)


def _write_filtered_filter_log(
    workbook: Any,
    filtered_results: list[FilteredSingleCutResult],
    fsc_sheet_names: dict[int, str],
    schema: SurveySchema,
    formats: dict[str, Any],
) -> None:
    ws = workbook.add_worksheet("Filter_Log")
    headers = ["Sheet Name", "Filter Question", "Filter Value", "Filter Description"]
    _write_header_row(ws, 0, headers, formats)
    row_index = 1
    for result_index, result in enumerate(filtered_results):
        sheet_name = fsc_sheet_names[result_index]
        for filter_spec in result.filters_applied:
            _write(ws, row_index, 0, sheet_name)
            _write(ws, row_index, 1, filter_spec.filter_question_id)
            _write(
                ws,
                row_index,
                2,
                "" if filter_spec.filter_value is None else filter_spec.filter_value,
            )
            _write(
                ws,
                row_index,
                3,
                _filter_description(filter_spec, schema, compact_breakdown=False),
            )
            row_index += 1
    ws.freeze_panes(1, 0)
    _autofit(ws)


def _write_data_quality(
    workbook: Any,
    quality_report: DataQualityReport,
    formats: dict[str, Any],
) -> None:
    ws = workbook.add_worksheet("Data_Quality")
    totals = [
        ("Total rows:", quality_report.total_rows),
        ("Total columns:", quality_report.total_columns),
        ("Columns in datamap:", quality_report.columns_in_datamap),
        ("Columns NOT in datamap:", len(quality_report.columns_not_in_datamap)),
    ]
    row_index = 0
    for label, value in totals:
        _write(ws, row_index, 0, label, formats["bold"])
        _write(ws, row_index, 1, value)
        row_index += 1

    row_index += 1
    _write(ws, row_index, 0, "Per-column missing %", formats["bold"])
    row_index += 1
    _write_header_row(ws, row_index, ["Column", "Missing %"], formats)
    row_index += 1
    for column, pct in sorted(
        quality_report.per_column_missing_pct.items(),
        key=lambda item: item[1],
        reverse=True,
    ):
        if pct > 0:
            _write(ws, row_index, 0, column)
            _write(ws, row_index, 1, pct, formats["pct"])
            row_index += 1

    row_index += 1
    _write(ws, row_index, 0, "Per-column out-of-range %", formats["bold"])
    row_index += 1
    _write_header_row(ws, row_index, ["Column", "Out-of-range %"], formats)
    row_index += 1
    for column, pct in quality_report.per_column_out_of_range_pct.items():
        if pct > 0:
            _write(ws, row_index, 0, column)
            _write(ws, row_index, 1, pct, formats["pct"])
            row_index += 1

    row_index += 1
    _write(ws, row_index, 0, "Coercion log", formats["bold"])
    row_index += 1
    _write_header_row(
        ws,
        row_index,
        ["Column", "From Type", "To Type", "Values Coerced", "Rows Affected"],
        formats,
    )
    row_index += 1
    for entry in quality_report.coercion_log:
        coerced_values = ", ".join(str(value) for value in entry.get("values_coerced", ()))
        _write(ws, row_index, 0, entry.get("column", ""))
        _write(ws, row_index, 1, entry.get("from_type", ""))
        _write(ws, row_index, 2, entry.get("to_type", ""))
        _write(ws, row_index, 3, coerced_values[:100])
        _write(ws, row_index, 4, entry.get("rows_affected", ""))
        row_index += 1

    ws.freeze_panes(6, 0)
    _autofit(ws)


def _write_warnings(
    workbook: Any,
    quality_report: DataQualityReport,
    results: list[SingleCutResult],
    skips: list[SkipRecord],
    formats: dict[str, Any],
    cross_cut_results: list[CrossCutResult],
    cross_cut_skips: list[SkipRecord],
) -> None:
    ws = workbook.add_worksheet("Warnings")
    _write_header_row(ws, 0, ["Source", "Warning"], formats)
    row_index = 1
    for warning in quality_report.warnings:
        _write_row(ws, row_index, ["data_quality", warning])
        row_index += 1
    for result in results:
        for warning in result.warnings:
            _write_row(ws, row_index, [f"result:{result.question_id}", warning])
            row_index += 1
    for result in cross_cut_results:
        for warning in result.warnings:
            _write_row(ws, row_index, [f"cross_result:{result.cross_cut_id}", warning])
            row_index += 1
    for skip in skips:
        if skip.details:
            _write_row(ws, row_index, [f"skip:{skip.question_id}", skip.details])
            row_index += 1
    for skip in cross_cut_skips:
        if skip.details:
            _write_row(ws, row_index, [f"cross_skip:{skip.question_id}", skip.details])
            row_index += 1
    ws.freeze_panes(1, 0)
    _autofit(ws)


def _safe_sheet_name(name: str) -> str:
    """Truncate to 31 chars and remove forbidden chars."""
    forbidden = set(":\\/?*[]")
    cleaned = "".join(char for char in name if char not in forbidden)
    cleaned = cleaned.strip() or "Sheet"
    return cleaned[:31]


def _question_text(schema: SurveySchema, question_id: str) -> str:
    spec = schema.get_question(question_id)
    return spec.question_text if spec is not None else question_id


def _question_label(schema: SurveySchema, question_id: str) -> str:
    return f"{question_id}: {_question_text(schema, question_id)}"


def _filter_value_label(filter_expr: str | None, filter_spec: Any) -> str:
    if not filter_expr:
        return "<unknown>"
    value_text = filter_expr.split("==", 1)[-1].strip().strip("\"'")
    try:
        value: int | str = int(value_text)
    except ValueError:
        value = value_text
    if filter_spec is not None and value in filter_spec.option_map:
        return str(filter_spec.option_map[value])
    return str(value)


def _segment_profile_display_mode(display_mode: str) -> str:
    return "counts" if display_mode == "counts" else "both"


def _filter_log_entry_count(cross_cut_results: list[CrossCutResult]) -> int:
    return sum(1 for result in cross_cut_results if result.result_table.get("filter_expr"))


def _filtered_sheet_names(
    filtered_results: list[FilteredSingleCutResult],
    used_sheet_names: set[str],
) -> dict[int, str]:
    target_counts: dict[str, int] = defaultdict(int)
    for result in filtered_results:
        target_counts[result.target_question_id] += 1

    target_seen: dict[str, int] = defaultdict(int)
    sheet_names: dict[int, str] = {}
    for index, result in enumerate(filtered_results):
        target_seen[result.target_question_id] += 1
        base_name = f"FSC_{result.target_question_id}"
        if target_counts[result.target_question_id] > 1:
            base_name = f"{base_name}_{target_seen[result.target_question_id]:02d}"
        sheet_names[index] = _unique_sheet_name(base_name, used_sheet_names)
    return sheet_names


def _filter_description(
    filter_spec: FilterSpec,
    schema: SurveySchema,
    compact_breakdown: bool,
) -> str:
    if filter_spec.filter_value is None:
        if compact_breakdown:
            return f"{filter_spec.filter_question_id} (breakdown)"
        return f"{filter_spec.filter_question_id} (breakdown - no specific value)"

    label = _filter_option_label(filter_spec, schema)
    if label:
        return (
            f"{filter_spec.filter_question_id} == {filter_spec.filter_value} "
            f"({label})"
        )
    return f"{filter_spec.filter_question_id} == {filter_spec.filter_value}"


def _filter_option_label(filter_spec: FilterSpec, schema: SurveySchema) -> str | None:
    spec = schema.get_question(filter_spec.filter_question_id)
    if spec is None:
        return None
    value = filter_spec.filter_value
    if value in spec.option_map:
        return str(spec.option_map[value])
    value_as_string = str(value)
    for option_code, label in spec.option_map.items():
        if str(option_code) == value_as_string:
            return str(label)
    return None


def _filtered_audit_records(
    filtered_results: list[FilteredSingleCutResult],
    log: CalculationLog,
) -> tuple[AuditRecord, ...]:
    source_ids = {
        result.target_question_id
        for result in filtered_results
    }
    for result in filtered_results:
        for filter_spec in result.filters_applied:
            source_ids.add(filter_spec.filter_question_id)

    records = []
    for record in log.all_records():
        if record.source_question_id in source_ids:
            records.append(record)
            continue
        if record.output_sheet.startswith(("FSC_", "CC_")):
            records.append(record)
    return tuple(records)


def _cross_tab_column_codes(matrix: dict, column_label_map: dict) -> list[Any]:
    codes = set(column_label_map)
    for row_payload in matrix.values():
        if isinstance(row_payload, dict):
            codes.update(row_payload)
    return sorted(codes, key=_sort_key)


def _label_for_code(label_map: dict, code: Any) -> str:
    if code in label_map:
        return str(label_map[code])
    code_as_string = str(code)
    if code_as_string in label_map:
        return str(label_map[code_as_string])
    return code_as_string


def _nested_lookup(matrix: dict, row_code: Any, column_code: Any) -> Any:
    row_payload = matrix.get(row_code, matrix.get(str(row_code), {}))
    if not isinstance(row_payload, dict):
        return ""
    return row_payload.get(column_code, row_payload.get(str(column_code), 0))


def _unique_sheet_name(name: str, used_sheet_names: set[str]) -> str:
    base = _safe_sheet_name(name)
    candidate = base
    suffix = 1
    while candidate in used_sheet_names:
        suffix_text = f"_{suffix}"
        candidate = f"{base[:31 - len(suffix_text)]}{suffix_text}"
        suffix += 1
    used_sheet_names.add(candidate)
    return candidate


def _autofit(ws: Any, max_width: int = 60) -> None:
    """Set widths from tracked cell content, capped at max_width."""
    widths = getattr(ws, "_sie_widths", {})
    for col_index, width in widths.items():
        ws.set_column(col_index, col_index, min(width + 2, max_width))


def _denominator_description(result: SingleCutResult) -> str:
    policy = result.denominator_policy.value
    if policy == "VALID_RESPONSES":
        return f"Valid responses (n={result.valid_n})"
    if policy == "ALL_RESPONDENTS":
        return f"All respondents (n={result.valid_n})"
    if policy == "EXPOSED_TO_QUESTION":
        return f"Exposed to question (n={result.valid_n})"
    return policy


def _write_header_row(
    ws: Any, row_index: int, headers: list[str], formats: dict[str, Any]
) -> None:
    for column_index, header in enumerate(headers):
        _write(ws, row_index, column_index, header, formats["header"])


def _write_row(
    ws: Any, row_index: int, values: list[Any], cell_format: Any | None = None
) -> None:
    for column_index, value in enumerate(values):
        _write(ws, row_index, column_index, value, cell_format)


def _write(ws: Any, row: int, col: int, value: Any, cell_format: Any | None = None) -> None:
    if isinstance(value, float) and math.isnan(value):
        value = ""
    ws.write(row, col, value, cell_format)
    _track_cell(ws, row, col, value)


def _write_url(
    ws: Any,
    row: int,
    col: int,
    url: str,
    cell_format: Any,
    string: str,
) -> None:
    ws.write_url(row, col, url, cell_format, string=string)
    _track_cell(ws, row, col, string)


def _merge_range(
    ws: Any,
    first_row: int,
    first_col: int,
    last_row: int,
    last_col: int,
    value: Any,
    cell_format: Any | None = None,
) -> None:
    if hasattr(ws, "merge_range"):
        ws.merge_range(first_row, first_col, last_row, last_col, value, cell_format)
        _track_cell(ws, first_row, first_col, value)
        return
    _write(ws, first_row, first_col, value, cell_format)


def _track_cell(ws: Any, row: int, col: int, value: Any) -> None:
    if not hasattr(ws, "_sie_widths"):
        setattr(ws, "_sie_widths", {})
    if not hasattr(ws, "_sie_last_row"):
        setattr(ws, "_sie_last_row", 0)
    widths = getattr(ws, "_sie_widths")
    widths[col] = max(widths.get(col, 0), len(_cell_display(value)))
    setattr(ws, "_sie_last_row", max(getattr(ws, "_sie_last_row"), row))


def _find_last_used_row(ws: Any) -> int:
    return int(getattr(ws, "_sie_last_row", 0))


def _cell_display(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.isoformat()
    return str(value)


def _sort_key(value: Any) -> tuple[int, Any]:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return (0, value)
    return (1, str(value))


def _rate_sort_value(value: Any) -> float:
    try:
        rate = float(value)
    except (TypeError, ValueError):
        return float("-inf")
    return rate if not math.isnan(rate) else float("-inf")


def _quote_sheet_name(sheet_name: str) -> str:
    return sheet_name.replace("'", "''")


class _FallbackFormat:
    def __init__(self, properties: dict[str, Any], style_id: int) -> None:
        self.properties = properties
        self.style_id = style_id


class _FallbackWorkbook:
    def __init__(self, output_path: str) -> None:
        self.output_path = output_path
        self._worksheets: list[_FallbackWorksheet] = []

    def add_format(self, properties: dict[str, Any]) -> _FallbackFormat:
        return _FallbackFormat(properties, _fallback_style_id(properties))

    def add_worksheet(self, name: str) -> "_FallbackWorksheet":
        worksheet = _FallbackWorksheet(name)
        self._worksheets.append(worksheet)
        return worksheet

    def close(self) -> None:
        Path(self.output_path).parent.mkdir(parents=True, exist_ok=True)
        with ZipFile(self.output_path, "w", ZIP_DEFLATED) as archive:
            archive.writestr("[Content_Types].xml", self._content_types_xml())
            archive.writestr("_rels/.rels", self._root_rels_xml())
            archive.writestr("docProps/core.xml", self._core_xml())
            archive.writestr("docProps/app.xml", self._app_xml())
            archive.writestr("xl/workbook.xml", self._workbook_xml())
            archive.writestr("xl/_rels/workbook.xml.rels", self._workbook_rels_xml())
            archive.writestr("xl/styles.xml", self._styles_xml())
            for index, worksheet in enumerate(self._worksheets, start=1):
                archive.writestr(
                    f"xl/worksheets/sheet{index}.xml", worksheet.to_xml()
                )

    def _content_types_xml(self) -> str:
        sheet_overrides = "".join(
            f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.'
            'spreadsheetml.worksheet+xml"/>'
            for index in range(1, len(self._worksheets) + 1)
        )
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
            '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
            f"{sheet_overrides}</Types>"
        )

    def _root_rels_xml(self) -> str:
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
            "</Relationships>"
        )

    def _workbook_xml(self) -> str:
        sheets = "".join(
            f'<sheet name="{escape(worksheet.name)}" sheetId="{index}" r:id="rId{index}"/>'
            for index, worksheet in enumerate(self._worksheets, start=1)
        )
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f"<sheets>{sheets}</sheets></workbook>"
        )

    def _workbook_rels_xml(self) -> str:
        sheet_rels = "".join(
            f'<Relationship Id="rId{index}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
            for index in range(1, len(self._worksheets) + 1)
        )
        style_id = len(self._worksheets) + 1
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f"{sheet_rels}"
            f'<Relationship Id="rId{style_id}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
            'Target="styles.xml"/>'
            "</Relationships>"
        )

    def _styles_xml(self) -> str:
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<numFmts count="3">'
            '<numFmt numFmtId="164" formatCode="0.0%"/>'
            '<numFmt numFmtId="165" formatCode="#,##0"/>'
            '<numFmt numFmtId="166" formatCode="0.00"/>'
            "</numFmts>"
            '<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font>'
            '<font><b/><sz val="11"/><name val="Calibri"/></font>'
            '<font><u/><color rgb="000000FF"/><sz val="11"/><name val="Calibri"/></font></fonts>'
            '<fills count="3"><fill><patternFill patternType="none"/></fill>'
            '<fill><patternFill patternType="gray125"/></fill>'
            '<fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/><bgColor indexed="64"/></patternFill></fill></fills>'
            '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>'
            '<border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="7">'
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
            '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
            '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>'
            '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
            '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
            '<xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
            '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
            "</cellXfs>"
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            "</styleSheet>"
        )

    def _core_xml(self) -> str:
        timestamp = datetime.now(timezone.utc).isoformat()
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" '
            'xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            "<dc:creator>Survey Insight Engine</dc:creator>"
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{timestamp}</dcterms:created>'
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{timestamp}</dcterms:modified>'
            "</cp:coreProperties>"
        )

    def _app_xml(self) -> str:
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            "<Application>Survey Insight Engine</Application>"
            "</Properties>"
        )


class _FallbackWorksheet:
    def __init__(self, name: str) -> None:
        self.name = name
        self.cells: dict[tuple[int, int], tuple[Any, int]] = {}
        self.hyperlinks: dict[tuple[int, int], str] = {}
        self.widths: dict[int, float] = {}
        self.freeze: tuple[int, int] | None = None

    def write(self, row: int, col: int, value: Any, cell_format: _FallbackFormat | None = None) -> None:
        style_id = cell_format.style_id if cell_format is not None else 0
        self.cells[(row, col)] = (value, style_id)

    def write_url(
        self,
        row: int,
        col: int,
        url: str,
        cell_format: _FallbackFormat | None = None,
        string: str | None = None,
    ) -> None:
        display = string if string is not None else url
        self.write(row, col, display, cell_format)
        if url.startswith("internal:"):
            self.hyperlinks[(row, col)] = url.removeprefix("internal:")

    def freeze_panes(self, row: int, col: int) -> None:
        self.freeze = (row, col)

    def set_column(self, first_col: int, last_col: int, width: float) -> None:
        for col in range(first_col, last_col + 1):
            self.widths[col] = width

    def to_xml(self) -> str:
        rows: dict[int, list[tuple[int, Any, int]]] = defaultdict(list)
        for (row, col), (value, style_id) in self.cells.items():
            rows[row].append((col, value, style_id))
        dimension_xml = self._dimension_xml()
        row_xml = []
        for row in sorted(rows):
            cells = "".join(
                self._cell_xml(row, col, value, style_id)
                for col, value, style_id in sorted(rows[row])
            )
            row_xml.append(f'<row r="{row + 1}">{cells}</row>')
        cols_xml = ""
        if self.widths:
            cols_xml = "<cols>" + "".join(
                f'<col min="{col + 1}" max="{col + 1}" width="{width}" customWidth="1"/>'
                for col, width in sorted(self.widths.items())
            ) + "</cols>"
        hyperlinks_xml = ""
        if self.hyperlinks:
            links = "".join(
                f'<hyperlink ref="{_cell_ref(row, col)}" location="{escape(location)}" display="{escape(str(self.cells[(row, col)][0]))}"/>'
                for (row, col), location in sorted(self.hyperlinks.items())
            )
            hyperlinks_xml = f"<hyperlinks>{links}</hyperlinks>"
        sheet_views = self._sheet_views_xml()
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f"{dimension_xml}{sheet_views}{cols_xml}<sheetData>{''.join(row_xml)}</sheetData>{hyperlinks_xml}"
            "</worksheet>"
        )

    def _dimension_xml(self) -> str:
        if not self.cells:
            return '<dimension ref="A1"/>'
        max_row = max(row for row, _ in self.cells)
        max_col = max(col for _, col in self.cells)
        return f'<dimension ref="A1:{_cell_ref(max_row, max_col)}"/>'

    def _sheet_views_xml(self) -> str:
        if self.freeze is None:
            return '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        row, col = self.freeze
        top_left = _cell_ref(row, col)
        return (
            '<sheetViews><sheetView workbookViewId="0">'
            f'<pane ySplit="{row}" xSplit="{col}" topLeftCell="{top_left}" '
            'activePane="bottomRight" state="frozen"/>'
            "</sheetView></sheetViews>"
        )

    def _cell_xml(self, row: int, col: int, value: Any, style_id: int) -> str:
        cell_ref = _cell_ref(row, col)
        style_attr = f' s="{style_id}"' if style_id else ""
        if value is None or value == "":
            return f'<c r="{cell_ref}"{style_attr}/>'
        if isinstance(value, bool):
            return f'<c r="{cell_ref}" t="b"{style_attr}><v>{1 if value else 0}</v></c>'
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return f'<c r="{cell_ref}"{style_attr}><v>{value}</v></c>'
        text = escape(str(value))
        preserve = ' xml:space="preserve"' if str(value).strip() != str(value) else ""
        return f'<c r="{cell_ref}" t="inlineStr"{style_attr}><is><t{preserve}>{text}</t></is></c>'


def _fallback_style_id(properties: dict[str, Any]) -> int:
    if properties.get("num_format") == "0.0%":
        return 3
    if properties.get("num_format") == "#,##0":
        return 4
    if properties.get("num_format") == "0.00":
        return 5
    if properties.get("underline"):
        return 6
    if properties.get("bold") and properties.get("bg_color"):
        return 2
    if properties.get("bold"):
        return 1
    if properties.get("italic"):
        return 0
    return 0


def _cell_ref(row: int, col: int) -> str:
    return f"{_column_name(col)}{row + 1}"


def _column_name(col: int) -> str:
    name = ""
    col += 1
    while col:
        col, remainder = divmod(col - 1, 26)
        name = chr(65 + remainder) + name
    return name
