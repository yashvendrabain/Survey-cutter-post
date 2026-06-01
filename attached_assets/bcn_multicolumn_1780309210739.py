"""BCN/Growth Agenda multi-column data-map adapter."""

from __future__ import annotations

import re
from typing import Any

from src.adapters.base import AdapterDetectionResult, DataMapAdapter
from src.datamap_parser import (
    DATAMAP_SHEET_NAME,
    _State,
    _attach_numeric_label_metadata,
    _capture_option_row,
    _capture_type_hint,
    _cell,
    _detect_datamap_format,
    _finalise_block,
    _header_match,
    _is_blank_row,
    _merge_per_row_children,
    _normalise_row,
    _row_preview,
    _start_block,
    _validate_conditional_on,
)


BCN_SUB_COLUMN_RE = re.compile(
    r"^Q(\d+)([a-zA-Z_])(\d+)(?:c(\d+))?(?:oe)?$",
    re.IGNORECASE,
)


class BcnMulticolumnAdapter(DataMapAdapter):
    """Parser adapter for legacy BCN-style data map workbooks."""

    name = "bcn_multicolumn"

    def detect(self, workbook: Any, raw_df: Any | None = None) -> AdapterDetectionResult:
        if DATAMAP_SHEET_NAME not in getattr(workbook, "sheetnames", []):
            return AdapterDetectionResult(0.0, "Sheet1 not found")

        worksheet = workbook[DATAMAP_SHEET_NAME]
        confidence = 0.0
        reasons: list[str] = []

        detected_format = _detect_datamap_format(worksheet)
        if detected_format == "bcn_multicolumn":
            confidence = 0.7
            reasons.append("BCN-style header/type-hint signals found")

        sub_column_count = _raw_sub_column_count(raw_df)
        if sub_column_count >= 10:
            confidence += 0.5
            reasons.append(f"{sub_column_count} raw columns match BCN sub-column pattern")
        if sub_column_count >= 50:
            confidence += 0.2
            reasons.append("50+ raw sub-column pattern matches")

        confidence = min(confidence, 1.0)
        return AdapterDetectionResult(
            confidence,
            "; ".join(reasons) if reasons else "no BCN multi-column signals",
        )

    def parse(self, workbook: Any, raw_df: Any | None = None) -> dict:
        del raw_df
        return _parse_bcn_multicolumn_workbook(workbook)


def _raw_sub_column_count(raw_df: Any | None) -> int:
    if raw_df is None:
        return 0
    columns = getattr(raw_df, "columns", raw_df)
    try:
        return sum(1 for column in columns if BCN_SUB_COLUMN_RE.match(str(column)))
    except TypeError:
        return 0


def _parse_bcn_multicolumn_workbook(workbook: Any) -> dict:
    worksheet = workbook[DATAMAP_SHEET_NAME]
    questions = []
    parser_warnings: list[str] = []
    state = _State.BETWEEN_BLOCKS
    current_block = None
    total_rows = 0

    for row_number, raw_row in enumerate(worksheet.iter_rows(values_only=True), start=1):
        total_rows = row_number
        row = _normalise_row(raw_row)
        col_a = _cell(row, 0)
        col_b = _cell(row, 1)
        col_c = _cell(row, 2)
        header_match = _header_match(col_a)
        is_blank = _is_blank_row(col_a, col_b)
        is_option_style = col_a is None and col_b is not None

        if state is _State.BETWEEN_BLOCKS:
            if is_blank:
                continue
            if header_match:
                current_block = _start_block(col_a, header_match, row_number)
                state = _State.IN_HEADER
                continue
            parser_warnings.append(
                f"orphan row at row {row_number}: {_row_preview(row)!r}"
            )
            continue

        if current_block is None:
            raise RuntimeError("parser entered a block state without a block")

        if state is _State.IN_HEADER:
            if is_blank:
                current_block["warnings"].append(
                    "header followed by blank row, no type hint"
                )
                questions.append(_finalise_block(current_block))
                current_block = None
                state = _State.BETWEEN_BLOCKS
                continue
            if header_match:
                parser_warnings.append(
                    f"header found mid-block at row {row_number} "
                    "â€” finalised previous block early"
                )
                questions.append(_finalise_block(current_block))
                current_block = _start_block(col_a, header_match, row_number)
                state = _State.IN_HEADER
                continue
            if is_option_style:
                current_block["warnings"].append(
                    f"option row encountered before type hint at row {row_number}"
                )
                continue

            _capture_type_hint(current_block, col_a, row_number)
            state = _State.IN_TYPE_HINT
            continue

        if state is _State.IN_TYPE_HINT:
            if is_blank:
                questions.append(_finalise_block(current_block))
                current_block = None
                state = _State.BETWEEN_BLOCKS
                continue
            if header_match:
                parser_warnings.append(
                    f"header found mid-block at row {row_number} "
                    "â€” finalised previous block early"
                )
                questions.append(_finalise_block(current_block))
                current_block = _start_block(col_a, header_match, row_number)
                state = _State.IN_HEADER
                continue
            if is_option_style:
                _capture_option_row(current_block, col_b, col_c, row_number)
                state = _State.IN_OPTIONS
                continue

            current_block["warnings"].append(
                f"unrecognised row after type hint at row {row_number}: "
                f"{_row_preview(row)!r}"
            )
            continue

        if state is _State.IN_OPTIONS:
            if is_blank:
                questions.append(_finalise_block(current_block))
                current_block = None
                state = _State.BETWEEN_BLOCKS
                continue
            if header_match:
                parser_warnings.append(
                    f"header found mid-block at row {row_number} "
                    "â€” finalised previous block early"
                )
                questions.append(_finalise_block(current_block))
                current_block = _start_block(col_a, header_match, row_number)
                state = _State.IN_HEADER
                continue
            if is_option_style:
                _capture_option_row(current_block, col_b, col_c, row_number)
                continue

            current_block["warnings"].append(
                f"unrecognised option row at row {row_number}: {_row_preview(row)!r}"
            )

    if current_block is not None:
        questions.append(_finalise_block(current_block))

    merged_questions = _merge_per_row_children(questions)
    merged_questions = _validate_conditional_on(merged_questions)
    merged_questions = _attach_numeric_label_metadata(merged_questions)
    return {
        "questions": merged_questions,
        "source_path": str(getattr(workbook, "_survey_source_path", "<workbook>")),
        "sheet_name": DATAMAP_SHEET_NAME,
        "total_rows_in_sheet": total_rows,
        "parser_warnings": parser_warnings,
    }
