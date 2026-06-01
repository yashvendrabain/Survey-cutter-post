"""Compact two-column data-map adapter."""

from __future__ import annotations

from typing import Any

from src.adapters.base import AdapterDetectionResult, DataMapAdapter
from src.datamap_parser import (
    DATAMAP_SHEET_NAME,
    _attach_numeric_label_metadata,
    _detect_datamap_format,
    _merge_per_row_children,
    _parse_compact_datamap,
    _validate_conditional_on,
)


class CompactTwoColumnAdapter(DataMapAdapter):
    """Parser adapter for compact QID/option two-column data maps."""

    name = "compact_two_column"

    def detect(self, workbook: Any, raw_df: Any | None = None) -> AdapterDetectionResult:
        del raw_df
        if DATAMAP_SHEET_NAME not in getattr(workbook, "sheetnames", []):
            return AdapterDetectionResult(0.0, "Sheet1 not found")

        worksheet = workbook[DATAMAP_SHEET_NAME]
        detected_format = _detect_datamap_format(worksheet)
        if detected_format == "compact_two_column":
            return AdapterDetectionResult(
                0.9,
                "column A/B compact question-option pattern found",
            )
        return AdapterDetectionResult(0.0, "no compact two-column signals")

    def parse(self, workbook: Any, raw_df: Any | None = None) -> dict:
        del raw_df
        worksheet = workbook[DATAMAP_SHEET_NAME]
        questions = _parse_compact_datamap(worksheet)
        merged_questions = _merge_per_row_children(questions)
        merged_questions = _validate_conditional_on(merged_questions)
        merged_questions = _attach_numeric_label_metadata(merged_questions)
        return {
            "questions": merged_questions,
            "source_path": str(getattr(workbook, "_survey_source_path", "<workbook>")),
            "sheet_name": DATAMAP_SHEET_NAME,
            "total_rows_in_sheet": int(worksheet.max_row or 0),
            "parser_warnings": [],
        }
