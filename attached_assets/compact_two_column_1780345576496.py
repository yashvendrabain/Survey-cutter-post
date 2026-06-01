"""Compact two-column data-map adapter."""

from __future__ import annotations

import re
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


DATAMAP_KEYWORDS = (
    "datamap",
    "data map",
    "codebook",
    "schema",
    "dictionary",
    "questions",
    "questionnaire",
    "variables",
    "metadata",
)


class CompactTwoColumnAdapter(DataMapAdapter):
    """Parser adapter for compact QID/option two-column data maps."""

    name = "compact_two_column"

    def detect(self, workbook: Any, raw_df: Any | None = None) -> AdapterDetectionResult:
        del raw_df
        sheet_name = _find_datamap_sheet(workbook)
        if sheet_name is None:
            return AdapterDetectionResult(0.0, "no data map sheet found by name")

        worksheet = workbook[sheet_name]
        detected_format = _detect_datamap_format(worksheet)
        if detected_format == "compact_two_column":
            return AdapterDetectionResult(
                0.9,
                f"column A/B compact question-option pattern found in {sheet_name!r}",
            )
        return AdapterDetectionResult(0.0, f"sheet {sheet_name!r} not compact format")

    def parse(self, workbook: Any, raw_df: Any | None = None) -> dict:
        del raw_df
        sheet_name = _find_datamap_sheet(workbook)
        if sheet_name is None:
            raise ValueError("no data map sheet found by name")
        worksheet = workbook[sheet_name]
        questions = _parse_compact_datamap(worksheet)
        merged_questions = _merge_per_row_children(questions)
        merged_questions = _validate_conditional_on(merged_questions)
        merged_questions = _attach_numeric_label_metadata(merged_questions)
        return {
            "questions": merged_questions,
            "source_path": str(getattr(workbook, "_survey_source_path", "<workbook>")),
            "sheet_name": sheet_name,
            "total_rows_in_sheet": int(worksheet.max_row or 0),
            "parser_warnings": [],
        }


def _find_datamap_sheet(workbook: Any) -> str | None:
    """Locate the data map sheet by normalized name match."""

    sheet_names = list(getattr(workbook, "sheetnames", []))
    for name in sheet_names:
        normalized = re.sub(r"[\s_]+", "", str(name).strip().lower())
        if any(re.sub(r"[\s_]+", "", keyword) in normalized for keyword in DATAMAP_KEYWORDS):
            return name
    if DATAMAP_SHEET_NAME in sheet_names:
        return DATAMAP_SHEET_NAME
    return None
