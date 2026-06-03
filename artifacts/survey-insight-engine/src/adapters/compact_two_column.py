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


RAW_COL_MATCH_THRESHOLD = 0.30
QID_SAMPLE_LIMIT = 20
_QID_PREFIX_SUFFIXES = (":", "r", "_", ".", "-", " ")


DATAMAP_KEYWORDS = (
    "datamap",
    "codebook",
    "schema",
)


class CompactTwoColumnAdapter(DataMapAdapter):
    """Parser adapter for compact QID/option two-column data maps."""

    name = "compact_two_column"

    def detect(self, workbook: Any, raw_df: Any | None = None) -> AdapterDetectionResult:
        sheet_name = _find_datamap_sheet(workbook)
        if sheet_name is None:
            return AdapterDetectionResult(0.0, "no data map sheet found by name")

        worksheet = workbook[sheet_name]
        detected_format = _detect_datamap_format(worksheet)
        if detected_format != "compact_two_column":
            return AdapterDetectionResult(0.0, f"sheet {sheet_name!r} not compact format")

        # Raw-data cross-check: a compact codebook only justifies high confidence
        # if the QIDs it declares actually appear in the raw data columns. Without
        # this check the adapter over-claims on workbooks whose codebook happens
        # to be in compact form but whose raw data uses an alien column-naming
        # convention (causing the wizard path to be missed).
        if raw_df is not None:
            sampled_qids = _sample_qids_from_compact_sheet(
                worksheet, limit=QID_SAMPLE_LIMIT
            )
            if sampled_qids:
                ratio = _qid_match_ratio(
                    sampled_qids,
                    [str(column) for column in getattr(raw_df, "columns", [])],
                )
                if ratio < RAW_COL_MATCH_THRESHOLD:
                    return AdapterDetectionResult(
                        0.0,
                        f"compact codebook in {sheet_name!r} but only "
                        f"{ratio:.0%} of QIDs match raw data columns; "
                        "treating as unknown survey format",
                    )

        return AdapterDetectionResult(
            0.9,
            f"column A/B compact question-option pattern found in {sheet_name!r}",
        )

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


def _sample_qids_from_compact_sheet(worksheet: Any, *, limit: int) -> list[str]:
    """Sample question-header rows from a compact codebook.

    A header row has a non-numeric label in column A and a non-empty
    question text in column B. Option rows (numeric col A) are skipped.
    """

    qids: list[str] = []
    for row in worksheet.iter_rows(min_col=1, max_col=2, values_only=True):
        if len(qids) >= limit:
            break
        col_a = row[0] if len(row) > 0 else None
        col_b = row[1] if len(row) > 1 else None
        if col_a is None or col_b is None:
            continue
        text_a = str(col_a).strip()
        text_b = str(col_b).strip()
        if not text_a or not text_b:
            continue
        if re.fullmatch(r"-?\d+(?:\.\d+)?", text_a):
            continue
        qids.append(text_a)
    return qids


def _qid_match_ratio(qids: list[str], raw_columns: list[str]) -> float:
    """Fraction of QIDs that appear in raw columns exactly or as a prefix.

    A match is either:
      - exact case-insensitive equality, or
      - the QID followed by any of _QID_PREFIX_SUFFIXES (":", "r", "_",
        ".", "-", " ") as the next character. This covers known formats:
        "Q1" exact, "Q6: Field sales" (colon-prefix multi-select),
        "Q1r1" (BCN sub-column), "Q1_a", "Q1.1", "Q1 something".

    The suffix requirement prevents false positives like QID "q1" matching
    raw column "q10".
    """

    if not qids:
        return 0.0
    raw_lc = [column.casefold() for column in raw_columns]
    raw_set = set(raw_lc)
    hits = 0
    for qid in qids:
        q = qid.casefold()
        if q in raw_set:
            hits += 1
            continue
        if any(
            column.startswith(q + suffix)
            for suffix in _QID_PREFIX_SUFFIXES
            for column in raw_lc
        ):
            hits += 1
    return hits / len(qids)
