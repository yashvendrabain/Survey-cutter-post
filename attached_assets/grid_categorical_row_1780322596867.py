"""Utility for grid rows named like ``Q27: row label``."""

from __future__ import annotations

import re

from src.adapters.label_pattern_subcolumn import infer_label_pattern_sub_columns
from src.datamap_parser import ParsedQuestion


_WARNING_PREFIX = "Grid categorical row:"
_LABEL_PATTERN_WARNING_PREFIXES = (
    "Could not match column",
    "Accepted substring match",
    "Duplicate label-pattern match",
)


def apply_grid_categorical_row_matching(
    question: ParsedQuestion,
    raw_columns: tuple[str, ...],
) -> ParsedQuestion:
    """Infer grid rows from raw headers like ``Q27: Field sales``."""

    if question["sub_columns"] or question.get("children"):
        return question
    if not 2 <= len(question["options"]) <= 10:
        return question
    if _has_bcn_style_sub_column(question["canonical_id"], raw_columns):
        return question

    inferred_sub_columns, warnings = _infer_grid_row_sub_columns(question, raw_columns)
    if len(inferred_sub_columns) < 2:
        if warnings:
            flagged = dict(question)
            flagged["warnings"] = list(question.get("warnings", [])) + warnings
            return flagged  # type: ignore[return-value]
        return question
    if _columns_match_option_labels(question, tuple(column for column, _label in inferred_sub_columns)):
        return question

    promoted = dict(question)
    promoted["sub_columns"] = inferred_sub_columns
    promoted["type_hint"] = "values_range"
    promoted["warnings"] = _without_stale_label_pattern_warnings(question) + warnings
    return promoted  # type: ignore[return-value]


def grid_categorical_row_warnings(question: ParsedQuestion) -> list[str]:
    return [
        warning
        for warning in question.get("warnings", []) or []
        if str(warning).startswith(_WARNING_PREFIX)
    ]


def _infer_grid_row_sub_columns(
    question: ParsedQuestion,
    raw_columns: tuple[str, ...],
) -> tuple[list[tuple[str, str]], list[str]]:
    canonical_id = question["canonical_id"]
    row_pattern = re.compile(
        rf"^\s*{re.escape(canonical_id)}\s*(?::|[-\u2013\u2014])\s*(?P<label>.*?)\s*$",
        re.IGNORECASE,
    )
    inferred: list[tuple[str, str]] = []
    warnings: list[str] = []
    used_labels: set[str] = set()

    for raw_column in raw_columns:
        raw_column_text = str(raw_column)
        match = row_pattern.match(raw_column_text)
        if match is None:
            if _looks_like_unmatched_question_column(canonical_id, raw_column_text):
                warnings.append(
                    f"{_WARNING_PREFIX} unmatched column {raw_column_text!r} "
                    f"for {canonical_id}."
                )
            continue

        row_label = re.sub(r"\s+", " ", match.group("label")).strip()
        if not row_label or " :: " in row_label:
            warnings.append(
                f"{_WARNING_PREFIX} unmatched column {raw_column_text!r} "
                f"for {canonical_id}."
            )
            continue
        label_key = row_label.casefold()
        if label_key in used_labels:
            warnings.append(
                f"{_WARNING_PREFIX} duplicate row label {row_label!r} "
                f"for {canonical_id}; ignored column {raw_column_text!r}."
            )
            continue

        inferred.append((raw_column_text, row_label))
        used_labels.add(label_key)

    return inferred, warnings


def _columns_match_option_labels(
    question: ParsedQuestion,
    raw_columns: tuple[str, ...],
) -> bool:
    matched_sub_columns, _warnings = infer_label_pattern_sub_columns(
        question,
        raw_columns,
    )
    return bool(matched_sub_columns)


def _has_bcn_style_sub_column(canonical_id: str, raw_columns: tuple[str, ...]) -> bool:
    pattern = re.compile(rf"^\s*{re.escape(canonical_id)}r\d+", re.IGNORECASE)
    return any(pattern.match(str(raw_column)) for raw_column in raw_columns)


def _looks_like_unmatched_question_column(canonical_id: str, raw_column: str) -> bool:
    text = str(raw_column).strip()
    if re.match(rf"^{re.escape(canonical_id)}r\d+", text, re.IGNORECASE):
        return False
    return bool(
        re.match(
            rf"^{re.escape(canonical_id)}\s*(?::|[-\u2013\u2014])",
            text,
            re.IGNORECASE,
        )
    )


def _without_stale_label_pattern_warnings(question: ParsedQuestion) -> list[str]:
    return [
        warning
        for warning in question.get("warnings", []) or []
        if not any(
            str(warning).startswith(prefix)
            for prefix in _LABEL_PATTERN_WARNING_PREFIXES
        )
    ]
