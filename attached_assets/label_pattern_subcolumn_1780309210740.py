"""Utility for raw columns named like ``Q3: option label``."""

from __future__ import annotations

import re

from src.datamap_parser import ParsedQuestion


_TRAILING_PUNCTUATION = ".,;:!?- "


def apply_label_pattern_matching(
    question: ParsedQuestion,
    raw_columns: tuple[str, ...],
) -> ParsedQuestion:
    """Infer binary sub-columns from raw headers like ``Q3: Option label``."""

    if question["sub_columns"] or not question["options"] or question.get("children"):
        return question

    inferred_sub_columns, warnings = infer_label_pattern_sub_columns(
        question,
        raw_columns,
    )
    if len(inferred_sub_columns) < 2:
        if warnings:
            flagged = dict(question)
            flagged["warnings"] = list(question.get("warnings", [])) + warnings
            return flagged  # type: ignore[return-value]
        return question

    promoted = dict(question)
    promoted["sub_columns"] = inferred_sub_columns
    promoted["options"] = []
    promoted["type_hint"] = "values_range"
    promoted["value_range"] = (0, 1)
    promoted["warnings"] = list(question.get("warnings", [])) + warnings
    return promoted  # type: ignore[return-value]


def infer_label_pattern_sub_columns(
    question: ParsedQuestion,
    raw_columns: tuple[str, ...],
) -> tuple[list[tuple[str, str]], list[str]]:
    """Match ``QID: label`` / ``QID - label`` raw columns to option labels."""

    canonical_id = question["canonical_id"]
    label_pattern = re.compile(
        rf"^\s*{re.escape(canonical_id)}\s*(?::|[-\u2013\u2014])\s*(?P<label>.+?)\s*$",
        re.IGNORECASE,
    )
    options = [(str(code), str(label).strip()) for code, label in question["options"]]
    inferred: list[tuple[str, str]] = []
    warnings: list[str] = []
    used_labels: set[str] = set()

    for raw_column in raw_columns:
        match = label_pattern.match(str(raw_column))
        if match is None:
            continue

        column_label = match.group("label").strip()
        option_label, match_kind = _match_label_pattern_option(column_label, options)
        if option_label is None:
            warnings.append(
                f"Could not match column {raw_column!r} to any option in {canonical_id}."
            )
            continue
        if option_label in used_labels:
            warnings.append(
                f"Duplicate label-pattern match for {canonical_id} option {option_label!r}; "
                f"ignored column {raw_column!r}."
            )
            continue

        if match_kind == "substring":
            warnings.append(
                f"Accepted substring match for column {raw_column!r} "
                f"to option {option_label!r} in {canonical_id}."
            )
        inferred.append((str(raw_column), option_label))
        used_labels.add(option_label)

    return inferred, warnings


def label_pattern_warnings(question: ParsedQuestion) -> list[str]:
    prefixes = (
        "Could not match column",
        "Accepted substring match",
        "Duplicate label-pattern match",
    )
    return [
        warning
        for warning in question.get("warnings", []) or []
        if any(str(warning).startswith(prefix) for prefix in prefixes)
    ]


def _match_label_pattern_option(
    column_label: str,
    options: list[tuple[str, str]],
) -> tuple[str | None, str | None]:
    stripped = column_label.strip()

    exact_matches = [label for _code, label in options if label.strip() == stripped]
    if len(exact_matches) == 1:
        return exact_matches[0], "exact"

    casefolded = stripped.casefold()
    case_matches = [
        label for _code, label in options if label.strip().casefold() == casefolded
    ]
    if len(case_matches) == 1:
        return case_matches[0], "case_insensitive"

    normalized = _normalise_label_pattern_text(stripped)
    normalized_matches = [
        label
        for _code, label in options
        if _normalise_label_pattern_text(label) == normalized
    ]
    if len(normalized_matches) == 1:
        return normalized_matches[0], "normalized"

    substring_matches = []
    for _code, label in options:
        option_norm = _normalise_label_pattern_text(label)
        if (
            normalized
            and option_norm
            and (normalized in option_norm or option_norm in normalized)
        ):
            substring_matches.append(label)
    if len(substring_matches) == 1:
        return substring_matches[0], "substring"

    return None, None


def _normalise_label_pattern_text(value: str) -> str:
    without_parentheticals = re.sub(r"\([^)]*\)", "", str(value))
    stripped = without_parentheticals.rstrip(_TRAILING_PUNCTUATION)
    return re.sub(r"\s+", " ", stripped).strip().casefold()
