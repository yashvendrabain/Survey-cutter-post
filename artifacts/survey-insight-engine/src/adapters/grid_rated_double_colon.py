"""Utility for rated grids named like ``Q14: dim1 :: dim2``."""

from __future__ import annotations

from collections import Counter
import re

from src.datamap_parser import ParsedQuestion


_WARNING_PREFIX = "Grid rated double-colon:"
_LABEL_PATTERN_WARNING_PREFIXES = (
    "Could not match column",
    "Accepted substring match",
    "Duplicate label-pattern match",
)


def apply_grid_rated_double_colon_matching(
    question: ParsedQuestion,
    raw_columns: tuple[str, ...],
) -> ParsedQuestion:
    """Infer rated-grid cells from raw headers like ``Q14: Planned :: Revenue``."""

    if question["sub_columns"] or question.get("children"):
        return question
    if len(question["options"]) > 1:
        return question

    cells, warnings = _parse_double_colon_cells(question, raw_columns)
    if len(cells) < 2:
        if warnings:
            flagged = dict(question)
            flagged["warnings"] = list(question.get("warnings", [])) + warnings
            return flagged  # type: ignore[return-value]
        return question

    warnings.extend(_grid_structure_warnings(question["canonical_id"], cells))

    promoted = dict(question)
    promoted["sub_columns"] = [
        (raw_column, f"{dim1_label} :: {dim2_label}")
        for raw_column, dim1_label, dim2_label in cells
    ]
    promoted["warnings"] = _without_stale_label_pattern_warnings(question) + warnings
    return promoted  # type: ignore[return-value]


def grid_rated_double_colon_warnings(question: ParsedQuestion) -> list[str]:
    return [
        warning
        for warning in question.get("warnings", []) or []
        if str(warning).startswith(_WARNING_PREFIX)
    ]


def _parse_double_colon_cells(
    question: ParsedQuestion,
    raw_columns: tuple[str, ...],
) -> tuple[list[tuple[str, str, str]], list[str]]:
    canonical_id = question["canonical_id"]
    cell_pattern = re.compile(
        rf"^\s*{re.escape(canonical_id)}\s*:\s*"
        r"(?P<dim1>.+?)\s*::\s*(?P<dim2>.+?)\s*$",
        re.IGNORECASE,
    )
    cells: list[tuple[str, str, str]] = []
    warnings: list[str] = []

    for raw_column in raw_columns:
        raw_column_text = str(raw_column)
        match = cell_pattern.match(raw_column_text)
        if match is None:
            continue

        dim1_label = re.sub(r"\s+", " ", match.group("dim1")).strip()
        dim2_label = re.sub(r"\s+", " ", match.group("dim2")).strip()
        if not dim1_label or not dim2_label:
            warnings.append(
                f"{_WARNING_PREFIX} unmatched column {raw_column_text!r} "
                f"for {canonical_id}."
            )
            continue
        cells.append((raw_column_text, dim1_label, dim2_label))

    return cells, warnings


def _grid_structure_warnings(
    canonical_id: str,
    cells: list[tuple[str, str, str]],
) -> list[str]:
    warnings: list[str] = []
    dim1_values = _ordered_unique(dim1_label for _raw, dim1_label, _dim2 in cells)
    dim2_values = _ordered_unique(dim2_label for _raw, _dim1, dim2_label in cells)
    observed_pairs = [(dim1_label, dim2_label) for _raw, dim1_label, dim2_label in cells]

    duplicate_pairs = [
        pair for pair, count in Counter(observed_pairs).items() if count > 1
    ]
    for dim1_label, dim2_label in duplicate_pairs:
        warnings.append(
            f"{_WARNING_PREFIX} duplicate cell {dim1_label!r} :: {dim2_label!r} "
            f"for {canonical_id}."
        )

    if len(dim1_values) < 2 or len(dim2_values) < 2:
        warnings.append(
            f"{_WARNING_PREFIX} non-grid pattern for {canonical_id}; "
            f"found {len(dim1_values)} first-dimension values and "
            f"{len(dim2_values)} second-dimension values."
        )
        return warnings

    observed_pair_set = set(observed_pairs)
    missing_pairs = [
        (dim1_label, dim2_label)
        for dim1_label in dim1_values
        for dim2_label in dim2_values
        if (dim1_label, dim2_label) not in observed_pair_set
    ]
    if missing_pairs:
        preview = ", ".join(
            f"{dim1_label!r} :: {dim2_label!r}"
            for dim1_label, dim2_label in missing_pairs[:3]
        )
        suffix = "..." if len(missing_pairs) > 3 else ""
        warnings.append(
            f"{_WARNING_PREFIX} partial grid for {canonical_id}; "
            f"missing {len(missing_pairs)} cell(s): {preview}{suffix}."
        )

    return warnings


def _ordered_unique(values: object) -> list[str]:
    ordered: list[str] = []
    seen: set[str] = set()
    for value in values:  # type: ignore[assignment]
        text = str(value)
        key = text.casefold()
        if key in seen:
            continue
        ordered.append(text)
        seen.add(key)
    return ordered


def _without_stale_label_pattern_warnings(question: ParsedQuestion) -> list[str]:
    return [
        warning
        for warning in question.get("warnings", []) or []
        if not any(
            str(warning).startswith(prefix)
            for prefix in _LABEL_PATTERN_WARNING_PREFIXES
        )
    ]
