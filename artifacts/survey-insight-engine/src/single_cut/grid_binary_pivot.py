"""Pivoted binary grid single-cut calculator."""

from __future__ import annotations

from collections import defaultdict
import re
from numbers import Real
from typing import Any

import pandas as pd

from src.calc_primitives import _make_audit_record
from src.calculation_log import CalculationLog
from src.models import (
    DenominatorPolicy,
    GridBinaryPivotResult,
    GridBinaryPivotRow,
    QuestionSpec,
    QuestionType,
)


REJECTION_PREFIXES = (
    "NO TO: ",
    "NOT: ",
    "No to: ",
    "NOT SELECTED: ",
    "NOT SELECTED - ",
    "NO - ",
    "Not selected: ",
)


def compute_grid_binary_pivot(
    question_spec: QuestionSpec,
    df: pd.DataFrame,
    log: CalculationLog,
    filter_mask: pd.Series | None = None,
    filter_expr: str | None = None,
) -> GridBinaryPivotResult:
    """Compute count/% cells for a binary row x c-column grid."""

    if question_spec.question_type is not QuestionType.GRID_BINARY_SELECT:
        raise ValueError(
            f"unsupported question_type for binary grid: {question_spec.question_type}"
        )

    working_df = df
    if filter_mask is not None:
        working_df = working_df[filter_mask]
        filter_expr = filter_expr or "<unnamed filter>"
    else:
        filter_expr = None

    grouped_columns, group_order = _group_grid_columns(question_spec, working_df)
    if not grouped_columns:
        raise ValueError("binary grid raw columns not found in dataframe")

    column_headers = [_grid_group_label(question_spec, group_key) for group_key in group_order]
    all_columns = [column for columns in grouped_columns.values() for column in columns.values()]
    coerced = pd.DataFrame(
        {
            column: working_df[column].map(
                lambda value, label=_grid_group_label(question_spec, _row_and_group_ids(column)[1] or "1"): _coerce_binary_value(value, label)
            )
            for column in all_columns
        },
        index=working_df.index,
    )
    total_respondents = int(coerced.notna().any(axis=1).sum())
    filtered_n = int(len(working_df))
    missing_n = int(filtered_n - total_respondents)

    rows: list[GridBinaryPivotRow] = []
    total_responses = 0
    row_labels = question_spec.grid_row_labels or {}
    for row_id, columns_by_group in grouped_columns.items():
        row_frame = coerced[list(columns_by_group.values())]
        row_denominator = int(row_frame.notna().any(axis=1).sum())
        counts: list[int] = []
        pcts: list[float] = []
        for group_key in group_order:
            column = columns_by_group.get(group_key)
            count = int((coerced[column] == 1).sum()) if column is not None else 0
            pct = float(count / row_denominator) if row_denominator else 0.0
            counts.append(count)
            pcts.append(pct)
            total_responses += count
            if column is not None:
                audit = _make_audit_record(
                    metric_name=f"{row_id}_{group_key}_selection_rate",
                    question_id=question_spec.canonical_id,
                    source_columns=(column,),
                    filter_expr=filter_expr,
                    numerator=count,
                    denominator=row_denominator,
                    formula="count(selected) / count(row answered)",
                    value_raw=pct,
                    valid_n=row_denominator,
                    missing_n=filtered_n - row_denominator,
                    output_sheet=f"SC_{question_spec.canonical_id}",
                )
                log.record(audit)
        first_column = next(iter(columns_by_group.values()))
        rows.append(
            GridBinaryPivotRow(
                row_id=row_id,
                row_label=_base_row_label(row_labels.get(first_column, row_id)),
                counts_per_column=counts,
                pcts_per_column=pcts,
            )
        )

    return GridBinaryPivotResult(
        question_id=question_spec.canonical_id,
        question_type=QuestionType.GRID_BINARY_SELECT,
        valid_n=total_respondents,
        missing_n=missing_n,
        denominator_policy=DenominatorPolicy.VALID_RESPONSES,
        question_text=question_spec.question_text,
        column_headers=column_headers,
        rows=rows,
        total_respondents=total_respondents,
        total_responses=total_responses,
    )


def _group_grid_columns(
    question_spec: QuestionSpec,
    df: pd.DataFrame,
) -> tuple[dict[str, dict[str, str]], list[str]]:
    grouped: dict[str, dict[str, str]] = defaultdict(dict)
    group_order: list[str] = []
    for column in question_spec.raw_columns:
        if column not in df.columns:
            continue
        row_id, group_key = _row_and_group_ids(column)
        if group_key is None:
            group_key = "1"
        grouped[row_id][group_key] = column
        if group_key not in group_order:
            group_order.append(group_key)
    group_order.sort(key=lambda key: (int(key) if str(key).isdigit() else 10**9, key))
    return dict(sorted(grouped.items(), key=lambda item: _row_sort_key(item[0]))), group_order


def _row_and_group_ids(column: str) -> tuple[str, str | None]:
    match = re.match(r"^(?P<row>.+?r\d+)c(?P<group>\d+)$", str(column))
    if match is None:
        return str(column), None
    return match.group("row"), match.group("group")


def _row_sort_key(row_id: str) -> tuple[int, str]:
    match = re.search(r"r(\d+)$", row_id)
    if match is None:
        return (10**9, row_id)
    return (int(match.group(1)), row_id)


def _coerce_binary_value(value: Any, selected_label: str) -> int | None:
    if pd.isna(value):
        return None
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, Real):
        numeric = float(value)
        if numeric == 1.0:
            return 1
        if numeric == 0.0:
            return 0
    text = str(value).strip()
    if not text:
        return None
    lowered = text.casefold()
    if lowered in {"1", "checked", "selected", "yes", "true"}:
        return 1
    if lowered in {"0", "unchecked", "not selected", "unselected", "no", "false"}:
        return 0
    if any(lowered.startswith(prefix.casefold()) for prefix in REJECTION_PREFIXES):
        return 0
    if lowered == str(selected_label).strip().casefold():
        return 1
    return None


def _grid_group_label(question_spec: QuestionSpec, group_key: str) -> str:
    for candidate in (group_key, int(group_key) if str(group_key).isdigit() else group_key):
        if candidate in question_spec.option_map:
            return str(question_spec.option_map[candidate])
    return f"c{group_key}"


def _base_row_label(label: str) -> str:
    return re.sub(r"\s{2,}", " ", str(label)).strip(" -:|")
