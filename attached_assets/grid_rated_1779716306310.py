"""Rated grid single-cut calculator."""

from __future__ import annotations

from collections import defaultdict
import re
from typing import Any

import pandas as pd

from src.calc_primitives import _make_audit_record
from src.calculation_log import CalculationLog
from src.models import (
    DenominatorPolicy,
    GridRatedResult,
    GridRatedRow,
    QuestionSpec,
    QuestionType,
)
from src.single_cut._conditional import apply_conditional_on_filter


MISSING_VALUE_TOKENS = {
    "i don't know",
    "i don\u2019t know",
    "this was not something i considered",
    "not applicable",
    "n/a",
}


def compute_grid_rated(
    question_spec: QuestionSpec,
    df: pd.DataFrame,
    log: CalculationLog,
    filter_mask: pd.Series | None = None,
    filter_expr: str | None = None,
) -> GridRatedResult:
    """Compute mean values for each row x c-column of a rated grid."""

    if question_spec.question_type is not QuestionType.GRID_RATED:
        raise ValueError(
            f"unsupported question_type for rated grid: {question_spec.question_type}"
        )

    working_df = df
    if filter_mask is not None:
        working_df = working_df[filter_mask]
        filter_expr = filter_expr or "<unnamed filter>"
    else:
        filter_expr = None
    working_df, filter_expr = apply_conditional_on_filter(
        question_spec,
        working_df,
        filter_expr,
    )

    grouped_columns, group_order = _group_grid_columns(question_spec, working_df)
    grouped_columns, group_order = _drop_empty_c_columns(
        grouped_columns,
        group_order,
        working_df,
    )
    if not grouped_columns:
        raise ValueError("rated grid raw columns not found in dataframe")

    column_headers = [_grid_group_label(question_spec, group_key) for group_key in group_order]
    data = working_df[
        [column for group in grouped_columns.values() for column in group.values()]
    ].apply(lambda column: column.map(_coerce_to_numeric))
    total_respondents = int(data.notna().any(axis=1).sum())
    total_responses = int(data.notna().sum().sum())
    filtered_n = int(len(working_df))
    missing_n = int(filtered_n - total_respondents)

    rows: list[GridRatedRow] = []
    row_labels = question_spec.grid_row_labels or {}
    for row_id, columns_by_group in grouped_columns.items():
        means: list[float] = []
        valid_ns: list[int] = []
        for group_key in group_order:
            column = columns_by_group.get(group_key)
            if column is None:
                mean_value = 0.0
                valid_n = 0
                numerator = 0.0
            else:
                series = data[column].dropna()
                valid_n = int(len(series))
                numerator = float(series.sum()) if valid_n else 0.0
                mean_value = float(numerator / valid_n) if valid_n else 0.0
                audit = _make_audit_record(
                    metric_name=f"{row_id}_{group_key}_mean",
                    question_id=question_spec.canonical_id,
                    source_columns=(column,),
                    filter_expr=filter_expr,
                    numerator=numerator,
                    denominator=valid_n,
                    formula="sum(numeric values) / count(numeric values)",
                    value_raw=mean_value,
                    valid_n=valid_n,
                    missing_n=filtered_n - valid_n,
                    output_sheet=f"SC_{question_spec.canonical_id}",
                )
                log.record(audit)
            means.append(mean_value)
            valid_ns.append(valid_n)
        delta = means[0] - means[1] if len(means) == 2 else None
        first_column = next(iter(columns_by_group.values()))
        rows.append(
            GridRatedRow(
                row_id=row_id,
                row_label=_base_row_label(row_labels.get(first_column, row_id)),
                means_per_column=means,
                valid_n_per_column=valid_ns,
                delta=delta,
            )
        )

    return GridRatedResult(
        question_id=question_spec.canonical_id,
        question_type=QuestionType.GRID_RATED,
        valid_n=total_respondents,
        missing_n=missing_n,
        denominator_policy=DenominatorPolicy.VALID_RESPONSES,
        question_text=question_spec.question_text,
        column_headers=column_headers,
        rows=rows,
        total_respondents=total_respondents,
        total_responses=total_responses,
        show_delta=len(column_headers) == 2,
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


def _drop_empty_c_columns(
    grouped_columns: dict[str, dict[str, str]],
    group_order: list[str],
    working_df: pd.DataFrame,
) -> tuple[dict[str, dict[str, str]], list[str]]:
    non_empty_groups: list[str] = []
    for group_key in group_order:
        columns_in_group = [
            column
            for columns_by_group in grouped_columns.values()
            if (column := columns_by_group.get(group_key)) is not None
        ]
        if columns_in_group and working_df[columns_in_group].notna().any().any():
            non_empty_groups.append(group_key)

    new_group_order = [group_key for group_key in group_order if group_key in non_empty_groups]
    new_grouped = {}
    for row_id, columns_by_group in grouped_columns.items():
        pruned = {
            group_key: column
            for group_key, column in columns_by_group.items()
            if group_key in new_group_order
        }
        if pruned:
            new_grouped[row_id] = pruned
    return new_grouped, new_group_order


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


def _coerce_to_numeric(value: Any) -> float | None:
    if pd.isna(value):
        return None
    text = str(value).strip()
    if text.casefold() in MISSING_VALUE_TOKENS:
        return None
    match = re.match(r"^-?\d+(?:\.\d+)?", text)
    if match is None:
        return None
    try:
        return float(match.group(0))
    except (TypeError, ValueError):
        return None


def _grid_group_label(question_spec: QuestionSpec, group_key: str) -> str:
    if str(group_key) in question_spec.grid_column_labels:
        return str(question_spec.grid_column_labels[str(group_key)])
    for candidate in (group_key, int(group_key) if str(group_key).isdigit() else group_key):
        if candidate in question_spec.option_map:
            return str(question_spec.option_map[candidate])
    if group_key == "1":
        return "Winner - All"
    if group_key == "2":
        return "Other considered vendor"
    return f"Group {group_key}"


def _base_row_label(label: str) -> str:
    return re.sub(r"\s{2,}", " ", str(label)).strip(" -:|")
