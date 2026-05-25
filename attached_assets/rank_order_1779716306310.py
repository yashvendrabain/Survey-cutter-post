"""Rank-order single-cut calculator."""

from __future__ import annotations

from typing import Any

import pandas as pd

from src.calc_primitives import _make_audit_record
from src.calculation_log import CalculationLog
from src.models import (
    DenominatorPolicy,
    QuestionSpec,
    QuestionType,
    RankOrderResult,
    RankOrderRow,
)
from src.single_cut._conditional import apply_conditional_on_filter


def compute_rank_order(
    question_spec: QuestionSpec,
    df: pd.DataFrame,
    log: CalculationLog,
    filter_mask: pd.Series | None = None,
    filter_expr: str | None = None,
) -> RankOrderResult:
    """Compute per-option counts for each rank position."""

    if question_spec.question_type is not QuestionType.RANK_ORDER:
        raise ValueError(
            f"unsupported question_type for rank order: {question_spec.question_type}"
        )
    if question_spec.value_range is None:
        raise ValueError("rank-order questions require value_range")

    low, high = question_spec.value_range
    if low != 1 or high < 1:
        raise ValueError("rank-order value_range must start at 1")
    rank_k = int(high)

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

    columns = [column for column in question_spec.raw_columns if column in working_df]
    if not columns:
        raise ValueError("rank-order raw columns not found in dataframe")

    ranked_values = working_df[columns].apply(lambda column: column.map(_coerce_rank_value))
    answered_mask = ranked_values.notna().any(axis=1)
    total_respondents = int(answered_mask.sum())
    total_responses = int(ranked_values.notna().sum().sum())
    filtered_n = int(len(working_df))
    missing_n = int(filtered_n - total_respondents)

    rows: list[RankOrderRow] = []
    for column in columns:
        series = ranked_values[column]
        counts_per_rank = [int((series == rank).sum()) for rank in range(1, rank_k + 1)]
        pcts_per_rank = [
            float(count / total_respondents) if total_respondents else 0.0
            for count in counts_per_rank
        ]
        for rank, count in enumerate(counts_per_rank, start=1):
            audit = _make_audit_record(
                metric_name=f"rank_{rank}_count",
                question_id=question_spec.canonical_id,
                source_columns=(column,),
                filter_expr=filter_expr,
                numerator=count,
                denominator=total_respondents,
                formula=f"count({column} == {rank}) / total_respondents",
                value_raw=float(pcts_per_rank[rank - 1]),
                valid_n=total_respondents,
                missing_n=missing_n,
                output_sheet=f"SC_{question_spec.canonical_id}",
            )
            log.record(audit)
        rows.append(
            RankOrderRow(
                option_id=column,
                option_label=_rank_option_label(question_spec, column),
                counts_per_rank=counts_per_rank,
                pcts_per_rank=pcts_per_rank,
            )
        )

    rows.sort(key=lambda row: (-sum(row.counts_per_rank), row.option_id))
    return RankOrderResult(
        question_id=question_spec.canonical_id,
        question_type=QuestionType.RANK_ORDER,
        valid_n=total_respondents,
        missing_n=missing_n,
        denominator_policy=DenominatorPolicy.VALID_RESPONSES,
        question_text=question_spec.question_text,
        K=rank_k,
        rows=rows,
        total_respondents=total_respondents,
        total_responses=total_responses,
    )


def _coerce_rank_value(value: Any) -> int | None:
    if pd.isna(value):
        return None
    try:
        rank = int(float(str(value).strip()))
    except (TypeError, ValueError):
        return None
    return rank


def _rank_option_label(question_spec: QuestionSpec, column: str) -> str:
    if column in question_spec.option_map:
        return str(question_spec.option_map[column])
    if question_spec.grid_row_labels and column in question_spec.grid_row_labels:
        return str(question_spec.grid_row_labels[column])
    return column
