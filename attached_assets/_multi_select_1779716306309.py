"""Multi-select single-cut calculator."""

from __future__ import annotations

import pandas as pd

from src.calc_primitives import selection_rate
from src.calculation_log import CalculationLog
from src.models import (
    DenominatorPolicy,
    MultiSelectResult,
    QuestionSpec,
    QuestionType,
)
from src.single_cut._conditional import apply_conditional_on_filter


def compute_multi_select(
    question_spec: QuestionSpec,
    df: pd.DataFrame,
    log: CalculationLog,
    filter_mask: pd.Series | None = None,
    filter_expr: str | None = None,
) -> MultiSelectResult:
    if question_spec.question_type is not QuestionType.MULTI_SELECT_BINARY:
        raise ValueError(
            f"unsupported question_type for multi-select: {question_spec.question_type}"
        )

    working_df = df
    if filter_mask is not None:
        working_df = working_df[filter_mask]
        if filter_expr is None:
            filter_expr = "<unnamed filter>"
    else:
        filter_expr = None
    working_df, filter_expr = apply_conditional_on_filter(
        question_spec,
        working_df,
        filter_expr,
    )

    warnings: list[str] = []
    binary_columns: dict[str, pd.Series] = {}
    for sub_column_id in question_spec.raw_columns:
        if sub_column_id not in working_df.columns:
            warnings.append(f"sub-column {sub_column_id} not found in data")
            continue
        binary_columns[sub_column_id] = working_df[sub_column_id]

    if not binary_columns:
        raise ValueError("no sub-columns present in raw data")

    denominator_policy = (
        "valid_responses"
        if question_spec.denominator_policy is DenominatorPolicy.VALID_RESPONSES
        else "all_respondents"
    )
    result_dict, audit = selection_rate(
        binary_columns=binary_columns,
        question_id=question_spec.canonical_id,
        output_sheet=f"SC_{question_spec.canonical_id}",
        denominator_policy=denominator_policy,
        all_respondents_n=len(working_df) if denominator_policy == "all_respondents" else None,
        filter_expr=filter_expr,
    )
    log.record(audit)

    respondents_who_answered_any = int(
        pd.DataFrame(binary_columns).notna().any(axis=1).sum()
    )
    if respondents_who_answered_any == 0:
        warnings.append("all sub-columns are 100% missing")

    selections = {}
    for sub_column_id, payload in result_dict.items():
        label = question_spec.option_map.get(sub_column_id)
        if label is None:
            label = sub_column_id
            warnings.append(f"no label for {sub_column_id}; using id as label")
        selections[sub_column_id] = {
            "label": label,
            "count": payload["count"],
            "selection_rate": payload["selection_rate"],
        }

    return MultiSelectResult(
        question_id=question_spec.canonical_id,
        question_type=QuestionType.MULTI_SELECT_BINARY,
        valid_n=audit.valid_n,
        missing_n=audit.missing_n,
        denominator_policy=question_spec.denominator_policy,
        selections=selections,
        respondents_who_answered_any=respondents_who_answered_any,
        audit_records=(audit,),
        warnings=tuple(warnings),
    )
