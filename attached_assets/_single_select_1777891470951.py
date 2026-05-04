"""Single-select single-cut calculator."""

from __future__ import annotations

import pandas as pd

from src.calc_primitives import rate_per_value
from src.calculation_log import CalculationLog
from src.models import QuestionSpec, QuestionType, SingleSelectResult


def compute_single_select(
    question_spec: QuestionSpec,
    df: pd.DataFrame,
    log: CalculationLog,
    filter_mask: pd.Series | None = None,
    filter_expr: str | None = None,
) -> SingleSelectResult:
    column_name = question_spec.canonical_id
    if column_name not in df.columns:
        raise ValueError(f"raw column not found in data: {column_name}")

    series = df[column_name]
    if filter_mask is not None:
        series = series[filter_mask]
        if filter_expr is None:
            filter_expr = "<unnamed filter>"
    else:
        filter_expr = None

    result_dict, audit = rate_per_value(
        series=series,
        question_id=question_spec.canonical_id,
        source_columns=(column_name,),
        output_sheet="Single Cuts",
        filter_expr=filter_expr,
    )
    log.record(audit)

    warnings: list[str] = []
    distribution = {}
    for code, count_dict in result_dict.items():
        label = question_spec.option_map.get(code)
        if label is None:
            label = str(code)
            warnings.append(f"unmapped option code in raw data: {code!r}")
        distribution[code] = {
            "label": label,
            "count": count_dict["count"],
            "rate": count_dict["rate"],
        }

    return SingleSelectResult(
        question_id=question_spec.canonical_id,
        question_type=QuestionType.SINGLE_SELECT,
        valid_n=audit.valid_n,
        missing_n=audit.missing_n,
        denominator_policy=question_spec.denominator_policy,
        distribution=distribution,
        audit_records=(audit,),
        warnings=tuple(warnings),
    )
