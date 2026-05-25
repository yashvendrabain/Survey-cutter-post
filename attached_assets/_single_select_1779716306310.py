"""Single-select single-cut calculator."""

from __future__ import annotations

import pandas as pd

from src.calc_primitives import rate_per_value
from src.calculation_log import CalculationLog
from src.models import QuestionSpec, QuestionType, SingleSelectResult
from src.single_cut._conditional import apply_conditional_on_filter


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
    series = working_df[column_name]

    result_dict, audit = rate_per_value(
        series=series,
        question_id=question_spec.canonical_id,
        source_columns=(column_name,),
        output_sheet="Single Cuts",
        filter_expr=filter_expr,
    )
    log.record(audit)

    if audit.valid_n == 0:
        return SingleSelectResult(
            question_id=question_spec.canonical_id,
            question_type=QuestionType.SINGLE_SELECT,
            valid_n=0,
            missing_n=int(series.isna().sum()),
            denominator_policy=question_spec.denominator_policy,
            distribution={},
            warnings=("all values are missing; skipped",),
            audit_records=(audit,),
        )

    warnings: list[str] = []
    distribution = {}
    for code, count_dict in result_dict.items():
        label = _label_for_value(code, question_spec.option_map)
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


def _label_for_value(value: object, option_map: dict[int | str, str]) -> str | None:
    for candidate in _option_lookup_candidates(value):
        if candidate in option_map:
            return option_map[candidate]

    value_text = str(value)
    for label in option_map.values():
        if value_text == str(label):
            return str(label)

    return None


def _option_lookup_candidates(value: object) -> tuple[object, ...]:
    candidates: list[object] = [value]
    if isinstance(value, float) and value.is_integer():
        candidates.append(int(value))
    if isinstance(value, str):
        stripped = value.strip()
        candidates.append(stripped)
        try:
            numeric = float(stripped)
        except ValueError:
            numeric = None
        if numeric is not None and numeric.is_integer():
            candidates.append(int(numeric))
    candidates.append(str(value))

    deduped: list[object] = []
    for candidate in candidates:
        if candidate not in deduped:
            deduped.append(candidate)
    return tuple(deduped)
