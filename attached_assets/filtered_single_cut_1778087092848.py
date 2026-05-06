"""Filtered single-cut dispatch for the Survey Insight Engine."""

from __future__ import annotations

import pandas as pd

try:
    from config import LOW_SAMPLE_THRESHOLD
except ModuleNotFoundError as exc:
    if exc.name != "config":
        raise
    LOW_SAMPLE_THRESHOLD = 30

from src.calculation_log import CalculationLog
from src.cross_cut_engine import compute_cross_cuts
from src.models import (
    AnalysisType,
    CrossCutSpec,
    FilteredSingleCutResult,
    FilterSpec,
    QuestionType,
    SurveySchema,
)
from src.single_cut._grid import compute_grid
from src.single_cut._multi_select import compute_multi_select
from src.single_cut._numeric import compute_numeric
from src.single_cut._single_select import compute_single_select


def compute_filtered_single_cut(
    target_question_id: str,
    filters: list[FilterSpec],
    schema: SurveySchema,
    df: pd.DataFrame,
    log: CalculationLog,
) -> FilteredSingleCutResult:
    """Run a single cut with value filters or dispatch no-value filters to cross cuts."""

    target_spec = schema.get_question(target_question_id)
    if target_spec is None:
        raise ValueError(f"target question {target_question_id!r} not in schema")

    breakdown_filters = [filter_spec for filter_spec in filters if filter_spec.is_breakdown()]
    value_filters = [filter_spec for filter_spec in filters if not filter_spec.is_breakdown()]

    if len(breakdown_filters) > 1:
        raise ValueError("at most one breakdown filter is supported")

    mask, filter_expr, filtered_n = _build_value_filter_mask(value_filters, df)

    if breakdown_filters:
        return _compute_breakdown(
            target_question_id=target_question_id,
            target_type=target_spec.question_type,
            breakdown_filter=breakdown_filters[0],
            filters=filters,
            schema=schema,
            df=df,
            mask=mask,
            filter_expr=filter_expr,
            filtered_n=filtered_n,
            log=log,
        )

    result = _compute_filtered_single_cut_result(
        target_spec=target_spec,
        df=df,
        log=log,
        mask=mask,
        filter_expr=filter_expr,
    )

    return FilteredSingleCutResult(
        target_question_id=target_question_id,
        filters_applied=tuple(filters),
        dispatch_mode="single_cut_filtered",
        single_cut_result=result,
        cross_cut_result=None,
        filtered_n=filtered_n,
        audit_records=tuple(result.audit_records),
        warnings=_build_warnings(filtered_n),
    )


def _build_value_filter_mask(
    value_filters: list[FilterSpec],
    df: pd.DataFrame,
) -> tuple[pd.Series | None, str | None, int]:
    if not value_filters:
        return None, None, int(len(df))

    mask = pd.Series(True, index=df.index)
    filter_descriptions = []
    for filter_spec in value_filters:
        if filter_spec.filter_question_id not in df.columns:
            raise ValueError(
                f"filter column {filter_spec.filter_question_id!r} not in raw data"
            )
        mask = mask & (df[filter_spec.filter_question_id] == filter_spec.filter_value)
        filter_descriptions.append(
            f"{filter_spec.filter_question_id} == {filter_spec.filter_value!r}"
        )

    return mask, " AND ".join(filter_descriptions), int(mask.sum())


def _compute_breakdown(
    target_question_id: str,
    target_type: QuestionType,
    breakdown_filter: FilterSpec,
    filters: list[FilterSpec],
    schema: SurveySchema,
    df: pd.DataFrame,
    mask: pd.Series | None,
    filter_expr: str | None,
    filtered_n: int,
    log: CalculationLog,
) -> FilteredSingleCutResult:
    target_spec = schema.get_question(target_question_id)
    breakdown_spec = schema.get_question(breakdown_filter.filter_question_id)
    if target_spec is None:
        raise ValueError(f"target question {target_question_id!r} not in schema")
    if breakdown_spec is None:
        raise ValueError(
            f"breakdown question {breakdown_filter.filter_question_id!r} not in schema"
        )

    if target_type in (
        QuestionType.SINGLE_SELECT,
        QuestionType.DEMOGRAPHIC_OR_SEGMENT,
    ):
        analysis_type = AnalysisType.CROSS_TAB
    elif target_type in (
        QuestionType.DIRECT_NUMERIC,
        QuestionType.NUMERIC_ALLOCATION,
    ):
        analysis_type = AnalysisType.GROUP_COMPARISON
    else:
        raise ValueError(f"breakdown not supported for target type {target_type.value!r}")

    spec = CrossCutSpec(
        cross_cut_id=(
            f"FILTERED_{target_question_id}_BY_{breakdown_filter.filter_question_id}"
        ),
        title=(
            f"{target_spec.question_text!r} broken down by "
            f"{breakdown_spec.question_text!r}"
        ),
        analysis_type=analysis_type,
        source_question_ids=(breakdown_filter.filter_question_id, target_question_id),
        filter_expr=filter_expr,
        display_mode="all",
    )

    sliced_df = df[mask].copy() if mask is not None else df
    cc_results, cc_skips = compute_cross_cuts([spec], schema, sliced_df, log)

    if cc_skips or not cc_results:
        detail = cc_skips[0].details if cc_skips else "no result"
        raise ValueError(f"cross cut breakdown failed: {detail}")

    cross_cut_result = cc_results[0]
    return FilteredSingleCutResult(
        target_question_id=target_question_id,
        filters_applied=tuple(filters),
        dispatch_mode="cross_cut_breakdown",
        single_cut_result=None,
        cross_cut_result=cross_cut_result,
        filtered_n=filtered_n,
        audit_records=tuple(cross_cut_result.audit_records),
        warnings=_build_warnings(filtered_n),
    )


def _compute_filtered_single_cut_result(
    target_spec,
    df: pd.DataFrame,
    log: CalculationLog,
    mask: pd.Series | None,
    filter_expr: str | None,
):
    target_type = target_spec.question_type

    if target_type == QuestionType.SINGLE_SELECT:
        return compute_single_select(target_spec, df, log, mask, filter_expr)
    if target_type == QuestionType.MULTI_SELECT_BINARY:
        return compute_multi_select(target_spec, df, log, mask, filter_expr)
    if target_type in (
        QuestionType.DIRECT_NUMERIC,
        QuestionType.NUMERIC_ALLOCATION,
    ):
        return compute_numeric(target_spec, df, log, mask, filter_expr)
    if target_type == QuestionType.GRID_SINGLE_SELECT:
        return compute_grid(target_spec, df, log, mask, filter_expr)
    raise ValueError(f"filtered analysis not supported for type {target_type.value!r}")


def _build_warnings(filtered_n: int) -> tuple[str, ...]:
    warnings = []
    if filtered_n < LOW_SAMPLE_THRESHOLD:
        warnings.append(
            f"Filtered sample size {filtered_n} is below reliability threshold "
            f"({LOW_SAMPLE_THRESHOLD}); results may not be reliable."
        )
    return tuple(warnings)
