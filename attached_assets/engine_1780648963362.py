"""Public single-cut analysis orchestrator."""

from __future__ import annotations

import pandas as pd

from src.calculation_log import CalculationLog
from src.models import (
    QuestionSpec,
    QuestionType,
    SingleCutResult,
    SkipRecord,
    SurveySchema,
)


UNSUPPORTED_TYPES = {
    QuestionType.OPEN_TEXT,
    QuestionType.METADATA_OR_ID,
    QuestionType.UNKNOWN,
    QuestionType.DEMOGRAPHIC_OR_SEGMENT,
}


def compute_single_cuts(
    schema: SurveySchema,
    df: pd.DataFrame,
    log: CalculationLog,
    filter_mask: pd.Series | None = None,
    filter_expr: str | None = None,
) -> tuple[list[SingleCutResult], list[SkipRecord]]:
    results: list[SingleCutResult] = []
    skips: list[SkipRecord] = []

    for spec in schema.questions:
        if not spec.analysis_eligible:
            skips.append(
                SkipRecord(
                    question_id=spec.question_id,
                    canonical_id=spec.canonical_id,
                    question_type=spec.question_type,
                    skip_reason="ineligible",
                    details=spec.exclusion_reason,
                )
            )
            continue

        if spec.question_type in UNSUPPORTED_TYPES:
            skips.append(
                SkipRecord(
                    question_id=spec.question_id,
                    canonical_id=spec.canonical_id,
                    question_type=spec.question_type,
                    skip_reason=f"unsupported_type: {spec.question_type.value}",
                )
            )
            continue

        if _all_raw_columns_empty(spec, df, filter_mask):
            skips.append(
                SkipRecord(
                    question_id=spec.question_id,
                    canonical_id=spec.canonical_id,
                    question_type=spec.question_type,
                    skip_reason="all raw columns empty in dataset",
                    details="all raw columns empty in dataset",
                )
            )
            continue

        try:
            result = _dispatch(spec, df, log, filter_mask, filter_expr)
        except Exception as exc:
            skips.append(
                SkipRecord(
                    question_id=spec.question_id,
                    canonical_id=spec.canonical_id,
                    question_type=spec.question_type,
                    skip_reason="calculation_error",
                    details=f"{type(exc).__name__}: {exc}",
                )
            )
            continue

        results.append(result)

    return results, skips


def _all_raw_columns_empty(
    spec: QuestionSpec,
    df: pd.DataFrame,
    filter_mask: pd.Series | None,
) -> bool:
    present_columns = [column for column in spec.raw_columns if column in df.columns]
    if not present_columns:
        return False

    data = df[present_columns]
    if filter_mask is not None:
        data = data[filter_mask]
    return not bool(data.notna().any().any())


def _dispatch(
    spec: QuestionSpec,
    df: pd.DataFrame,
    log: CalculationLog,
    filter_mask: pd.Series | None,
    filter_expr: str | None,
) -> SingleCutResult:
    from src.single_cut.grid_binary_pivot import compute_grid_binary_pivot
    from src.single_cut.grid_rated import compute_grid_rated
    from src.single_cut.nps import compute_nps
    from src.single_cut.rank_order import compute_rank_order
    from src.single_cut._grid import compute_grid
    from src.single_cut._multi_select import compute_multi_select
    from src.single_cut._numeric import compute_numeric
    from src.single_cut._single_select import compute_single_select

    if spec.question_type is QuestionType.SINGLE_SELECT:
        return compute_single_select(spec, df, log, filter_mask, filter_expr)
    if spec.question_type is QuestionType.MULTI_SELECT_BINARY:
        return compute_multi_select(spec, df, log, filter_mask, filter_expr)
    if spec.question_type is QuestionType.RANK_ORDER:
        return compute_rank_order(spec, df, log, filter_mask, filter_expr)
    if spec.question_type in (
        QuestionType.DIRECT_NUMERIC,
        QuestionType.NUMERIC_ALLOCATION,
    ):
        return compute_numeric(spec, df, log, filter_mask, filter_expr)
    if spec.question_type is QuestionType.NPS:
        return compute_nps(spec, df, log, filter_mask, filter_expr)
    if spec.question_type is QuestionType.GRID_RATED:
        return compute_grid_rated(spec, df, log, filter_mask, filter_expr)
    if spec.question_type is QuestionType.GRID_BINARY_SELECT:
        return compute_grid_binary_pivot(spec, df, log, filter_mask, filter_expr)
    if spec.question_type is QuestionType.GRID_SINGLE_SELECT:
        return compute_grid(spec, df, log, filter_mask, filter_expr)
    raise ValueError(f"unsupported question type: {spec.question_type}")
