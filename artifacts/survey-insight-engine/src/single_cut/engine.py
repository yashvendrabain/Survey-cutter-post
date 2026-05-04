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


def _dispatch(
    spec: QuestionSpec,
    df: pd.DataFrame,
    log: CalculationLog,
    filter_mask: pd.Series | None,
    filter_expr: str | None,
) -> SingleCutResult:
    from src.single_cut._grid import compute_grid
    from src.single_cut._multi_select import compute_multi_select
    from src.single_cut._numeric import compute_numeric
    from src.single_cut._single_select import compute_single_select

    if spec.question_type is QuestionType.SINGLE_SELECT:
        return compute_single_select(spec, df, log, filter_mask, filter_expr)
    if spec.question_type is QuestionType.MULTI_SELECT_BINARY:
        return compute_multi_select(spec, df, log, filter_mask, filter_expr)
    if spec.question_type in (
        QuestionType.DIRECT_NUMERIC,
        QuestionType.NUMERIC_ALLOCATION,
    ):
        return compute_numeric(spec, df, log, filter_mask, filter_expr)
    if spec.question_type is QuestionType.GRID_SINGLE_SELECT:
        return compute_grid(spec, df, log, filter_mask, filter_expr)
    raise ValueError(f"unsupported question type: {spec.question_type}")
