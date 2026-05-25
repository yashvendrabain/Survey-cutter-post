"""Shared conditional-display filtering for single-cut calculators."""

from __future__ import annotations

import pandas as pd

from src.models import QuestionSpec


def apply_conditional_on_filter(
    question_spec: QuestionSpec,
    df: pd.DataFrame,
    filter_expr: str | None,
) -> tuple[pd.DataFrame, str | None]:
    """Restrict to respondents who answered the gating question, when known."""

    conditional_id = question_spec.conditional_on
    if not conditional_id or conditional_id not in df.columns:
        return df, filter_expr

    conditional_expr = f"conditional_on({conditional_id} not null)"
    combined_expr = (
        f"{filter_expr} AND {conditional_expr}" if filter_expr else conditional_expr
    )
    return df[df[conditional_id].notna()], combined_expr
