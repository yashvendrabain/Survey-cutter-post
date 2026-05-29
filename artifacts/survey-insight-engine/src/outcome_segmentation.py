"""Deterministic outcome segmentation and differentiator ranking."""

from __future__ import annotations

from datetime import datetime, timezone
import math
from typing import Any

import numpy as np
import pandas as pd

try:
    from scipy.stats import chi2_contingency
except ModuleNotFoundError:
    chi2_contingency = None

from src.models import (
    AuditRecord,
    DifferentiatorResult,
    OutcomeSegmentationResult,
    ProfileTrait,
    QuestionSpec,
    QuestionType,
    SegmentDefinition,
    SurveySchema,
    WinnerProfile,
)


_SUPPORTED_TYPES = {
    QuestionType.SINGLE_SELECT,
    QuestionType.MULTI_SELECT_BINARY,
    QuestionType.DIRECT_NUMERIC,
}


def compute_outcome_segmentation(
    decoded_df: pd.DataFrame,
    schema: SurveySchema,
    outcome_question_id: str,
    segment_definition: SegmentDefinition,
    audit_log: list[AuditRecord],
    min_sample_size: int = 30,
) -> OutcomeSegmentationResult:
    """Compute differentiator ranking and winner profile for an outcome variable."""

    manual_mode = segment_definition.segment_mode == "manual_uuid"
    outcome_spec = schema.get_question(outcome_question_id)
    if outcome_spec is None and not manual_mode:
        raise ValueError(f"outcome question {outcome_question_id!r} not in schema")

    winner_mask, loser_mask, valid_mask, segment_warnings = _build_segment_masks(
        decoded_df,
        outcome_spec,
        segment_definition,
        respondent_id_column=schema.respondent_id_column,
    )
    winner_n = int(winner_mask.sum())
    loser_n = int(loser_mask.sum())
    total_n = int(valid_mask.sum())

    warnings: list[str] = list(segment_warnings)
    if any(
        warning.startswith("Insufficient data for quartile split")
        for warning in warnings
    ):
        profile = WinnerProfile(
            outcome_question_id=outcome_question_id,
            winner_label=segment_definition.winner_label,
            winner_n=winner_n,
            loser_n=loser_n,
            defining_traits=(),
            loser_label=segment_definition.loser_label,
        )
        return OutcomeSegmentationResult(
            outcome_question_id=outcome_question_id,
            segment_definition=segment_definition,
            winner_n=winner_n,
            loser_n=loser_n,
            total_n=total_n,
            differentiators=(),
            winner_profile=profile,
            skipped_questions=(),
            warnings=tuple(warnings),
        )

    if winner_n < min_sample_size:
        warnings.append(
            f"Winner sample size {winner_n} is below minimum {min_sample_size}."
        )
    if loser_n < min_sample_size:
        warnings.append(
            f"{segment_definition.loser_label} sample size {loser_n} is below minimum {min_sample_size}."
        )
    if warnings:
        profile = WinnerProfile(
            outcome_question_id=outcome_question_id,
            winner_label=segment_definition.winner_label,
            winner_n=winner_n,
            loser_n=loser_n,
            defining_traits=(),
            loser_label=segment_definition.loser_label,
        )
        return OutcomeSegmentationResult(
            outcome_question_id=outcome_question_id,
            segment_definition=segment_definition,
            winner_n=winner_n,
            loser_n=loser_n,
            total_n=total_n,
            differentiators=(),
            winner_profile=profile,
            skipped_questions=(),
            warnings=tuple(warnings),
        )

    differentiators: list[DifferentiatorResult] = []
    skipped: list[tuple[str, str]] = []

    for question in schema.questions:
        if question.canonical_id == outcome_question_id:
            continue
        if question.question_type not in _SUPPORTED_TYPES:
            skipped.append((question.canonical_id, "unsupported_type"))
            continue

        try:
            result = _compute_question_differentiator(
                decoded_df=decoded_df,
                question=question,
                winner_mask=winner_mask,
                loser_mask=loser_mask,
            )
        except _SkipQuestion as exc:
            skipped.append((question.canonical_id, exc.reason))
            continue

        differentiators.append(result)
        audit_log.append(
            _make_audit_record(
                question=question,
                segment_definition=segment_definition,
                winner_count=int(round(result.top_option_winner_rate * winner_n)),
                winner_n=winner_n,
                loser_n=loser_n,
                total_rows=len(decoded_df),
                cramers_v=result.cramers_v,
            )
        )

    ranked = tuple(
        sorted(
            differentiators,
            key=lambda item: (
                -item.cramers_v,
                -abs(_finite_lift_for_sort(item.top_option_lift) - 1.0),
                item.question_id,
            ),
        )
    )
    profile, profile_warnings = _build_winner_profile(
        outcome_question_id=outcome_question_id,
        winner_label=segment_definition.winner_label,
        loser_label=segment_definition.loser_label,
        winner_n=winner_n,
        loser_n=loser_n,
        differentiators=ranked,
        decoded_df=decoded_df,
        schema=schema,
        winner_mask=winner_mask,
        loser_mask=loser_mask,
        audit_log=audit_log,
    )
    warnings.extend(profile_warnings)

    return OutcomeSegmentationResult(
        outcome_question_id=outcome_question_id,
        segment_definition=segment_definition,
        winner_n=winner_n,
        loser_n=loser_n,
        total_n=total_n,
        differentiators=ranked,
        winner_profile=profile,
        skipped_questions=tuple(skipped),
        warnings=tuple(warnings),
    )


class _SkipQuestion(Exception):
    def __init__(self, reason: str) -> None:
        super().__init__(reason)
        self.reason = reason


def _build_segment_masks(
    df: pd.DataFrame,
    outcome_spec: QuestionSpec | None,
    segment_definition: SegmentDefinition,
    respondent_id_column: str | None = None,
) -> tuple[pd.Series, pd.Series, pd.Series, tuple[str, ...]]:
    if segment_definition.segment_mode == "manual_uuid":
        column = (
            segment_definition.manual_cohort_id_column
            or _manual_respondent_id_column(df, respondent_id_column)
        )
        if column not in df.columns:
            raise ValueError(f"manual cohort id column {column!r} not in data")
        winner_set = {str(value).strip() for value in segment_definition.manual_winner_uuids}
        laggard_set = {str(value).strip() for value in segment_definition.manual_laggard_uuids}
        respondent_ids = df[column].astype(str).str.strip()
        winner_mask = respondent_ids.isin(winner_set)
        loser_mask = respondent_ids.isin(laggard_set)
        valid_mask = winner_mask | loser_mask
        return winner_mask, loser_mask, valid_mask, ()

    if outcome_spec is None:
        raise ValueError("outcome question is required unless segment_mode is manual_uuid")
    column = _primary_column(outcome_spec)
    if column not in df.columns:
        raise ValueError(f"outcome column {column!r} not in data")

    warnings: list[str] = []
    series = df[column]
    valid_mask = series.notna()
    if segment_definition.segment_mode == "categorical":
        winner_values = _expanded_option_values(
            segment_definition.winner_values,
            outcome_spec.option_map,
        )
        winner_mask = series.isin(winner_values) & valid_mask
        if segment_definition.laggard_values:
            laggard_values = _expanded_option_values(
                segment_definition.laggard_values,
                outcome_spec.option_map,
            )
            loser_mask = series.isin(laggard_values) & valid_mask
        else:
            loser_mask = (~winner_mask) & valid_mask
    elif segment_definition.segment_mode == "numeric_threshold":
        numeric = pd.to_numeric(series, errors="coerce")
        valid_mask = numeric.notna()
        threshold = float(segment_definition.winner_threshold)
        if segment_definition.threshold_direction == "gte":
            winner_mask = (numeric >= threshold) & valid_mask
        else:
            winner_mask = (numeric <= threshold) & valid_mask
        if segment_definition.laggard_threshold is not None:
            laggard_threshold = float(segment_definition.laggard_threshold)
            if segment_definition.laggard_threshold_direction == "gte":
                loser_mask = (numeric >= laggard_threshold) & valid_mask
            else:
                loser_mask = (numeric <= laggard_threshold) & valid_mask
        else:
            loser_mask = (~winner_mask) & valid_mask
    else:
        numeric = pd.to_numeric(series, errors="coerce")
        valid_outcome = numeric.dropna()
        if len(valid_outcome) < 8:
            empty_mask = pd.Series(False, index=df.index)
            return (
                empty_mask,
                empty_mask,
                empty_mask,
                ("Insufficient data for quartile split (need 8+ valid rows)",),
            )

        q1_threshold = valid_outcome.quantile(0.25)
        q3_threshold = valid_outcome.quantile(0.75)
        if segment_definition.quartile_winner == "top":
            winner_mask = (numeric >= q3_threshold) & numeric.notna()
            loser_mask = (numeric <= q1_threshold) & numeric.notna()
        else:
            winner_mask = (numeric <= q1_threshold) & numeric.notna()
            loser_mask = (numeric >= q3_threshold) & numeric.notna()

        valid_mask = winner_mask | loser_mask
        middle_excluded = int(
            len(valid_outcome) - int(winner_mask.sum()) - int(loser_mask.sum())
        )
        if middle_excluded > 0:
            warnings.append(
                f"Quartile split: {middle_excluded} middle-50% respondents "
                f"excluded from analysis (Q1={q1_threshold:.2f}, Q3={q3_threshold:.2f})"
            )

    return winner_mask, loser_mask, valid_mask, tuple(warnings)


def _manual_respondent_id_column(
    df: pd.DataFrame,
    respondent_id_column: str | None,
) -> str:
    if respondent_id_column and respondent_id_column in df.columns:
        return respondent_id_column
    if "record" in df.columns:
        return "record"
    if len(df.columns) > 0:
        return str(df.columns[0])
    raise ValueError("manual_uuid segmentation requires a respondent-id column")


def _compute_question_differentiator(
    decoded_df: pd.DataFrame,
    question: QuestionSpec,
    winner_mask: pd.Series,
    loser_mask: pd.Series,
) -> DifferentiatorResult:
    if question.question_type is QuestionType.SINGLE_SELECT:
        return _single_select_differentiator(decoded_df, question, winner_mask, loser_mask)
    if question.question_type is QuestionType.MULTI_SELECT_BINARY:
        return _multi_select_differentiator(decoded_df, question, winner_mask, loser_mask)
    if question.question_type is QuestionType.DIRECT_NUMERIC:
        return _numeric_differentiator(decoded_df, question, winner_mask, loser_mask)
    raise _SkipQuestion("unsupported_type")


def _single_select_differentiator(
    df: pd.DataFrame,
    question: QuestionSpec,
    winner_mask: pd.Series,
    loser_mask: pd.Series,
) -> DifferentiatorResult:
    column = _primary_column(question)
    if column not in df.columns:
        raise _SkipQuestion("missing_column")
    series = df[column]
    return _categorical_differentiator(
        question=question,
        series=series,
        label_map=question.option_map,
        question_type="single_select",
        winner_mask=winner_mask,
        loser_mask=loser_mask,
    )


def _multi_select_differentiator(
    df: pd.DataFrame,
    question: QuestionSpec,
    winner_mask: pd.Series,
    loser_mask: pd.Series,
) -> DifferentiatorResult:
    best: DifferentiatorResult | None = None
    missing_columns = 0
    for column in question.raw_columns:
        if column not in df.columns:
            missing_columns += 1
            continue
        selected = ((df[column] != 0) & df[column].notna()).astype(int)
        label = question.option_map.get(column, column)
        try:
            result = _categorical_differentiator(
                question=question,
                series=selected,
                label_map={1: label, 0: f"Not {label}"},
                question_type="multi_select_binary",
                winner_mask=winner_mask,
                loser_mask=loser_mask,
                top_value=1,
            )
        except _SkipQuestion:
            continue
        if best is None or result.cramers_v > best.cramers_v:
            best = result

    if best is None:
        raise _SkipQuestion("missing_column" if missing_columns else "insufficient_variation")
    return best


def _numeric_differentiator(
    df: pd.DataFrame,
    question: QuestionSpec,
    winner_mask: pd.Series,
    loser_mask: pd.Series,
) -> DifferentiatorResult:
    column = _primary_column(question)
    if column not in df.columns:
        raise _SkipQuestion("missing_column")
    numeric = pd.to_numeric(df[column], errors="coerce")
    valid = numeric.dropna()
    if valid.nunique() < 2:
        raise _SkipQuestion("insufficient_variation")

    try:
        binned = pd.qcut(numeric, q=4, labels=False, duplicates="drop")
    except ValueError as exc:
        raise _SkipQuestion("insufficient_variation") from exc

    if binned.dropna().nunique() < 2:
        raise _SkipQuestion("insufficient_variation")

    max_bin = int(binned.dropna().max())
    label_map = {
        bin_id: _quartile_label(int(bin_id), max_bin)
        for bin_id in sorted(binned.dropna().unique())
    }
    return _categorical_differentiator(
        question=question,
        series=binned,
        label_map=label_map,
        question_type="direct_numeric",
        winner_mask=winner_mask,
        loser_mask=loser_mask,
    )


def _categorical_differentiator(
    question: QuestionSpec,
    series: pd.Series,
    label_map: dict[Any, str],
    question_type: str,
    winner_mask: pd.Series,
    loser_mask: pd.Series,
    top_value: Any | None = None,
) -> DifferentiatorResult:
    valid_winner = series[winner_mask].notna()
    valid_loser = series[loser_mask].notna()
    winner_valid_n = int(valid_winner.sum())
    loser_valid_n = int(valid_loser.sum())
    if winner_valid_n < 10:
        raise _SkipQuestion("low_winner_sample")
    if loser_valid_n < 10:
        raise _SkipQuestion("low_loser_sample")

    values = sorted(series[series.notna()].unique(), key=lambda value: str(value))
    rows: list[list[int]] = []
    row_values: list[Any] = []
    for value in values:
        winner_count = int((series[winner_mask] == value).sum())
        loser_count = int((series[loser_mask] == value).sum())
        if winner_count + loser_count > 0:
            rows.append([winner_count, loser_count])
            row_values.append(value)

    table = pd.DataFrame(rows, index=row_values, columns=["winner", "loser"])
    table = table.loc[table.sum(axis=1) > 0]
    if len(table) < 2:
        raise _SkipQuestion("insufficient_variation")

    cramers_v, p_value = _cramers_v(table)
    if p_value is None and cramers_v == 0.0:
        raise _SkipQuestion("chi2_invalid")

    option_stats = _option_stats(table, winner_valid_n, loser_valid_n)
    if top_value is not None and top_value in option_stats:
        selected_value = top_value
    else:
        selected_value = max(
            option_stats,
            key=lambda value: abs(
                option_stats[value]["winner_rate"] - option_stats[value]["loser_rate"]
            ),
        )
    selected = option_stats[selected_value]
    lift, warnings = _lift(selected["winner_rate"], selected["loser_rate"])

    return DifferentiatorResult(
        question_id=question.canonical_id,
        question_text=question.question_text,
        question_type=question_type,
        cramers_v=cramers_v,
        top_option_label=str(label_map.get(selected_value, selected_value)),
        top_option_winner_rate=selected["winner_rate"],
        top_option_loser_rate=selected["loser_rate"],
        top_option_lift=lift,
        winner_n=winner_valid_n,
        loser_n=loser_valid_n,
        p_value=p_value,
        warnings=tuple(warnings),
    )


def _cramers_v(contingency_table: pd.DataFrame) -> tuple[float, float | None]:
    n = float(contingency_table.to_numpy().sum())
    rows, cols = contingency_table.shape
    if n == 0 or min(rows, cols) == 1:
        return 0.0, None
    try:
        if chi2_contingency is None:
            chi2, p_value, _dof, _expected = _chi2_contingency_fallback(contingency_table)
        else:
            chi2, p_value, _dof, _expected = chi2_contingency(contingency_table)
    except ValueError as exc:
        raise _SkipQuestion("chi2_invalid") from exc
    if chi2 < 0:
        return 0.0, None
    denominator = n * min(rows - 1, cols - 1)
    if denominator <= 0:
        return 0.0, None
    cramers_v = math.sqrt(float(chi2) / denominator)
    cramers_v = min(max(cramers_v, 0.0), 1.0)
    if p_value is None or pd.isna(p_value):
        return cramers_v, None
    return cramers_v, float(p_value)


def _chi2_contingency_fallback(
    contingency_table: pd.DataFrame,
) -> tuple[float, float, int, np.ndarray]:
    observed = contingency_table.to_numpy(dtype=float)
    row_sums = observed.sum(axis=1, keepdims=True)
    col_sums = observed.sum(axis=0, keepdims=True)
    total = observed.sum()
    if total <= 0:
        raise ValueError("observed table has no observations")
    expected = row_sums @ col_sums / total
    if np.any(expected == 0):
        raise ValueError("expected frequency is zero")
    chi2 = float(((observed - expected) ** 2 / expected).sum())
    dof = int((observed.shape[0] - 1) * (observed.shape[1] - 1))
    if dof == 1:
        p_value = float(math.erfc(math.sqrt(max(chi2, 0.0) / 2.0)))
    else:
        p_value = float("nan")
    return chi2, p_value, dof, expected


def _option_stats(
    table: pd.DataFrame,
    winner_n: int,
    loser_n: int,
) -> dict[Any, dict[str, float]]:
    return {
        value: {
            "winner_rate": float(row["winner"] / winner_n) if winner_n else 0.0,
            "loser_rate": float(row["loser"] / loser_n) if loser_n else 0.0,
        }
        for value, row in table.iterrows()
    }


def _lift(winner_rate: float, loser_rate: float) -> tuple[float, list[str]]:
    if loser_rate == 0 and winner_rate > 0:
        return 999.0, ["infinite_lift_loser_zero"]
    if loser_rate == 0 and winner_rate == 0:
        return 1.0, []
    return float(winner_rate / loser_rate), []


def _build_winner_profile(
    outcome_question_id: str,
    winner_label: str,
    loser_label: str,
    winner_n: int,
    loser_n: int,
    differentiators: tuple[DifferentiatorResult, ...],
    decoded_df: pd.DataFrame,
    schema: SurveySchema,
    winner_mask: pd.Series,
    loser_mask: pd.Series,
    audit_log: list[AuditRecord],
) -> tuple[WinnerProfile, list[str]]:
    qualifying = [
        diff
        for diff in differentiators
        if diff.top_option_lift > 1.2 and diff.cramers_v > 0.1
    ]
    if len(qualifying) < 3:
        qualifying = [
            diff
            for diff in differentiators
            if diff.top_option_lift > 1.1 and diff.cramers_v > 0.1
        ]

    traits: list[ProfileTrait] = []
    for diff in qualifying[:5]:
        question = schema.get_question(diff.question_id)
        if question is None:
            lag_label, lag_w_rate, lag_l_rate, lag_gap = ("", 0.0, 0.0, 0.0)
        else:
            lag_label, lag_w_rate, lag_l_rate, lag_gap = _compute_laggard_top_option(
                decoded_df=decoded_df,
                question=question,
                winner_mask=winner_mask,
                loser_mask=loser_mask,
                winner_n=winner_n,
                loser_n=loser_n,
            )
            if lag_label:
                audit_log.append(
                    _make_laggard_top_option_audit_record(
                        question=question,
                        outcome_question_id=outcome_question_id,
                        winner_n=winner_n,
                        loser_n=loser_n,
                        total_rows=len(decoded_df),
                        laggard_rate=lag_l_rate,
                    )
                )

        traits.append(
            ProfileTrait(
                question_id=diff.question_id,
                question_text=diff.question_text,
                option_label=diff.top_option_label,
                winner_rate=diff.top_option_winner_rate,
                loser_rate=diff.top_option_loser_rate,
                lift=diff.top_option_lift,
                rate_gap=diff.top_option_winner_rate - diff.top_option_loser_rate,
                laggard_top_option_label=lag_label,
                laggard_top_option_winner_rate=lag_w_rate,
                laggard_top_option_loser_rate=lag_l_rate,
                laggard_top_option_gap=lag_gap,
            )
        )
    warnings: list[str] = []
    if len(traits) < 3:
        warnings.append("winner_profile_has_fewer_than_3_traits")

    return (
        WinnerProfile(
            outcome_question_id=outcome_question_id,
            winner_label=winner_label,
            winner_n=winner_n,
            loser_n=loser_n,
            defining_traits=tuple(traits),
            loser_label=loser_label,
        ),
        warnings,
    )


def _compute_laggard_top_option(
    decoded_df: pd.DataFrame,
    question: QuestionSpec,
    winner_mask: pd.Series,
    loser_mask: pd.Series,
    winner_n: int,
    loser_n: int,
) -> tuple[str, float, float, float]:
    """
    Return the option most chosen by laggards for a differentiator question.
    """
    if winner_n == 0 or loser_n == 0:
        return ("", 0.0, 0.0, 0.0)

    try:
        if question.question_type is QuestionType.SINGLE_SELECT:
            col = _primary_column(question)
            if col not in decoded_df.columns:
                return ("", 0.0, 0.0, 0.0)
            series = decoded_df[col]

            best_label = ""
            best_loser_rate = -1.0
            best_winner_rate = 0.0

            for code, label in question.option_map.items():
                mask = _option_value_mask(series, code, label)
                w_rate = float((mask & winner_mask).sum()) / winner_n
                l_rate = float((mask & loser_mask).sum()) / loser_n
                if l_rate > best_loser_rate:
                    best_loser_rate = l_rate
                    best_winner_rate = w_rate
                    best_label = str(label)

            if not best_label:
                return ("", 0.0, 0.0, 0.0)

            gap = best_loser_rate - best_winner_rate
            return (best_label, best_winner_rate, best_loser_rate, gap)

        if question.question_type is QuestionType.MULTI_SELECT_BINARY:
            best_label = ""
            best_loser_rate = -1.0
            best_winner_rate = 0.0

            for col, label in question.option_map.items():
                if col not in decoded_df.columns:
                    continue
                selected = (decoded_df[col] != 0) & decoded_df[col].notna()
                w_rate = float((selected & winner_mask).sum()) / winner_n
                l_rate = float((selected & loser_mask).sum()) / loser_n
                if l_rate > best_loser_rate:
                    best_loser_rate = l_rate
                    best_winner_rate = w_rate
                    best_label = str(label)

            if not best_label:
                return ("", 0.0, 0.0, 0.0)

            gap = best_loser_rate - best_winner_rate
            return (best_label, best_winner_rate, best_loser_rate, gap)

        col = _primary_column(question)
        if col not in decoded_df.columns:
            return ("", 0.0, 0.0, 0.0)
        numeric = pd.to_numeric(decoded_df[col], errors="coerce")
        valid = numeric.dropna()
        if len(valid) < 4:
            return ("", 0.0, 0.0, 0.0)

        bins = pd.qcut(
            numeric,
            q=4,
            labels=[
                "Q1 (bottom quartile)",
                "Q2",
                "Q3",
                "Q4 (top quartile)",
            ],
            duplicates="drop",
        )

        best_label = ""
        best_loser_rate = -1.0
        best_winner_rate = 0.0

        categories = bins.cat.categories if hasattr(bins, "cat") else []
        for bin_label in categories:
            mask = bins == bin_label
            w_rate = float((mask & winner_mask).sum()) / winner_n
            l_rate = float((mask & loser_mask).sum()) / loser_n
            if l_rate > best_loser_rate:
                best_loser_rate = l_rate
                best_winner_rate = w_rate
                best_label = str(bin_label)

        if not best_label:
            return ("", 0.0, 0.0, 0.0)

        gap = best_loser_rate - best_winner_rate
        return (best_label, best_winner_rate, best_loser_rate, gap)

    except Exception:
        return ("", 0.0, 0.0, 0.0)


def _option_value_mask(series: pd.Series, code: object, label: object) -> pd.Series:
    mask = series == code
    mask = mask | (series == label)
    code_text = str(code)
    label_text = str(label)
    return mask | (series.astype("string") == code_text) | (
        series.astype("string") == label_text
    )


def _expanded_option_values(
    values: tuple[int | str, ...],
    option_map: dict[int | str, str],
) -> tuple[int | str, ...]:
    expanded: list[int | str] = []
    for value in values:
        candidates: list[int | str] = [value]
        if value in option_map:
            candidates.append(option_map[value])
        if isinstance(value, str):
            stripped = value.strip()
            candidates.append(stripped)
            try:
                numeric = float(stripped)
            except ValueError:
                numeric = None
            if numeric is not None and numeric.is_integer():
                numeric_int = int(numeric)
                candidates.append(numeric_int)
                if numeric_int in option_map:
                    candidates.append(option_map[numeric_int])
        for candidate in candidates:
            if candidate not in expanded:
                expanded.append(candidate)
    return tuple(expanded)


def _make_laggard_top_option_audit_record(
    question: QuestionSpec,
    outcome_question_id: str,
    winner_n: int,
    loser_n: int,
    total_rows: int,
    laggard_rate: float,
) -> AuditRecord:
    valid_n = int(winner_n + loser_n)
    laggard_count = int(round(laggard_rate * loser_n))
    return AuditRecord(
        output_sheet=f"OS_{outcome_question_id}",
        metric_name="laggard_top_option_rate",
        source_question_id=question.canonical_id,
        source_columns=question.raw_columns or (question.canonical_id,),
        filter_expr=f"winner_profile: {outcome_question_id}",
        numerator=laggard_count,
        denominator=loser_n,
        formula="laggard_rate = (laggard_mask & option_mask).sum() / loser_n",
        value_raw=laggard_rate,
        valid_n=valid_n,
        missing_n=int(total_rows - valid_n),
        timestamp=datetime.now(timezone.utc),
    )


def _make_audit_record(
    question: QuestionSpec,
    segment_definition: SegmentDefinition,
    winner_count: int,
    winner_n: int,
    loser_n: int,
    total_rows: int,
    cramers_v: float,
) -> AuditRecord:
    valid_n = int(winner_n + loser_n)
    return AuditRecord(
        output_sheet=f"OS_{segment_definition.outcome_question_id}",
        metric_name="differentiator_cramers_v",
        source_question_id=question.canonical_id,
        source_columns=question.raw_columns,
        filter_expr=f"segment_definition: {segment_definition}",
        numerator=winner_count,
        denominator=winner_n,
        formula="cramers_v = sqrt(chi2 / (n * min(rows-1, cols-1)))",
        value_raw=cramers_v,
        valid_n=valid_n,
        missing_n=int(total_rows - valid_n),
        timestamp=datetime.now(timezone.utc),
    )


def _primary_column(question: QuestionSpec) -> str:
    return question.raw_columns[0] if question.raw_columns else question.canonical_id


def _quartile_label(bin_id: int, max_bin: int) -> str:
    if bin_id == 0:
        return "Q1 (bottom quartile)"
    if bin_id == max_bin:
        return "Q4 (top quartile)"
    return f"Q{bin_id + 1}"


def _finite_lift_for_sort(lift: float) -> float:
    if np.isinf(lift):
        return 999.0
    return float(lift)
