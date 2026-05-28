"""Hypothesis validation for survey-analysis results."""

from __future__ import annotations

from dataclasses import asdict
from datetime import datetime, timezone
import json
import math
import os
from statistics import NormalDist
from typing import Any
from uuid import uuid4

import numpy as np
import pandas as pd

from config import PORTKEY_BASE_URL, PORTKEY_DEFAULT_MODEL
from src.calculation_log import CalculationLog
from src.models import (
    AuditRecord,
    HypothesisGroup,
    HypothesisResult,
    HypothesisSpec,
    HypothesisStatistic,
    QuestionSpec,
    QuestionType,
    SurveySchema,
)

try:
    from openai import OpenAI
except ModuleNotFoundError:
    OpenAI = None  # type: ignore[assignment]

try:
    from scipy import stats as scipy_stats
    try:
        from scipy.stats.contingency import association as scipy_association
    except Exception:  # pragma: no cover - depends on SciPy version.
        scipy_association = None
except ModuleNotFoundError:  # pragma: no cover - local runtime may omit SciPy.
    scipy_stats = None
    scipy_association = None


_MIN_SAMPLE_SIZE = 50
_EFFECT_THRESHOLDS = {
    "pearson_r": 0.3,
    "spearman_rho": 0.3,
    "chi_square_cramer_v": 0.2,
    "welch_t": 0.5,
    "anova_oneway": 0.06,
}


def validate_hypothesis(
    spec: HypothesisSpec,
    active_df: pd.DataFrame,
    schema: SurveySchema,
    calculation_log: CalculationLog,
    cross_cut_lookup: dict[Any, str] | None = None,
    active_filters_summary: str = "(all respondents)",
) -> HypothesisResult:
    """Run the full deterministic hypothesis-validation pipeline."""

    cohort_n = int(len(active_df))
    outcome_question = schema.get_question(spec.outcome_question_id)
    predictor_question = schema.get_question(spec.predictor_question_id)
    warnings: list[str] = []
    related_cross_cut_id = _lookup_cross_cut_id(
        cross_cut_lookup,
        spec.outcome_question_id,
        spec.predictor_question_id,
    )
    if related_cross_cut_id is None:
        warnings.append("no_matching_cross_cut_found")

    source_columns = _source_columns_for_questions(outcome_question, predictor_question)
    entry_ids: list[str] = []
    entry_ids.append(
        _record_hypothesis_metric(
            calculation_log,
            spec,
            "cohort_n",
            source_columns,
            active_filters_summary,
            "COUNT respondents in active_df after active filters",
            float(cohort_n),
            valid_n=cohort_n,
            numerator=cohort_n,
            denominator=None,
        )
    )

    if outcome_question is None or predictor_question is None:
        return HypothesisResult(
            spec=spec,
            verdict="INCONCLUSIVE",
            verdict_reason="could_not_classify_questions",
            statistic=None,
            cohort_n=cohort_n,
            pairwise_n=0,
            active_filters_summary=active_filters_summary,
            related_cross_cut_id=related_cross_cut_id,
            calculation_log_entry_ids=tuple(entry_ids),
            warnings=tuple(warnings),
        )

    outcome_role = _classify_question_role(outcome_question)
    predictor_role = _classify_question_role(predictor_question)
    if "unsupported" in {outcome_role, predictor_role}:
        return HypothesisResult(
            spec=spec,
            verdict="INCONCLUSIVE",
            verdict_reason="could_not_classify_questions",
            statistic=None,
            cohort_n=cohort_n,
            pairwise_n=0,
            active_filters_summary=active_filters_summary,
            related_cross_cut_id=related_cross_cut_id,
            calculation_log_entry_ids=tuple(entry_ids),
            warnings=tuple(warnings),
        )

    prepared = _prepare_test_data(spec, active_df, schema)
    pairwise_n = int(prepared["pairwise_n"])
    entry_ids.append(
        _record_hypothesis_metric(
            calculation_log,
            spec,
            "pairwise_n",
            tuple(prepared["source_columns"]),
            active_filters_summary,
            "COUNT respondents where outcome IS NOT NULL and predictor IS NOT NULL",
            float(pairwise_n),
            valid_n=pairwise_n,
            numerator=pairwise_n,
            denominator=cohort_n,
        )
    )

    predictor_n_groups = int(prepared.get("predictor_n_groups", 0))
    try:
        test_name = _select_test(outcome_role, predictor_role, predictor_n_groups)
    except ValueError:
        return HypothesisResult(
            spec=spec,
            verdict="INCONCLUSIVE",
            verdict_reason="could_not_classify_questions",
            statistic=None,
            cohort_n=cohort_n,
            pairwise_n=pairwise_n,
            active_filters_summary=active_filters_summary,
            related_cross_cut_id=related_cross_cut_id,
            calculation_log_entry_ids=tuple(entry_ids),
            warnings=tuple(warnings),
        )

    statistic = _run_selected_test(test_name, prepared)
    entry_ids.extend(
        _record_statistic_metrics(
            calculation_log,
            spec,
            statistic,
            tuple(prepared["source_columns"]),
            active_filters_summary,
        )
    )

    sample_sizes = (
        list(statistic.sample_sizes_per_group)
        if statistic.sample_sizes_per_group
        else [statistic.sample_size]
    )
    if not _check_sample_size_floor(sample_sizes):
        verdict, reason = "INCONCLUSIVE", "insufficient_sample_size"
    else:
        verdict, reason = _verdict_from_statistic(
            statistic,
            spec.outcome_direction,
        )

    result = HypothesisResult(
        spec=spec,
        verdict=verdict,
        verdict_reason=reason,
        statistic=statistic,
        cohort_n=cohort_n,
        pairwise_n=pairwise_n,
        active_filters_summary=active_filters_summary,
        related_cross_cut_id=related_cross_cut_id,
        calculation_log_entry_ids=tuple(entry_ids),
        ai_prose_explanation="",
        warnings=tuple(warnings),
    )
    prose = explain_verdict(result)
    if prose:
        result = HypothesisResult(
            spec=result.spec,
            verdict=result.verdict,
            verdict_reason=result.verdict_reason,
            statistic=result.statistic,
            cohort_n=result.cohort_n,
            pairwise_n=result.pairwise_n,
            active_filters_summary=result.active_filters_summary,
            related_cross_cut_id=result.related_cross_cut_id,
            calculation_log_entry_ids=result.calculation_log_entry_ids,
            ai_prose_explanation=prose,
            warnings=result.warnings,
        )
    return result


def _classify_question_role(question: QuestionSpec) -> str:
    if question.question_type in {
        QuestionType.DIRECT_NUMERIC,
        QuestionType.NUMERIC_ALLOCATION,
        QuestionType.GRID_RATED,
    }:
        return "continuous"
    if question.question_type is QuestionType.SINGLE_SELECT:
        if getattr(question, "label_to_numeric_value", None):
            return "ordinal"
        return "categorical"
    if question.question_type is QuestionType.MULTI_SELECT_BINARY:
        return "categorical"
    if question.question_type is QuestionType.RANK_ORDER:
        return "ordinal"
    return "unsupported"


def _select_test(
    outcome_role: str,
    predictor_role: str,
    predictor_n_groups: int,
) -> str:
    """Return the statistical test selected by the locked auto-selector."""

    if outcome_role == "continuous" and predictor_role == "continuous":
        return "pearson_r"
    if outcome_role == "continuous" and predictor_role == "ordinal":
        return "spearman_rho"
    if outcome_role == "ordinal" and predictor_role == "ordinal":
        return "spearman_rho"
    if outcome_role == "continuous" and predictor_role == "categorical":
        if predictor_n_groups == 2:
            return "welch_t"
        if predictor_n_groups >= 3:
            return "anova_oneway"
    if outcome_role == "categorical" and predictor_role == "categorical":
        return "chi_square_cramer_v"
    raise ValueError(
        f"unsupported hypothesis role pair: {outcome_role} x {predictor_role}"
    )


def _prepare_test_data(
    spec: HypothesisSpec,
    active_df: pd.DataFrame,
    schema: SurveySchema,
) -> dict[str, Any]:
    outcome_question = schema.get_question(spec.outcome_question_id)
    predictor_question = schema.get_question(spec.predictor_question_id)
    if outcome_question is None or predictor_question is None:
        raise ValueError("hypothesis question id not found in schema")

    outcome_role = _classify_question_role(outcome_question)
    predictor_role = _classify_question_role(predictor_question)
    outcome = _series_for_question(active_df, outcome_question, outcome_role)
    predictor = _series_for_question(active_df, predictor_question, predictor_role)
    aligned = pd.DataFrame({"outcome": outcome, "predictor": predictor}).dropna()

    prepared: dict[str, Any] = {
        "outcome": aligned["outcome"],
        "predictor": aligned["predictor"],
        "outcome_role": outcome_role,
        "predictor_role": predictor_role,
        "pairwise_n": int(len(aligned)),
        "source_columns": tuple(
            dict.fromkeys(
                tuple(outcome_question.raw_columns) + tuple(predictor_question.raw_columns)
            )
        ),
    }

    if predictor_role == "categorical":
        grouped_predictor = _apply_predictor_groups(aligned["predictor"], spec)
        grouped_frame = pd.DataFrame(
            {"outcome": aligned["outcome"], "predictor": grouped_predictor}
        ).dropna()
        prepared["outcome"] = grouped_frame["outcome"]
        prepared["predictor"] = grouped_frame["predictor"]
        prepared["pairwise_n"] = int(len(grouped_frame))
        groups = [
            group["outcome"].to_numpy()
            for _label, group in grouped_frame.groupby("predictor", sort=True)
        ]
        prepared["groups"] = groups
        prepared["predictor_n_groups"] = len(groups)
        if outcome_role == "categorical":
            prepared["contingency_table"] = pd.crosstab(
                grouped_frame["outcome"],
                grouped_frame["predictor"],
            ).to_numpy(dtype=float)
    else:
        prepared["predictor_n_groups"] = 0
    return prepared


def _run_selected_test(test_name: str, prepared: dict[str, Any]) -> HypothesisStatistic:
    outcome = prepared["outcome"]
    predictor = prepared["predictor"]
    if test_name == "pearson_r":
        return _run_pearson(outcome, predictor)
    if test_name == "spearman_rho":
        return _run_spearman(outcome, predictor)
    if test_name == "welch_t":
        groups = prepared["groups"]
        return _run_welch_t(groups[0], groups[1])
    if test_name == "anova_oneway":
        return _run_anova(prepared["groups"])
    if test_name == "chi_square_cramer_v":
        return _run_chi_square(prepared["contingency_table"])
    raise ValueError(f"unsupported test: {test_name}")


def _run_pearson(outcome: Any, predictor: Any) -> HypothesisStatistic:
    x = _numeric_array(outcome)
    y = _numeric_array(predictor)
    n = int(min(len(x), len(y)))
    if n < 3:
        r, p_value = 0.0, 1.0
    elif scipy_stats is not None:
        r, p_value = scipy_stats.pearsonr(x, y)
    else:
        r = float(np.corrcoef(x, y)[0, 1])
        p_value = _normal_approx_correlation_p_value(r, n)
    ci_low, ci_high = _compute_pearson_ci(float(r), n)
    return HypothesisStatistic(
        test_name="pearson_r",
        effect_size=float(r),
        effect_size_label="Pearson r",
        p_value=float(p_value),
        confidence_interval_low=ci_low,
        confidence_interval_high=ci_high,
        sample_size=n,
        raw_test_statistic=float(r),
    )


def _run_spearman(outcome: Any, predictor: Any) -> HypothesisStatistic:
    x = _numeric_array(outcome)
    y = _numeric_array(predictor)
    n = int(min(len(x), len(y)))
    if n < 3:
        rho, p_value = 0.0, 1.0
    elif scipy_stats is not None:
        rho, p_value = scipy_stats.spearmanr(x, y)
    else:
        rho = float(np.corrcoef(pd.Series(x).rank(), pd.Series(y).rank())[0, 1])
        p_value = _normal_approx_correlation_p_value(rho, n)
    ci_low, ci_high = _compute_pearson_ci(float(rho), n)
    return HypothesisStatistic(
        test_name="spearman_rho",
        effect_size=float(rho),
        effect_size_label="Spearman rho",
        p_value=float(p_value),
        confidence_interval_low=ci_low,
        confidence_interval_high=ci_high,
        sample_size=n,
        raw_test_statistic=float(rho),
    )


def _run_welch_t(group_a: Any, group_b: Any) -> HypothesisStatistic:
    a = _numeric_array(group_a)
    b = _numeric_array(group_b)
    n1, n2 = int(len(a)), int(len(b))
    if n1 < 2 or n2 < 2:
        t_stat, p_value, df = 0.0, 1.0, None
    elif scipy_stats is not None:
        test = scipy_stats.ttest_ind(a, b, equal_var=False)
        t_stat = float(test.statistic)
        p_value = float(test.pvalue)
        df = _welch_df(a, b)
    else:
        t_stat, df = _welch_t_and_df(a, b)
        p_value = 2.0 * (1.0 - NormalDist().cdf(abs(t_stat)))
    effect = _compute_cohen_d(float(t_stat), n1, n2)
    return HypothesisStatistic(
        test_name="welch_t",
        effect_size=float(effect),
        effect_size_label="Cohen's d",
        p_value=float(p_value),
        confidence_interval_low=None,
        confidence_interval_high=None,
        sample_size=n1 + n2,
        sample_sizes_per_group=(n1, n2),
        degrees_of_freedom=None if df is None else float(df),
        raw_test_statistic=float(t_stat),
    )


def _run_anova(groups: list[np.ndarray]) -> HypothesisStatistic:
    cleaned = [_numeric_array(group) for group in groups]
    sample_sizes = tuple(int(len(group)) for group in cleaned)
    if len(cleaned) < 2 or any(size < 2 for size in sample_sizes):
        f_stat, p_value = 0.0, 1.0
    elif scipy_stats is not None:
        f_stat, p_value = scipy_stats.f_oneway(*cleaned)
    else:
        f_stat = _anova_f_stat(cleaned)
        p_value = math.exp(-0.5 * max(float(f_stat), 0.0))
    df_between = float(max(len(cleaned) - 1, 0))
    df_within = float(max(sum(sample_sizes) - len(cleaned), 0))
    eta_squared = _compute_eta_squared(float(f_stat), df_between, df_within)
    return HypothesisStatistic(
        test_name="anova_oneway",
        effect_size=float(eta_squared),
        effect_size_label="eta-squared",
        p_value=float(p_value),
        confidence_interval_low=None,
        confidence_interval_high=None,
        sample_size=sum(sample_sizes),
        sample_sizes_per_group=sample_sizes,
        degrees_of_freedom=df_between,
        raw_test_statistic=float(f_stat),
    )


def _run_chi_square(contingency_table: np.ndarray) -> HypothesisStatistic:
    arr = np.asarray(contingency_table)
    if not np.allclose(arr, np.round(arr)):
        raise ValueError(f"contingency table has non-integer counts: {arr}")
    table = arr.astype(np.int64)
    n = int(table.sum())
    if table.size == 0 or n == 0:
        chi2, p_value, df, cramer_v = 0.0, 1.0, 0.0, 0.0
    elif scipy_stats is not None:
        chi2, p_value, df, _expected = scipy_stats.chi2_contingency(table)
        if scipy_association is not None:
            cramer_v = float(scipy_association(table, method="cramer"))
        else:
            cramer_v = _manual_cramer_v(float(chi2), table)
    else:
        chi2 = _manual_chi_square(table)
        df = float((table.shape[0] - 1) * (table.shape[1] - 1))
        p_value = math.exp(-0.5 * max(float(chi2), 0.0))
        cramer_v = _manual_cramer_v(float(chi2), table)
    group_sizes = tuple(int(value) for value in table.sum(axis=0).tolist())
    return HypothesisStatistic(
        test_name="chi_square_cramer_v",
        effect_size=float(cramer_v),
        effect_size_label="Cramér's V",
        p_value=float(p_value),
        confidence_interval_low=None,
        confidence_interval_high=None,
        sample_size=n,
        sample_sizes_per_group=group_sizes,
        degrees_of_freedom=float(df),
        raw_test_statistic=float(chi2),
    )


def _compute_pearson_ci(
    r: float,
    n: int,
    alpha: float = 0.05,
) -> tuple[float, float]:
    """Fisher z-transform confidence interval for Pearson/Spearman r."""

    if n <= 3 or abs(r) >= 1.0:
        return (float(r), float(r))
    z = math.atanh(r)
    se = 1.0 / math.sqrt(n - 3)
    critical = NormalDist().inv_cdf(1.0 - alpha / 2.0)
    return (math.tanh(z - critical * se), math.tanh(z + critical * se))


def _compute_cohen_d(t: float, n1: int, n2: int) -> float:
    """Convert a Welch t statistic into Cohen's d."""

    if n1 <= 0 or n2 <= 0:
        return 0.0
    return float(t) * math.sqrt(1.0 / n1 + 1.0 / n2)


def _compute_eta_squared(f: float, df_between: float, df_within: float) -> float:
    if not math.isfinite(f):
        return 1.0
    numerator = f * df_between
    denominator = numerator + df_within
    if denominator == 0:
        return 0.0
    return float(numerator / denominator)


def _check_sample_size_floor(sizes: list[int], floor: int = _MIN_SAMPLE_SIZE) -> bool:
    return bool(sizes) and all(int(size) >= floor for size in sizes)


def _verdict_from_statistic(
    statistic: HypothesisStatistic,
    hypothesized_direction: str,
) -> tuple[str, str]:
    """Apply significance, effect-size, and direction thresholds."""

    threshold = _EFFECT_THRESHOLDS.get(statistic.test_name)
    if threshold is None:
        return "INCONCLUSIVE", "unsupported_test"
    if not math.isfinite(statistic.p_value) or statistic.p_value >= 0.05:
        return "INCONCLUSIVE", "not_statistically_significant"
    if abs(statistic.effect_size) < threshold:
        return "INCONCLUSIVE", "effect_below_threshold"

    sign = 1 if statistic.effect_size > 0 else -1 if statistic.effect_size < 0 else 0
    if hypothesized_direction in {"different"}:
        return "CONFIRMED", "effect_meets_threshold"
    if hypothesized_direction in {"higher", "correlated_positive"}:
        if sign > 0:
            return "CONFIRMED", "effect_meets_threshold"
        if sign < 0:
            return "REFUTED", "direction_opposite"
    if hypothesized_direction in {"lower", "correlated_negative"}:
        if sign < 0:
            return "CONFIRMED", "effect_meets_threshold"
        if sign > 0:
            return "REFUTED", "direction_opposite"
    return "INCONCLUSIVE", "direction_not_testable"


def parse_freetext_hypothesis(text: str, schema: SurveySchema) -> HypothesisSpec | None:
    api_key = os.environ.get("PORTKEY_API_KEY", "").strip()
    if not api_key or OpenAI is None:
        return None
    questions = [
        {
            "question_id": question.canonical_id,
            "question_text": question.question_text,
            "question_type": question.question_type.value,
        }
        for question in schema.questions
    ]
    system_prompt = (
        "Parse the analyst's hypothesis into a JSON HypothesisSpec matching this "
        "schema: {title, outcome_question_id, predictor_question_id, "
        "outcome_direction, predictor_groups, free_text}. Use only question IDs "
        "from the provided schema. Return ONLY JSON. If unsure, return the empty "
        "object {}."
    )
    try:
        client = OpenAI(api_key=api_key, base_url=PORTKEY_BASE_URL)
        response = client.chat.completions.create(
            model=PORTKEY_DEFAULT_MODEL,
            temperature=0.1,
            messages=[
                {"role": "system", "content": system_prompt},
                {
                    "role": "user",
                    "content": json.dumps(
                        {"hypothesis": text, "questions": questions},
                        ensure_ascii=False,
                    ),
                },
            ],
        )
        content = response.choices[0].message.content or "{}"
        payload = json.loads(content)
        if not payload:
            return None
        groups = tuple(
            HypothesisGroup(
                label=str(group["label"]),
                matching_values=tuple(group.get("matching_values", ())),
            )
            for group in payload.get("predictor_groups", ()) or ()
        )
        return HypothesisSpec(
            title=str(payload["title"]),
            outcome_question_id=str(payload["outcome_question_id"]),
            predictor_question_id=str(payload["predictor_question_id"]),
            outcome_direction=str(payload["outcome_direction"]),
            predictor_groups=groups,
            free_text=text,
        )
    except Exception:
        return None


def explain_verdict(result: HypothesisResult) -> str:
    api_key = os.environ.get("PORTKEY_API_KEY", "").strip()
    if not api_key or OpenAI is None or result.statistic is None:
        return ""
    system_prompt = (
        f"Write 1-2 sentences explaining why the verdict is {result.verdict}. "
        "Use ONLY the numbers in the provided statistic object. Do NOT invent "
        "values. Do NOT imply causality. Do NOT round or change numbers - quote "
        "them as-is."
    )
    try:
        client = OpenAI(api_key=api_key, base_url=PORTKEY_BASE_URL)
        response = client.chat.completions.create(
            model=PORTKEY_DEFAULT_MODEL,
            temperature=0.1,
            messages=[
                {"role": "system", "content": system_prompt},
                {
                    "role": "user",
                    "content": json.dumps(
                        {
                            "spec": asdict(result.spec),
                            "verdict": result.verdict,
                            "verdict_reason": result.verdict_reason,
                            "statistic": asdict(result.statistic),
                        },
                        ensure_ascii=False,
                    ),
                },
            ],
        )
        return (response.choices[0].message.content or "").strip()
    except Exception:
        return ""


def _series_for_question(
    df: pd.DataFrame,
    question: QuestionSpec,
    role: str,
) -> pd.Series:
    if question.question_type is QuestionType.MULTI_SELECT_BINARY:
        present_columns = [column for column in question.raw_columns if column in df.columns]
        if not present_columns:
            return pd.Series([np.nan] * len(df), index=df.index)
        selected = df[present_columns].apply(
            lambda row: any(_selected_value(value) for value in row),
            axis=1,
        )
        return selected.map(lambda value: "Selected" if value else "Not selected")

    if question.question_type is QuestionType.GRID_RATED:
        present_columns = [column for column in question.raw_columns if column in df.columns]
        if not present_columns:
            return pd.Series([np.nan] * len(df), index=df.index)
        numeric = df[present_columns].apply(pd.to_numeric, errors="coerce")
        return numeric.mean(axis=1, skipna=True)

    if question.question_type is QuestionType.NUMERIC_ALLOCATION:
        present_columns = [column for column in question.raw_columns if column in df.columns]
        if not present_columns:
            return pd.Series([np.nan] * len(df), index=df.index)
        numeric = df[present_columns].apply(pd.to_numeric, errors="coerce")
        return numeric.mean(axis=1, skipna=True)

    source_column = question.raw_columns[0] if question.raw_columns else question.canonical_id
    if source_column not in df.columns:
        return pd.Series([np.nan] * len(df), index=df.index)
    series = df[source_column]
    if role in {"continuous"}:
        return pd.to_numeric(series, errors="coerce")
    if role == "ordinal":
        return series.map(lambda value: _ordinal_value(value, question))
    return series.map(lambda value: _categorical_value(value, question))


def _ordinal_value(value: Any, question: QuestionSpec) -> float | None:
    if pd.isna(value):
        return None
    mapping = question.label_to_numeric_value or {}
    if isinstance(value, str) and value in mapping:
        return float(mapping[value])
    candidates: list[Any] = [value]
    if isinstance(value, str):
        stripped = value.strip()
        candidates.append(stripped)
        try:
            numeric = float(stripped)
        except ValueError:
            numeric = None
        if numeric is not None and numeric.is_integer():
            candidates.append(int(numeric))
    if isinstance(value, float) and value.is_integer():
        candidates.append(int(value))
    for candidate in candidates:
        label = question.option_map.get(candidate)
        if label in mapping:
            return float(mapping[label])
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _categorical_value(value: Any, question: QuestionSpec) -> Any:
    if pd.isna(value):
        return None
    candidates: list[Any] = [value]
    if isinstance(value, str):
        stripped = value.strip()
        candidates.append(stripped)
        try:
            numeric = float(stripped)
        except ValueError:
            numeric = None
        if numeric is not None and numeric.is_integer():
            candidates.append(int(numeric))
    if isinstance(value, float) and value.is_integer():
        candidates.append(int(value))
    for candidate in candidates:
        if candidate in question.option_map:
            return question.option_map[candidate]
    return value


def _apply_predictor_groups(series: pd.Series, spec: HypothesisSpec) -> pd.Series:
    if not spec.predictor_groups:
        return series
    group_lookup: list[tuple[str, set[Any], set[str]]] = [
        (group.label, set(group.matching_values), {str(value) for value in group.matching_values})
        for group in spec.predictor_groups
    ]

    def resolve(value: Any) -> str | None:
        for label, raw_values, string_values in group_lookup:
            if value in raw_values or str(value) in string_values:
                return label
        return None

    return series.map(resolve)


def _selected_value(value: Any) -> bool:
    if pd.isna(value):
        return False
    if isinstance(value, str):
        stripped = value.strip().lower()
        if stripped in {"", "0", "false", "no", "not selected", "unchecked"}:
            return False
        return True
    try:
        return float(value) != 0.0
    except (TypeError, ValueError):
        return bool(value)


def _numeric_array(values: Any) -> np.ndarray:
    return pd.to_numeric(pd.Series(values), errors="coerce").dropna().to_numpy(dtype=float)


def _normal_approx_correlation_p_value(r: float, n: int) -> float:
    if n <= 3 or abs(r) >= 1:
        return 0.0 if abs(r) >= 1 else 1.0
    z = abs(math.atanh(r) * math.sqrt(n - 3))
    return 2.0 * (1.0 - NormalDist().cdf(z))


def _welch_t_and_df(a: np.ndarray, b: np.ndarray) -> tuple[float, float]:
    var_a = float(np.var(a, ddof=1))
    var_b = float(np.var(b, ddof=1))
    n1, n2 = len(a), len(b)
    se2 = var_a / n1 + var_b / n2
    if se2 <= 0:
        mean_diff = float(np.mean(a)) - float(np.mean(b))
        if mean_diff == 0:
            return 0.0, float(n1 + n2 - 2)
        return math.copysign(1e12, mean_diff), float(n1 + n2 - 2)
    t_stat = (float(np.mean(a)) - float(np.mean(b))) / math.sqrt(se2)
    return t_stat, _welch_df(a, b)


def _welch_df(a: np.ndarray, b: np.ndarray) -> float:
    var_a = float(np.var(a, ddof=1))
    var_b = float(np.var(b, ddof=1))
    n1, n2 = len(a), len(b)
    numerator = (var_a / n1 + var_b / n2) ** 2
    denominator = ((var_a / n1) ** 2 / (n1 - 1)) + ((var_b / n2) ** 2 / (n2 - 1))
    return float(numerator / denominator) if denominator else float(n1 + n2 - 2)


def _anova_f_stat(groups: list[np.ndarray]) -> float:
    all_values = np.concatenate(groups)
    grand_mean = float(np.mean(all_values))
    ss_between = sum(len(group) * (float(np.mean(group)) - grand_mean) ** 2 for group in groups)
    ss_within = sum(float(((group - float(np.mean(group))) ** 2).sum()) for group in groups)
    df_between = len(groups) - 1
    df_within = len(all_values) - len(groups)
    if df_between <= 0 or df_within <= 0:
        return 0.0
    if ss_within <= 0:
        return 1e12 if ss_between > 0 else 0.0
    return (ss_between / df_between) / (ss_within / df_within)


def _manual_chi_square(table: np.ndarray) -> float:
    total = table.sum()
    if total <= 0:
        return 0.0
    expected = np.outer(table.sum(axis=1), table.sum(axis=0)) / total
    with np.errstate(divide="ignore", invalid="ignore"):
        chi = (table - expected) ** 2 / expected
    return float(np.nan_to_num(chi).sum())


def _manual_cramer_v(chi2: float, table: np.ndarray) -> float:
    n = float(table.sum())
    min_dim = min(table.shape) - 1
    if n <= 0 or min_dim <= 0:
        return 0.0
    return math.sqrt(max(chi2, 0.0) / (n * min_dim))


def _record_hypothesis_metric(
    calculation_log: CalculationLog,
    spec: HypothesisSpec,
    metric_name: str,
    source_columns: tuple[str, ...],
    active_filters_summary: str,
    formula: str,
    value: float,
    valid_n: int,
    numerator: float | int | None = None,
    denominator: float | int | None = None,
) -> str:
    entry_id = f"hypothesis:{uuid4().hex}:{metric_name}"
    calculation_log.record(
        AuditRecord(
            output_sheet="Hypothesis_Check",
            metric_name=entry_id,
            source_question_id=f"{spec.outcome_question_id},{spec.predictor_question_id}",
            source_columns=source_columns or ("<none>",),
            filter_expr=active_filters_summary,
            numerator=numerator,
            denominator=denominator,
            formula=formula,
            value_raw=float(value) if math.isfinite(float(value)) else 0.0,
            valid_n=int(valid_n),
            missing_n=0,
            timestamp=datetime.now(timezone.utc),
        )
    )
    return entry_id


def _record_statistic_metrics(
    calculation_log: CalculationLog,
    spec: HypothesisSpec,
    statistic: HypothesisStatistic,
    source_columns: tuple[str, ...],
    active_filters_summary: str,
) -> list[str]:
    metric_values = [
        ("effect_size", statistic.effect_size),
        ("p_value", statistic.p_value),
    ]
    if statistic.confidence_interval_low is not None:
        metric_values.append(("ci_low", statistic.confidence_interval_low))
    if statistic.confidence_interval_high is not None:
        metric_values.append(("ci_high", statistic.confidence_interval_high))
    return [
        _record_hypothesis_metric(
            calculation_log,
            spec,
            metric_name,
            source_columns,
            active_filters_summary,
            f"scipy.stats.{statistic.test_name}(outcome, predictor)",
            value,
            valid_n=statistic.sample_size,
        )
        for metric_name, value in metric_values
    ]


def _source_columns_for_questions(
    outcome_question: QuestionSpec | None,
    predictor_question: QuestionSpec | None,
) -> tuple[str, ...]:
    columns: list[str] = []
    if outcome_question is not None:
        columns.extend(outcome_question.raw_columns)
    if predictor_question is not None:
        columns.extend(predictor_question.raw_columns)
    return tuple(dict.fromkeys(columns)) or ("<none>",)


def _lookup_cross_cut_id(
    lookup: dict[Any, str] | None,
    outcome_question_id: str,
    predictor_question_id: str,
) -> str | None:
    if not lookup:
        return None
    keys = [
        (outcome_question_id, predictor_question_id),
        (predictor_question_id, outcome_question_id),
        f"{outcome_question_id}|{predictor_question_id}",
        f"{predictor_question_id}|{outcome_question_id}",
    ]
    for key in keys:
        if key in lookup:
            return lookup[key]
    return None
