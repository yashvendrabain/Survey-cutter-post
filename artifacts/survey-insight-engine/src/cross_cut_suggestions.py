"""Rule-based cross-cut suggestion engine."""

from __future__ import annotations

from dataclasses import dataclass, replace
import re
from typing import TypeAlias

from src.models import (
    AnalysisType,
    CrossCutSpec,
    OutcomeSegmentationResult,
    QuestionSpec,
    QuestionType,
    SurveySchema,
)


Candidate: TypeAlias = tuple[CrossCutSpec, str, int]


@dataclass(frozen=True, slots=True)
class CrossCutSuggestion:
    """Rule-based cross-cut suggestion with optional outcome relevance score."""

    spec: CrossCutSpec
    reason: str
    rule_score: int = 0
    outcome_relevance_score: float = 0.0

    @property
    def synthetic_question_title(self) -> str:
        return self.spec.title

    @property
    def business_question(self) -> str:
        return self.reason

    @property
    def source_question_ids(self) -> tuple[str, ...]:
        return self.spec.source_question_ids

    @property
    def analysis_type(self) -> AnalysisType:
        return self.spec.analysis_type

DEMOGRAPHIC_KEYWORDS = (
    "industry",
    "region",
    "country",
    "size",
    "function",
    "role",
    "seniority",
    "department",
    "segment",
    "tier",
)
EXPECTED_KEYWORDS = {"expected", "anticipated", "projected", "expect"}
REALIZED_KEYWORDS = {"realized", "actual", "achieved", "delivered"}
STOPWORDS = {
    "the",
    "a",
    "an",
    "is",
    "are",
    "of",
    "in",
    "on",
    "for",
    "to",
    "and",
    "or",
    "what",
    "your",
    "you",
}
TOKEN_PATTERN = re.compile(r"[A-Za-z0-9_]+")


def suggest_cross_cuts(
    schema: SurveySchema,
    max_suggestions: int = 15,
) -> list[tuple[CrossCutSpec, str]]:
    """Generate deterministic rule-based cross-cut suggestions."""

    if max_suggestions <= 0 or not schema.questions:
        return []

    candidates: list[Candidate] = []
    candidates.extend(_segment_profile_candidates(schema))
    candidates.extend(_demographic_cross_tab_candidates(schema))
    candidates.extend(_group_comparison_candidates(schema))
    candidates.extend(_expected_vs_realized_candidates(schema))

    deduped = _deduplicate_candidates(candidates)
    deduped.sort(key=lambda candidate: (-candidate[2], candidate[0].cross_cut_id))
    return [
        (spec, reason)
        for spec, reason, _ in deduped[:max_suggestions]
    ]


def score_suggestions_for_outcome(
    suggestions: list[CrossCutSuggestion] | list[tuple[CrossCutSpec, str]],
    segmentation_result: OutcomeSegmentationResult,
) -> list[CrossCutSuggestion]:
    """Re-score cross-cut suggestions by relevance to outcome segmentation."""

    top_differentiators = {
        differentiator.question_id: differentiator
        for differentiator in segmentation_result.differentiators[:10]
    }
    scored: list[CrossCutSuggestion] = []
    for raw_suggestion in suggestions:
        suggestion = _as_suggestion(raw_suggestion)
        source_ids = set(suggestion.source_question_ids)
        matching_ids = source_ids.intersection(top_differentiators)

        if segmentation_result.outcome_question_id in source_ids:
            score = 0.3
        else:
            score = 0.1

        if matching_ids:
            score = min(score + 0.5 * len(matching_ids), 0.9)
            if matching_ids == source_ids:
                score = 0.95
            if any(top_differentiators[qid].cramers_v > 0.3 for qid in matching_ids):
                score = min(score + 0.05, 0.99)

        scored.append(
            replace(
                suggestion,
                outcome_relevance_score=round(float(score), 4),
            )
        )

    return sorted(
        scored,
        key=lambda suggestion: (
            -suggestion.outcome_relevance_score,
            -suggestion.rule_score,
            suggestion.spec.cross_cut_id,
        ),
    )


def _segment_profile_candidates(schema: SurveySchema) -> list[Candidate]:
    candidates: list[Candidate] = []
    demographics = _demographic_questions(schema)
    targets = [
        question
        for question in _single_select_questions(schema)
        if not _is_demographic(question) and len(question.option_map) <= 10
    ]

    for demographic in demographics:
        emitted_for_demographic = 0
        for option_value, option_label in _filter_options(demographic):
            for target in targets:
                if emitted_for_demographic >= 5:
                    break
                candidate = _candidate(
                    cross_cut_id=(
                        f"SP_{demographic.canonical_id}_"
                        f"{_safe_id_part(option_value)}_{target.canonical_id}"
                    ),
                    title=f"{_question_title(target)} within {option_label}",
                    analysis_type=AnalysisType.SEGMENT_PROFILE,
                    source_question_ids=(
                        demographic.canonical_id,
                        target.canonical_id,
                    ),
                    reason=(
                        f"Profile of {target.canonical_id} within "
                        f"{option_label} segment"
                    ),
                    score=80,
                    filter_expr=f"{demographic.canonical_id} == {option_value}",
                    filter_mask_description=(
                        f"{demographic.canonical_id} = {option_label}"
                    ),
                )
                if candidate is not None:
                    candidates.append(candidate)
                    emitted_for_demographic += 1
            if emitted_for_demographic >= 5:
                break
    return candidates


def _demographic_cross_tab_candidates(schema: SurveySchema) -> list[Candidate]:
    candidates: list[Candidate] = []
    demographics = _demographic_questions(schema)
    for index, question_a in enumerate(demographics):
        for question_b in demographics[index + 1:]:
            first, second = sorted(
                (question_a, question_b),
                key=lambda question: question.canonical_id,
            )
            candidate = _candidate(
                cross_cut_id=f"CT_{first.canonical_id}_{second.canonical_id}",
                title=f"{_question_title(first)} x {_question_title(second)}",
                analysis_type=AnalysisType.CROSS_TAB,
                source_question_ids=(first.canonical_id, second.canonical_id),
                reason=(
                    f"How {first.canonical_id} respondents distribute "
                    f"across {second.canonical_id}"
                ),
                score=90,
            )
            if candidate is not None:
                candidates.append(candidate)
    return candidates


def _group_comparison_candidates(schema: SurveySchema) -> list[Candidate]:
    candidates: list[Candidate] = []
    for demographic in _demographic_questions(schema):
        for metric in _direct_numeric_questions(schema):
            candidate = _candidate(
                cross_cut_id=f"GC_{demographic.canonical_id}_{metric.canonical_id}",
                title=f"{_question_title(metric)} by {_question_title(demographic)}",
                analysis_type=AnalysisType.GROUP_COMPARISON,
                source_question_ids=(demographic.canonical_id, metric.canonical_id),
                reason=(
                    f"Average {metric.canonical_id} across "
                    f"{demographic.canonical_id} segments"
                ),
                score=85,
            )
            if candidate is not None:
                candidates.append(candidate)
    return candidates


def _expected_vs_realized_candidates(schema: SurveySchema) -> list[Candidate]:
    candidates: list[Candidate] = []
    numerics = _direct_numeric_questions(schema)
    for index, question_a in enumerate(numerics):
        for question_b in numerics[index + 1:]:
            ordered = _expected_realized_pair(question_a, question_b)
            if ordered is None:
                continue
            expected, realized = ordered
            candidate = _candidate(
                cross_cut_id=f"EVR_{expected.canonical_id}_{realized.canonical_id}",
                title=f"{_question_title(expected)} vs {_question_title(realized)}",
                analysis_type=AnalysisType.EXPECTED_VS_REALIZED,
                source_question_ids=(expected.canonical_id, realized.canonical_id),
                reason=(
                    "Gap between expected and realized "
                    f"({expected.canonical_id} vs {realized.canonical_id})"
                ),
                score=95,
            )
            if candidate is not None:
                candidates.append(candidate)
    return candidates


def _candidate(
    cross_cut_id: str,
    title: str,
    analysis_type: AnalysisType,
    source_question_ids: tuple[str, ...],
    reason: str,
    score: int,
    filter_expr: str | None = None,
    filter_mask_description: str | None = None,
) -> Candidate | None:
    try:
        spec = CrossCutSpec(
            cross_cut_id=cross_cut_id,
            title=title,
            analysis_type=analysis_type,
            source_question_ids=source_question_ids,
            filter_expr=filter_expr,
            filter_mask_description=filter_mask_description,
        )
    except ValueError:
        return None
    return spec, reason, score


def _as_suggestion(
    suggestion: CrossCutSuggestion | tuple[CrossCutSpec, str],
) -> CrossCutSuggestion:
    if isinstance(suggestion, CrossCutSuggestion):
        return suggestion
    spec, reason = suggestion
    return CrossCutSuggestion(spec=spec, reason=reason)


def _deduplicate_candidates(candidates: list[Candidate]) -> list[Candidate]:
    best_by_key: dict[tuple[AnalysisType, frozenset[str]], Candidate] = {}
    for candidate in candidates:
        spec = candidate[0]
        key = (spec.analysis_type, frozenset(spec.source_question_ids))
        existing = best_by_key.get(key)
        if existing is None or candidate[2] > existing[2]:
            best_by_key[key] = candidate
    return list(best_by_key.values())


def _demographic_questions(schema: SurveySchema) -> list[QuestionSpec]:
    return sorted(
        (
            question
            for question in schema.questions
            if _is_demographic(question)
        ),
        key=lambda question: question.canonical_id,
    )


def _single_select_questions(schema: SurveySchema) -> list[QuestionSpec]:
    return sorted(
        (
            question
            for question in schema.questions
            if question.question_type is QuestionType.SINGLE_SELECT
        ),
        key=lambda question: question.canonical_id,
    )


def _direct_numeric_questions(schema: SurveySchema) -> list[QuestionSpec]:
    return sorted(
        (
            question
            for question in schema.questions
            if question.question_type is QuestionType.DIRECT_NUMERIC
        ),
        key=lambda question: question.canonical_id,
    )


def _is_demographic(question: QuestionSpec) -> bool:
    if question.question_type in (
        QuestionType.SINGLE_SELECT,
        QuestionType.DEMOGRAPHIC_OR_SEGMENT,
    ):
        if not 2 <= len(question.option_map) <= 12:
            return False
    elif question.question_type is QuestionType.GRID_SINGLE_SELECT:
        row_count = len(question.grid_row_labels) if question.grid_row_labels else 0
        if not 2 <= row_count <= 12:
            return False
    else:
        return False
    text = question.question_text.lower()
    return any(keyword in text for keyword in DEMOGRAPHIC_KEYWORDS)


def _expected_realized_pair(
    question_a: QuestionSpec,
    question_b: QuestionSpec,
) -> tuple[QuestionSpec, QuestionSpec] | None:
    tokens_a = _tokens(question_a.question_text)
    tokens_b = _tokens(question_b.question_text)
    if len(tokens_a.intersection(tokens_b)) < 3:
        return None

    a_expected = bool(tokens_a.intersection(EXPECTED_KEYWORDS))
    b_expected = bool(tokens_b.intersection(EXPECTED_KEYWORDS))
    a_realized = bool(tokens_a.intersection(REALIZED_KEYWORDS))
    b_realized = bool(tokens_b.intersection(REALIZED_KEYWORDS))
    if a_expected and b_realized:
        return question_a, question_b
    if b_expected and a_realized:
        return question_b, question_a
    return None


def _tokens(text: str) -> set[str]:
    return {
        token
        for token in (match.group(0).lower() for match in TOKEN_PATTERN.finditer(text))
        if token not in STOPWORDS
    }


def _question_title(question: QuestionSpec) -> str:
    return f"{question.canonical_id}: {question.question_text}"


def _sorted_options(question: QuestionSpec) -> list[tuple[int | str, str]]:
    return sorted(
        question.option_map.items(),
        key=lambda item: (str(type(item[0]).__name__), str(item[0])),
    )


def _filter_options(question: QuestionSpec) -> list[tuple[int | str, str]]:
    if question.question_type is QuestionType.GRID_SINGLE_SELECT:
        return sorted(
            (question.grid_row_labels or {}).items(),
            key=lambda item: str(item[0]),
        )
    return _sorted_options(question)


def _safe_id_part(value: int | str) -> str:
    return re.sub(r"\W+", "_", str(value)).strip("_") or "value"
