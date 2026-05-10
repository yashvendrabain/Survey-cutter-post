"""Tests for Stage D deterministic outcome-aware suggestion scoring."""

from __future__ import annotations

import unittest

from src.cross_cut_suggestions import CrossCutSuggestion, score_suggestions_for_outcome
from src.models import (
    AnalysisType,
    CrossCutSpec,
    DenominatorPolicy,
    DifferentiatorResult,
    OutcomeSegmentationResult,
    QuestionType,
    SegmentDefinition,
    WinnerProfile,
)


def _spec(cross_cut_id: str, source_question_ids: tuple[str, ...]) -> CrossCutSpec:
    return CrossCutSpec(
        cross_cut_id=cross_cut_id,
        title=cross_cut_id,
        analysis_type=AnalysisType.CROSS_TAB,
        source_question_ids=source_question_ids,
    )


def _suggestion(cross_cut_id: str, source_question_ids: tuple[str, ...]) -> CrossCutSuggestion:
    return CrossCutSuggestion(
        spec=_spec(cross_cut_id, source_question_ids),
        reason=f"Reason for {cross_cut_id}",
        rule_score=80,
    )


def _diff(question_id: str, cramers_v: float = 0.2) -> DifferentiatorResult:
    return DifferentiatorResult(
        question_id=question_id,
        question_text=f"Question {question_id}",
        question_type=QuestionType.SINGLE_SELECT.value,
        cramers_v=cramers_v,
        top_option_label="Top option",
        top_option_winner_rate=0.6,
        top_option_loser_rate=0.3,
        top_option_lift=2.0,
        winner_n=50,
        loser_n=50,
        p_value=0.01,
    )


def _segmentation_result(
    differentiators: tuple[DifferentiatorResult, ...],
    outcome_question_id: str = "Q_OUT",
) -> OutcomeSegmentationResult:
    definition = SegmentDefinition(
        outcome_question_id=outcome_question_id,
        segment_mode="categorical",
        winner_values=(1,),
    )
    profile = WinnerProfile(
        outcome_question_id=outcome_question_id,
        winner_label="Winner",
        winner_n=50,
        loser_n=50,
        defining_traits=(),
    )
    return OutcomeSegmentationResult(
        outcome_question_id=outcome_question_id,
        segment_definition=definition,
        winner_n=50,
        loser_n=50,
        total_n=100,
        differentiators=differentiators,
        winner_profile=profile,
        skipped_questions=(),
    )


class TestStageD(unittest.TestCase):
    def test_score_suggestions_for_outcome_with_empty_differentiators(self) -> None:
        suggestions = [
            _suggestion("CT_Q1_Q2", ("Q1", "Q2")),
            _suggestion("CT_Q3_Q4", ("Q3", "Q4")),
        ]

        scored = score_suggestions_for_outcome(
            suggestions,
            _segmentation_result(()),
        )

        self.assertEqual([item.outcome_relevance_score for item in scored], [0.1, 0.1])

    def test_score_suggestions_for_outcome_with_matching_differentiators(self) -> None:
        suggestion = _suggestion("CT_Q1_Q2", ("Q1", "Q2"))

        scored = score_suggestions_for_outcome(
            [suggestion],
            _segmentation_result((_diff("Q1", cramers_v=0.35),)),
        )

        self.assertGreaterEqual(scored[0].outcome_relevance_score, 0.5)

    def test_score_suggestions_for_outcome_sorting(self) -> None:
        suggestions = [
            _suggestion("CT_Q5_Q6", ("Q5", "Q6")),
            _suggestion("CT_Q1_Q2", ("Q1", "Q2")),
            _suggestion("CT_Q3_Q4", ("Q3", "Q4")),
        ]

        scored = score_suggestions_for_outcome(
            suggestions,
            _segmentation_result(
                (
                    _diff("Q1", cramers_v=0.4),
                    _diff("Q2", cramers_v=0.2),
                    _diff("Q3", cramers_v=0.2),
                )
            ),
        )

        scores = [item.outcome_relevance_score for item in scored]
        self.assertEqual(scores, sorted(scores, reverse=True))
        self.assertEqual(scored[0].spec.cross_cut_id, "CT_Q1_Q2")

    def test_score_suggestions_for_outcome_all_match_bonus(self) -> None:
        suggestion = _suggestion("CT_Q1_Q2", ("Q1", "Q2"))

        scored = score_suggestions_for_outcome(
            [suggestion],
            _segmentation_result(
                (
                    _diff("Q1", cramers_v=0.35),
                    _diff("Q2", cramers_v=0.2),
                )
            ),
        )

        self.assertGreaterEqual(scored[0].outcome_relevance_score, 0.95)

    def test_score_suggestions_for_outcome_with_outcome_question_id_match(self) -> None:
        suggestion = _suggestion("CT_Q_OUT_Q1", ("Q_OUT", "Q1"))

        scored = score_suggestions_for_outcome(
            [suggestion],
            _segmentation_result((_diff("Q1", cramers_v=0.35),)),
        )

        self.assertEqual(scored[0].outcome_relevance_score, 0.85)


if __name__ == "__main__":
    unittest.main()
