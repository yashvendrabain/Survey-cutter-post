"""Tests for rule-based cross-cut suggestions."""

from __future__ import annotations

from datetime import datetime, timezone
import unittest

from src.cross_cut_suggestions import suggest_cross_cuts
from src.models import (
    AnalysisType,
    CrossCutSpec,
    QuestionSpec,
    QuestionType,
    SurveySchema,
)


TYPE_SCORES = {
    AnalysisType.EXPECTED_VS_REALIZED: 95,
    AnalysisType.CROSS_TAB: 90,
    AnalysisType.GROUP_COMPARISON: 85,
    AnalysisType.SEGMENT_PROFILE: 80,
}


def q(
    canonical_id: str,
    question_text: str,
    question_type: QuestionType,
    option_map: dict[int | str, str] | None = None,
) -> QuestionSpec:
    return QuestionSpec(
        question_id=f"[{canonical_id}]",
        canonical_id=canonical_id,
        question_text=question_text,
        question_type=question_type,
        raw_columns=(canonical_id,),
        option_map=option_map or {},
    )


def schema(*questions: QuestionSpec) -> SurveySchema:
    return SurveySchema(
        questions=tuple(questions),
        respondent_id_column="record",
        total_respondents=1,
        source_datamap_path="datamap.xlsx",
        source_rawdata_path="raw.csv",
        parsed_at=datetime(2026, 5, 5, tzinfo=timezone.utc),
    )


def industry_question() -> QuestionSpec:
    return q(
        "Q_IND",
        "Which industry best describes your company?",
        QuestionType.SINGLE_SELECT,
        {1: "Tech", 2: "Retail", 3: "Healthcare"},
    )


def region_question() -> QuestionSpec:
    return q(
        "Q_REGION",
        "What region is your headquarters in?",
        QuestionType.SINGLE_SELECT,
        {1: "North", 2: "South"},
    )


def target_question(canonical_id: str = "Q_TGT") -> QuestionSpec:
    return q(
        canonical_id,
        "Which product capability matters most?",
        QuestionType.SINGLE_SELECT,
        {1: "Speed", 2: "Quality", 3: "Cost", 4: "Support"},
    )


def metric_question(canonical_id: str = "Q_NUM") -> QuestionSpec:
    return q(
        canonical_id,
        "Projected annual pipeline revenue growth",
        QuestionType.DIRECT_NUMERIC,
    )


def expected_question() -> QuestionSpec:
    return q(
        "Q_EXPECTED",
        "Expected quarterly pipeline revenue growth next year",
        QuestionType.DIRECT_NUMERIC,
    )


def realized_question() -> QuestionSpec:
    return q(
        "Q_REALIZED",
        "Actual quarterly pipeline revenue growth next year",
        QuestionType.DIRECT_NUMERIC,
    )


class TestCrossCutSuggestions(unittest.TestCase):
    def test_no_demographic_questions_returns_empty_or_other_rules(self) -> None:
        suggestions = suggest_cross_cuts(
            schema(expected_question(), realized_question(), target_question())
        )

        self.assertEqual(len(suggestions), 1)
        self.assertEqual(
            suggestions[0][0].analysis_type,
            AnalysisType.EXPECTED_VS_REALIZED,
        )

    def test_segment_profile_suggested_for_industry_question(self) -> None:
        suggestions = suggest_cross_cuts(schema(industry_question(), target_question()))

        specs = [spec for spec, _ in suggestions]
        self.assertEqual(len(specs), 1)
        self.assertEqual(specs[0].analysis_type, AnalysisType.SEGMENT_PROFILE)
        self.assertEqual(specs[0].source_question_ids, ("Q_IND", "Q_TGT"))
        self.assertEqual(specs[0].filter_expr, "Q_IND == 1")

    def test_cross_tab_suggested_for_two_demographics(self) -> None:
        suggestions = suggest_cross_cuts(
            schema(industry_question(), region_question(), target_question())
        )

        cross_tabs = [
            spec for spec, _ in suggestions
            if spec.analysis_type is AnalysisType.CROSS_TAB
        ]
        self.assertEqual(len(cross_tabs), 1)
        self.assertEqual(cross_tabs[0].source_question_ids, ("Q_IND", "Q_REGION"))

    def test_group_comparison_suggested_for_demographic_x_numeric(self) -> None:
        suggestions = suggest_cross_cuts(schema(industry_question(), metric_question()))

        specs = [spec for spec, _ in suggestions]
        self.assertEqual(len(specs), 1)
        self.assertEqual(specs[0].analysis_type, AnalysisType.GROUP_COMPARISON)
        self.assertEqual(specs[0].source_question_ids, ("Q_IND", "Q_NUM"))

    def test_expected_vs_realized_keyword_detection(self) -> None:
        suggestions = suggest_cross_cuts(schema(expected_question(), realized_question()))

        self.assertEqual(len(suggestions), 1)
        spec, reason = suggestions[0]
        self.assertEqual(spec.analysis_type, AnalysisType.EXPECTED_VS_REALIZED)
        self.assertEqual(spec.source_question_ids, ("Q_EXPECTED", "Q_REALIZED"))
        self.assertIn("Gap between expected and realized", reason)

    def test_suggestions_capped_at_max(self) -> None:
        questions = [industry_question()] + [
            target_question(f"Q_TGT_{index}") for index in range(6)
        ]

        suggestions = suggest_cross_cuts(schema(*questions), max_suggestions=2)

        self.assertEqual(len(suggestions), 2)

    def test_dedup_by_source_question_pair(self) -> None:
        suggestions = suggest_cross_cuts(schema(industry_question(), target_question()))

        keys = [
            (spec.analysis_type, frozenset(spec.source_question_ids))
            for spec, _ in suggestions
        ]
        self.assertEqual(len(keys), len(set(keys)))
        self.assertEqual(len(suggestions), 1)

    def test_all_suggestions_are_valid_cross_cut_specs(self) -> None:
        test_schema = schema(
            industry_question(),
            region_question(),
            target_question(),
            metric_question(),
            expected_question(),
            realized_question(),
        )

        suggestions = suggest_cross_cuts(test_schema)

        canonical_ids = {question.canonical_id for question in test_schema.questions}
        for spec, reason in suggestions:
            self.assertIsInstance(spec, CrossCutSpec)
            self.assertTrue(reason)
            self.assertTrue(set(spec.source_question_ids).issubset(canonical_ids))

    def test_score_ordering_descending(self) -> None:
        suggestions = suggest_cross_cuts(
            schema(
                industry_question(),
                region_question(),
                target_question(),
                metric_question(),
                expected_question(),
                realized_question(),
            )
        )

        scores = [TYPE_SCORES[spec.analysis_type] for spec, _ in suggestions]
        self.assertEqual(scores, sorted(scores, reverse=True))

    def test_empty_schema_returns_empty(self) -> None:
        self.assertEqual(suggest_cross_cuts(schema()), [])

    def test_suggestions_deterministic_across_calls(self) -> None:
        test_schema = schema(
            industry_question(),
            region_question(),
            target_question(),
            metric_question(),
            expected_question(),
            realized_question(),
        )

        first = suggest_cross_cuts(test_schema)
        second = suggest_cross_cuts(test_schema)

        self.assertEqual(first, second)


if __name__ == "__main__":
    unittest.main()
