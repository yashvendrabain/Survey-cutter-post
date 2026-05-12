"""Tests for deterministic outcome segmentation."""

from __future__ import annotations

import unittest
from datetime import datetime, timezone

import pandas as pd

from src.models import (
    DenominatorPolicy,
    ProfileTrait,
    QuestionSpec,
    QuestionType,
    SegmentDefinition,
    SurveySchema,
)
from src.outcome_segmentation import compute_outcome_segmentation, _lift


UTC_NOW = datetime(2026, 5, 10, 12, 0, tzinfo=timezone.utc)


def _question(
    qid: str,
    question_type: QuestionType = QuestionType.SINGLE_SELECT,
    *,
    text: str | None = None,
    raw_columns: tuple[str, ...] | None = None,
    option_map: dict | None = None,
) -> QuestionSpec:
    if raw_columns is None:
        raw_columns = (qid,) if question_type is not QuestionType.METADATA_OR_ID else ()
    if option_map is None:
        if question_type is QuestionType.MULTI_SELECT_BINARY:
            option_map = {column: column for column in raw_columns}
        else:
            option_map = {1: "One", 2: "Two", 3: "Three"}
    return QuestionSpec(
        question_id=qid,
        canonical_id=qid,
        question_text=text or f"Question {qid}",
        question_type=question_type,
        raw_columns=raw_columns,
        option_map=option_map,
        value_range=(1, 3) if question_type is QuestionType.SINGLE_SELECT else None,
        denominator_policy=DenominatorPolicy.VALID_RESPONSES,
    )


def _schema(questions: list[QuestionSpec]) -> SurveySchema:
    return SurveySchema(
        questions=tuple(questions),
        respondent_id_column="record",
        total_respondents=200,
        source_datamap_path="synthetic_map.xlsx",
        source_rawdata_path="synthetic_raw.xlsx",
        parsed_at=UTC_NOW,
    )


class TestOutcomeSegmentation(unittest.TestCase):
    def test_basic_categorical_segmentation(self) -> None:
        df = pd.DataFrame(
            {
                "Q_OUTCOME": [1] * 50 + [2] * 50,
                "Q_STRONG": [1] * 40 + [2] * 10 + [1] * 10 + [2] * 40,
                "Q_FLAT": [1, 2] * 50,
            }
        )
        schema = _schema(
            [
                _question("Q_OUTCOME"),
                _question("Q_STRONG", text="Strong differentiator"),
                _question("Q_FLAT", text="Flat differentiator"),
            ]
        )
        audit_log = []

        result = compute_outcome_segmentation(
            df,
            schema,
            "Q_OUTCOME",
            SegmentDefinition("Q_OUTCOME", "categorical", winner_values=(1,)),
            audit_log,
        )

        self.assertEqual(result.differentiators[0].question_id, "Q_STRONG")
        self.assertTrue(all(0.0 <= diff.cramers_v <= 1.0 for diff in result.differentiators))

    def test_numeric_threshold_segmentation(self) -> None:
        df = pd.DataFrame(
            {
                "Q_REVENUE": list(range(100)),
                "Q_DRIVER": [1] * 45 + [2] * 5 + [1] * 10 + [2] * 40,
            }
        )
        schema = _schema(
            [
                _question("Q_REVENUE", QuestionType.DIRECT_NUMERIC, option_map={}),
                _question("Q_DRIVER"),
            ]
        )
        audit_log = []

        result = compute_outcome_segmentation(
            df,
            schema,
            "Q_REVENUE",
            SegmentDefinition(
                "Q_REVENUE",
                "numeric_threshold",
                winner_threshold=50,
                threshold_direction="gte",
            ),
            audit_log,
        )

        self.assertEqual(result.winner_n + result.loser_n, result.total_n)
        self.assertGreater(result.differentiators[0].top_option_lift, 1.0)

    def test_quartile_segmentation_top_winner(self) -> None:
        df = pd.DataFrame(
            {
                "Q_REVENUE": list(range(100)),
                "Q_DRIVER": [1] * 25 + [2] * 50 + [1] * 25,
            }
        )
        schema = _schema(
            [
                _question("Q_REVENUE", QuestionType.DIRECT_NUMERIC, option_map={}),
                _question("Q_DRIVER"),
            ]
        )
        audit_log = []

        result = compute_outcome_segmentation(
            df,
            schema,
            "Q_REVENUE",
            SegmentDefinition(
                outcome_question_id="Q_REVENUE",
                segment_mode="quartile",
                quartile_winner="top",
            ),
            audit_log,
        )

        self.assertEqual(result.winner_n, 25)
        self.assertEqual(result.loser_n, 25)
        self.assertEqual(result.total_n, 50)
        self.assertEqual(result.segment_definition.loser_label, "Laggard")

    def test_quartile_default_label_is_laggard(self) -> None:
        sd = SegmentDefinition(
            outcome_question_id="Q1",
            segment_mode="quartile",
        )

        self.assertEqual(sd.loser_label, "Laggard")
        self.assertEqual(sd.winner_label, "Winner")

    def test_multi_select_binary_handling(self) -> None:
        df = pd.DataFrame({"Q_OUTCOME": [1] * 50 + [2] * 50})
        for idx in range(1, 6):
            df[f"Q_MULTI_{idx}"] = 0
        df.loc[:44, "Q_MULTI_3"] = 1
        df.loc[50:54, "Q_MULTI_3"] = 1
        df.loc[10:30, "Q_MULTI_1"] = 1
        df.loc[60:80, "Q_MULTI_1"] = 1
        schema = _schema(
            [
                _question("Q_OUTCOME"),
                _question(
                    "Q_MULTI",
                    QuestionType.MULTI_SELECT_BINARY,
                    raw_columns=tuple(f"Q_MULTI_{idx}" for idx in range(1, 6)),
                    option_map={f"Q_MULTI_{idx}": f"Option {idx}" for idx in range(1, 6)},
                ),
            ]
        )
        audit_log = []

        result = compute_outcome_segmentation(
            df,
            schema,
            "Q_OUTCOME",
            SegmentDefinition("Q_OUTCOME", "categorical", winner_values=(1,)),
            audit_log,
        )

        self.assertEqual(result.differentiators[0].question_id, "Q_MULTI")
        self.assertEqual(result.differentiators[0].top_option_label, "Option 3")

    def test_direct_numeric_quartile_binning(self) -> None:
        df = pd.DataFrame(
            {
                "Q_OUTCOME": [1] * 50 + [2] * 50,
                "Q_AGE": list(range(50, 100)) + list(range(18, 68)),
            }
        )
        schema = _schema(
            [
                _question("Q_OUTCOME"),
                _question("Q_AGE", QuestionType.DIRECT_NUMERIC, text="Age", option_map={}),
            ]
        )
        audit_log = []

        result = compute_outcome_segmentation(
            df,
            schema,
            "Q_OUTCOME",
            SegmentDefinition("Q_OUTCOME", "categorical", winner_values=(1,)),
            audit_log,
        )

        self.assertIn(
            result.differentiators[0].top_option_label,
            {"Q1 (bottom quartile)", "Q2", "Q3", "Q4 (top quartile)"},
        )

    def test_sample_size_validation(self) -> None:
        df = pd.DataFrame(
            {
                "Q_OUTCOME": [1] * 5 + [2] * 95,
                "Q_DRIVER": [1, 2] * 50,
            }
        )
        schema = _schema([_question("Q_OUTCOME"), _question("Q_DRIVER")])
        audit_log = []

        result = compute_outcome_segmentation(
            df,
            schema,
            "Q_OUTCOME",
            SegmentDefinition("Q_OUTCOME", "categorical", winner_values=(1,)),
            audit_log,
            min_sample_size=30,
        )

        self.assertEqual(result.differentiators, ())
        self.assertTrue(any("Winner sample size" in warning for warning in result.warnings))

    def test_lift_edge_cases(self) -> None:
        df = pd.DataFrame(
            {
                "Q_OUTCOME": [1] * 50 + [2] * 50,
                "Q_EDGE": [1] * 25 + [2] * 25 + [2] * 50,
                "Q_TWOX": [1] * 30 + [2] * 20 + [1] * 15 + [2] * 35,
            }
        )
        schema = _schema([_question("Q_OUTCOME"), _question("Q_EDGE"), _question("Q_TWOX")])
        audit_log = []

        result = compute_outcome_segmentation(
            df,
            schema,
            "Q_OUTCOME",
            SegmentDefinition("Q_OUTCOME", "categorical", winner_values=(1,)),
            audit_log,
        )
        by_id = {diff.question_id: diff for diff in result.differentiators}

        self.assertEqual(by_id["Q_EDGE"].top_option_lift, 999.0)
        self.assertIn("infinite_lift_loser_zero", by_id["Q_EDGE"].warnings)
        self.assertEqual(_lift(0.0, 0.0)[0], 1.0)
        self.assertAlmostEqual(by_id["Q_TWOX"].top_option_lift, 2.0)

    def test_winner_profile_assembly(self) -> None:
        df = pd.DataFrame({"Q_OUTCOME": [1] * 60 + [2] * 60})
        questions = [_question("Q_OUTCOME")]
        for idx in range(1, 5):
            qid = f"Q_DRIVER_{idx}"
            df[qid] = [1] * 42 + [2] * 18 + [1] * 18 + [2] * 42
            questions.append(_question(qid, text=f"Driver {idx}"))
        schema = _schema(questions)
        audit_log = []

        result = compute_outcome_segmentation(
            df,
            schema,
            "Q_OUTCOME",
            SegmentDefinition("Q_OUTCOME", "categorical", winner_values=(1,)),
            audit_log,
        )

        self.assertEqual(len(result.winner_profile.defining_traits), 4)
        lifts = [trait.lift for trait in result.winner_profile.defining_traits]
        self.assertTrue(all(lift > 1.2 for lift in lifts))

    def test_profile_trait_includes_laggard_top_option(self) -> None:
        df = pd.DataFrame(
            {
                "Q_OUTCOME": [1] * 50 + [2] * 50,
                "Q_DRIVER": [1] * 40 + [2] * 10 + [1] * 15 + [2] * 35,
            }
        )
        schema = _schema(
            [
                _question("Q_OUTCOME"),
                _question(
                    "Q_DRIVER",
                    option_map={1: "Winner-led option", 2: "Laggard-led option"},
                ),
            ]
        )
        audit_log = []

        result = compute_outcome_segmentation(
            df,
            schema,
            "Q_OUTCOME",
            SegmentDefinition("Q_OUTCOME", "categorical", winner_values=(1,)),
            audit_log,
        )

        trait = result.winner_profile.defining_traits[0]
        self.assertEqual(trait.option_label, "Winner-led option")
        self.assertEqual(trait.laggard_top_option_label, "Laggard-led option")
        self.assertGreater(
            trait.laggard_top_option_loser_rate,
            trait.laggard_top_option_winner_rate,
        )
        self.assertNotEqual(trait.laggard_top_option_label, trait.option_label)

    def test_profile_trait_laggard_fields_default_to_empty(self) -> None:
        trait = ProfileTrait(
            question_id="Q1",
            question_text="Question 1",
            option_label="Winner option",
            winner_rate=0.6,
            loser_rate=0.3,
            lift=2.0,
            rate_gap=0.3,
        )

        self.assertEqual(trait.laggard_top_option_label, "")
        self.assertEqual(trait.laggard_top_option_winner_rate, 0.0)
        self.assertEqual(trait.laggard_top_option_loser_rate, 0.0)
        self.assertEqual(trait.laggard_top_option_gap, 0.0)

    def test_skipped_questions_logging(self) -> None:
        df = pd.DataFrame(
            {
                "Q_OUTCOME": [1] * 50 + [2] * 50,
                "Q_TEXT": ["a"] * 100,
                "record": range(100),
            }
        )
        schema = _schema(
            [
                _question("Q_OUTCOME"),
                _question("Q_TEXT", QuestionType.OPEN_TEXT, raw_columns=("Q_TEXT",), option_map={}),
                _question("record", QuestionType.METADATA_OR_ID, raw_columns=(), option_map={}),
            ]
        )
        audit_log = []

        result = compute_outcome_segmentation(
            df,
            schema,
            "Q_OUTCOME",
            SegmentDefinition("Q_OUTCOME", "categorical", winner_values=(1,)),
            audit_log,
        )

        self.assertIn(("Q_TEXT", "unsupported_type"), result.skipped_questions)
        self.assertIn(("record", "unsupported_type"), result.skipped_questions)

    def test_audit_log_population(self) -> None:
        df = pd.DataFrame(
            {
                "Q_OUTCOME": [1] * 50 + [2] * 50,
                "Q1": [1] * 35 + [2] * 15 + [1] * 15 + [2] * 35,
                "Q2": [1, 2] * 50,
                "Q3": list(range(100)),
            }
        )
        schema = _schema([_question("Q_OUTCOME"), _question("Q1"), _question("Q2"), _question("Q3", QuestionType.DIRECT_NUMERIC, option_map={})])
        audit_log = []

        compute_outcome_segmentation(
            df,
            schema,
            "Q_OUTCOME",
            SegmentDefinition("Q_OUTCOME", "categorical", winner_values=(1,)),
            audit_log,
        )

        metric_names = [record.metric_name for record in audit_log]
        self.assertEqual(metric_names.count("differentiator_cramers_v"), 3)
        self.assertEqual(metric_names.count("laggard_top_option_rate"), 2)
        for record in audit_log:
            if record.metric_name == "differentiator_cramers_v":
                self.assertIn(record.source_question_id, {"Q1", "Q2", "Q3"})
                self.assertEqual(
                    record.formula,
                    "cramers_v = sqrt(chi2 / (n * min(rows-1, cols-1)))",
                )

    def test_full_pipeline_with_synthetic_ltb_data(self) -> None:
        df = pd.DataFrame({"Q_WINNER": [1] * 100 + [2] * 100})
        questions = [_question("Q_WINNER", text="Which vendor was the winner?")]
        for idx in range(1, 5):
            qid = f"Q_LTB_{idx}"
            df[qid] = [1] * 70 + [2] * 30 + [1] * 30 + [2] * 70
            questions.append(_question(qid, text=f"LTB driver {idx}"))
        schema = _schema(questions)
        audit_log = []

        result = compute_outcome_segmentation(
            df,
            schema,
            "Q_WINNER",
            SegmentDefinition("Q_WINNER", "categorical", winner_values=(1,)),
            audit_log,
        )

        self.assertTrue(result.differentiators)
        self.assertTrue(result.winner_profile.defining_traits)
        missing_n = len(df) - result.total_n
        self.assertEqual(len(df), result.winner_n + result.loser_n + missing_n)


if __name__ == "__main__":
    unittest.main()
