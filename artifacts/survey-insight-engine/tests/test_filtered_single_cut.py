"""Tests for filtered single-cut dispatch."""

from __future__ import annotations

from datetime import datetime, timezone
import unittest

import pandas as pd

from src.calculation_log import CalculationLog
from src.models import (
    AnalysisType,
    DenominatorPolicy,
    FilteredSingleCutResult,
    FilterSpec,
    QuestionSpec,
    QuestionType,
    SingleCutResult,
    SurveySchema,
)
from src.filtered_single_cut import compute_filtered_single_cut
from tests.conftest import GOLDEN_30_RESPONDENTS_PATH


def load_filtered_golden() -> pd.DataFrame:
    dataframe = pd.read_csv(GOLDEN_30_RESPONDENTS_PATH)
    dataframe["Q_REGION"] = [1] * 10 + [2] * 10 + [3] * 10
    dataframe["Q_INDUSTRY"] = [1] * 15 + [2] * 15
    return dataframe


def make_schema() -> SurveySchema:
    questions = (
        QuestionSpec(
            question_id="[Q_REGION]",
            canonical_id="Q_REGION",
            question_text="Which region are you based in?",
            question_type=QuestionType.SINGLE_SELECT,
            raw_columns=("Q_REGION",),
            option_map={1: "APAC", 2: "EMEA", 3: "Americas"},
            denominator_policy=DenominatorPolicy.VALID_RESPONSES,
            is_demographic=True,
        ),
        QuestionSpec(
            question_id="[Q_INDUSTRY]",
            canonical_id="Q_INDUSTRY",
            question_text="Which industry best describes your company?",
            question_type=QuestionType.SINGLE_SELECT,
            raw_columns=("Q_INDUSTRY",),
            option_map={1: "Technology", 2: "Financial services"},
            denominator_policy=DenominatorPolicy.VALID_RESPONSES,
            is_demographic=True,
        ),
        QuestionSpec(
            question_id="[Q_SS_1]",
            canonical_id="Q_SS_1",
            question_text="Yes or no target",
            question_type=QuestionType.SINGLE_SELECT,
            raw_columns=("Q_SS_1",),
            option_map={1: "Yes", 2: "No"},
            denominator_policy=DenominatorPolicy.VALID_RESPONSES,
        ),
        QuestionSpec(
            question_id="[Q_MS_1]",
            canonical_id="Q_MS_1",
            question_text="Multi-select target",
            question_type=QuestionType.MULTI_SELECT_BINARY,
            raw_columns=("Q_MS_1r1", "Q_MS_1r2", "Q_MS_1r3"),
            option_map={
                "Q_MS_1r1": "Option 1",
                "Q_MS_1r2": "Option 2",
                "Q_MS_1r3": "Option 3",
            },
            denominator_policy=DenominatorPolicy.VALID_RESPONSES,
        ),
        QuestionSpec(
            question_id="[Q_NUM_1]",
            canonical_id="Q_NUM_1",
            question_text="Numeric target",
            question_type=QuestionType.DIRECT_NUMERIC,
            raw_columns=("Q_NUM_1",),
            option_map={},
            denominator_policy=DenominatorPolicy.VALID_RESPONSES,
        ),
        QuestionSpec(
            question_id="[Q_GRID_1]",
            canonical_id="Q_GRID_1",
            question_text="Grid target",
            question_type=QuestionType.GRID_SINGLE_SELECT,
            raw_columns=("Q_GRID_1r1", "Q_GRID_1r2", "Q_GRID_1r3"),
            option_map={
                1: "Strongly disagree",
                2: "Disagree",
                3: "Agree",
                4: "Strongly agree",
            },
            value_range=(1, 4),
            denominator_policy=DenominatorPolicy.VALID_RESPONSES,
            grid_row_labels={
                "Q_GRID_1r1": "Row 1",
                "Q_GRID_1r2": "Row 2",
                "Q_GRID_1r3": "Row 3",
            },
        ),
    )
    return SurveySchema(
        questions=questions,
        respondent_id_column="respondent_id",
        total_respondents=30,
        source_datamap_path="filtered_datamap.xlsx",
        source_rawdata_path="filtered_raw.csv",
        parsed_at=datetime(2026, 5, 6, tzinfo=timezone.utc),
    )


class TestFilteredSingleCut(unittest.TestCase):
    def test_filter_with_value_runs_filtered_single_select(self) -> None:
        dataframe = load_filtered_golden()
        result = compute_filtered_single_cut(
            "Q_SS_1",
            [FilterSpec("Q_REGION", 1)],
            make_schema(),
            dataframe,
            CalculationLog(),
        )

        self.assertEqual(result.dispatch_mode, "single_cut_filtered")
        self.assertEqual(result.filtered_n, 10)
        self.assertIsNotNone(result.single_cut_result)
        self.assertEqual(result.single_cut_result.valid_n, 10)
        self.assertEqual(result.single_cut_result.distribution[1]["count"], 10)

    def test_filter_with_value_runs_filtered_numeric(self) -> None:
        dataframe = load_filtered_golden()
        result = compute_filtered_single_cut(
            "Q_NUM_1",
            [FilterSpec("Q_REGION", 1)],
            make_schema(),
            dataframe,
            CalculationLog(),
        )

        self.assertEqual(result.filtered_n, 10)
        self.assertIsNotNone(result.single_cut_result)
        self.assertEqual(result.single_cut_result.mean, 5.5)

    def test_filter_with_value_runs_filtered_multi_select(self) -> None:
        dataframe = load_filtered_golden()
        result = compute_filtered_single_cut(
            "Q_MS_1",
            [FilterSpec("Q_REGION", 1)],
            make_schema(),
            dataframe,
            CalculationLog(),
        )

        self.assertEqual(result.filtered_n, 10)
        self.assertIsNotNone(result.single_cut_result)
        self.assertEqual(result.single_cut_result.respondents_who_answered_any, 10)
        self.assertEqual(result.single_cut_result.selections["Q_MS_1r1"]["count"], 10)

    def test_filter_with_value_runs_filtered_grid(self) -> None:
        dataframe = load_filtered_golden()
        result = compute_filtered_single_cut(
            "Q_GRID_1",
            [FilterSpec("Q_REGION", 1)],
            make_schema(),
            dataframe,
            CalculationLog(),
        )

        self.assertEqual(result.filtered_n, 10)
        self.assertIsNotNone(result.single_cut_result)
        self.assertEqual(result.single_cut_result.valid_n, 10)
        self.assertEqual(result.single_cut_result.rows["Q_GRID_1r1"].valid_n, 10)

    def test_filter_without_value_dispatches_to_cross_tab(self) -> None:
        dataframe = load_filtered_golden()
        result = compute_filtered_single_cut(
            "Q_SS_1",
            [FilterSpec("Q_REGION")],
            make_schema(),
            dataframe,
            CalculationLog(),
        )

        self.assertEqual(result.dispatch_mode, "cross_cut_breakdown")
        self.assertEqual(result.filtered_n, 30)
        self.assertIsNotNone(result.cross_cut_result)
        self.assertIs(result.cross_cut_result.analysis_type, AnalysisType.CROSS_TAB)
        self.assertEqual(result.cross_cut_result.result_table["grand_total"], 30)

    def test_filter_without_value_dispatches_to_group_comparison(self) -> None:
        dataframe = load_filtered_golden()
        result = compute_filtered_single_cut(
            "Q_NUM_1",
            [FilterSpec("Q_REGION")],
            make_schema(),
            dataframe,
            CalculationLog(),
        )

        self.assertEqual(result.dispatch_mode, "cross_cut_breakdown")
        self.assertEqual(result.filtered_n, 30)
        self.assertIsNotNone(result.cross_cut_result)
        self.assertIs(
            result.cross_cut_result.analysis_type,
            AnalysisType.GROUP_COMPARISON,
        )
        self.assertEqual(
            result.cross_cut_result.result_table["per_segment"][1]["mean"],
            5.5,
        )

    def test_multiple_value_filters_combine_with_and(self) -> None:
        dataframe = load_filtered_golden()
        result = compute_filtered_single_cut(
            "Q_SS_1",
            [FilterSpec("Q_REGION", 2), FilterSpec("Q_INDUSTRY", 1)],
            make_schema(),
            dataframe,
            CalculationLog(),
        )

        self.assertEqual(result.filtered_n, 5)
        self.assertIsNotNone(result.single_cut_result)
        self.assertEqual(result.single_cut_result.valid_n, 5)

    def test_value_filter_plus_breakdown_combines_correctly(self) -> None:
        dataframe = load_filtered_golden()
        result = compute_filtered_single_cut(
            "Q_NUM_1",
            [FilterSpec("Q_REGION", 1), FilterSpec("Q_INDUSTRY")],
            make_schema(),
            dataframe,
            CalculationLog(),
        )

        self.assertEqual(result.dispatch_mode, "cross_cut_breakdown")
        self.assertEqual(result.filtered_n, 10)
        self.assertIsNotNone(result.cross_cut_result)
        per_segment = result.cross_cut_result.result_table["per_segment"]
        self.assertEqual(tuple(per_segment.keys()), (1,))
        self.assertEqual(per_segment[1]["n"], 10)

    def test_low_sample_warning_emitted_when_filtered_n_below_30(self) -> None:
        dataframe = load_filtered_golden()
        result = compute_filtered_single_cut(
            "Q_SS_1",
            [FilterSpec("Q_REGION", 1)],
            make_schema(),
            dataframe,
            CalculationLog(),
        )

        self.assertIn("Filtered sample size 10", result.warnings[0])

    def test_two_breakdown_filters_raise_value_error(self) -> None:
        dataframe = load_filtered_golden()
        with self.assertRaisesRegex(ValueError, "at most one breakdown"):
            compute_filtered_single_cut(
                "Q_SS_1",
                [FilterSpec("Q_REGION"), FilterSpec("Q_INDUSTRY")],
                make_schema(),
                dataframe,
                CalculationLog(),
            )

    def test_invalid_target_question_raises(self) -> None:
        dataframe = load_filtered_golden()
        with self.assertRaisesRegex(ValueError, "target question"):
            compute_filtered_single_cut(
                "Q_DOES_NOT_EXIST",
                [FilterSpec("Q_REGION", 1)],
                make_schema(),
                dataframe,
                CalculationLog(),
            )

    def test_invalid_filter_column_raises(self) -> None:
        dataframe = load_filtered_golden()
        with self.assertRaisesRegex(ValueError, "filter column"):
            compute_filtered_single_cut(
                "Q_SS_1",
                [FilterSpec("Q_MISSING", 1)],
                make_schema(),
                dataframe,
                CalculationLog(),
            )

    def test_audit_records_preserved(self) -> None:
        dataframe = load_filtered_golden()
        log = CalculationLog()
        result = compute_filtered_single_cut(
            "Q_SS_1",
            [FilterSpec("Q_REGION", 1)],
            make_schema(),
            dataframe,
            log,
        )

        self.assertIsNotNone(result.single_cut_result)
        self.assertEqual(result.audit_records, result.single_cut_result.audit_records)
        self.assertGreater(len(result.audit_records), 0)
        self.assertEqual(len(log), len(result.audit_records))

    def test_filtered_single_cut_result_validation(self) -> None:
        base_result = SingleCutResult(
            question_id="Q1",
            question_type=QuestionType.SINGLE_SELECT,
            valid_n=1,
            missing_n=0,
            denominator_policy=DenominatorPolicy.VALID_RESPONSES,
        )

        with self.assertRaisesRegex(ValueError, "dispatch_mode"):
            FilteredSingleCutResult(
                target_question_id="Q1",
                filters_applied=(),
                dispatch_mode="bad",
                single_cut_result=base_result,
                cross_cut_result=None,
                filtered_n=1,
            )
        with self.assertRaisesRegex(ValueError, "single_cut_result required"):
            FilteredSingleCutResult(
                target_question_id="Q1",
                filters_applied=(),
                dispatch_mode="single_cut_filtered",
                single_cut_result=None,
                cross_cut_result=None,
                filtered_n=1,
            )
        with self.assertRaisesRegex(ValueError, "filtered_n"):
            FilteredSingleCutResult(
                target_question_id="Q1",
                filters_applied=(),
                dispatch_mode="single_cut_filtered",
                single_cut_result=base_result,
                cross_cut_result=None,
                filtered_n=-1,
            )


if __name__ == "__main__":
    unittest.main()
