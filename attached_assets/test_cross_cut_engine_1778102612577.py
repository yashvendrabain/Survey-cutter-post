"""Tests for the cross-cut engine."""

from __future__ import annotations

from datetime import datetime, timezone
import unittest

import pandas as pd

from src.calculation_log import CalculationLog
from src.cross_cut_engine import (
    _compute_group_comparison,
    _compute_cross_tab,
    _compute_segment_profile,
    compute_cross_cuts,
)
from src.models import (
    AnalysisType,
    CrossCutSpec,
    DenominatorPolicy,
    QuestionSpec,
    QuestionType,
    SurveySchema,
)
from tests.conftest import CROSS_CUT_30_RESPONDENTS_PATH


EXPECTED_CROSS_TAB_COUNTS = {
    1: {1: 5, 2: 5, 3: 0, 4: 0},
    2: {1: 3, 2: 3, 3: 2, 4: 2},
    3: {1: 0, 2: 0, 3: 5, 4: 5},
}
EXPECTED_CROSS_TAB_ROW_PCT = {
    1: {1: 0.5, 2: 0.5, 3: 0.0, 4: 0.0},
    2: {1: 0.3, 2: 0.3, 3: 0.2, 4: 0.2},
    3: {1: 0.0, 2: 0.0, 3: 0.5, 4: 0.5},
}
EXPECTED_CROSS_TAB_COL_PCT = {
    1: {1: 0.625, 2: 0.625, 3: 0.0, 4: 0.0},
    2: {1: 0.375, 2: 0.375, 3: 0.2857142857142857, 4: 0.2857142857142857},
    3: {1: 0.0, 2: 0.0, 3: 0.7142857142857143, 4: 0.7142857142857143},
}
EXPECTED_CROSS_TAB_ROW_TOTALS = {1: 10, 2: 10, 3: 10}
EXPECTED_CROSS_TAB_COLUMN_TOTALS = {1: 8, 2: 8, 3: 7, 4: 7}

EXPECTED_SEGMENT_MEANS = {
    1: 55.0,
    2: 50.0,
    3: 5.5,
}
EXPECTED_GROUP_OVERALL_MEAN = 36.833333333333336
EXPECTED_GROUP_OVERALL_STD = 32.92319519137223

EXPECTED_PAIRED_N = 12
EXPECTED_EXPECTED_MEAN = 70.0
EXPECTED_REALIZED_MEAN = 60.0
EXPECTED_GAP_MEAN = -5.0
EXPECTED_GAP_STD = 0.0


def load_cross_cut_golden() -> pd.DataFrame:
    return pd.read_csv(CROSS_CUT_30_RESPONDENTS_PATH)


def make_schema() -> SurveySchema:
    questions = (
        QuestionSpec(
            question_id="[Q_SEG_1]",
            canonical_id="Q_SEG_1",
            question_text="Segment",
            question_type=QuestionType.SINGLE_SELECT,
            raw_columns=("Q_SEG_1",),
            option_map={1: "Segment 1", 2: "Segment 2", 3: "Segment 3"},
            denominator_policy=DenominatorPolicy.VALID_RESPONSES,
        ),
        QuestionSpec(
            question_id="[Q_TGT_1]",
            canonical_id="Q_TGT_1",
            question_text="Target categorical",
            question_type=QuestionType.SINGLE_SELECT,
            raw_columns=("Q_TGT_1",),
            option_map={1: "A", 2: "B", 3: "C", 4: "D"},
            denominator_policy=DenominatorPolicy.VALID_RESPONSES,
        ),
        QuestionSpec(
            question_id="[Q_NUM_3]",
            canonical_id="Q_NUM_3",
            question_text="Numeric metric",
            question_type=QuestionType.DIRECT_NUMERIC,
            raw_columns=("Q_NUM_3",),
            option_map={},
            denominator_policy=DenominatorPolicy.VALID_RESPONSES,
        ),
        QuestionSpec(
            question_id="[Q_EXP_1]",
            canonical_id="Q_EXP_1",
            question_text="Expected",
            question_type=QuestionType.DIRECT_NUMERIC,
            raw_columns=("Q_EXP_1",),
            option_map={},
            denominator_policy=DenominatorPolicy.VALID_RESPONSES,
        ),
        QuestionSpec(
            question_id="[Q_REAL_1]",
            canonical_id="Q_REAL_1",
            question_text="Realized",
            question_type=QuestionType.DIRECT_NUMERIC,
            raw_columns=("Q_REAL_1",),
            option_map={},
            denominator_policy=DenominatorPolicy.VALID_RESPONSES,
        ),
    )
    return SurveySchema(
        questions=questions,
        respondent_id_column="respondent_id",
        total_respondents=30,
        source_datamap_path="cross_datamap.xlsx",
        source_rawdata_path="cross_raw.csv",
        parsed_at=datetime(2026, 5, 5, tzinfo=timezone.utc),
    )


def make_schema_with_extra(*extra_questions: QuestionSpec) -> SurveySchema:
    base = make_schema()
    return SurveySchema(
        questions=base.questions + tuple(extra_questions),
        respondent_id_column=base.respondent_id_column,
        total_respondents=base.total_respondents,
        source_datamap_path=base.source_datamap_path,
        source_rawdata_path=base.source_rawdata_path,
        parsed_at=base.parsed_at,
    )


def grid_single_select_dimension_spec() -> QuestionSpec:
    return QuestionSpec(
        question_id="[Q_GRID_SEG]",
        canonical_id="Q_GRID_SEG",
        question_text="Grid-like segment",
        question_type=QuestionType.GRID_SINGLE_SELECT,
        raw_columns=("Q_GRID_SEGr1", "Q_GRID_SEGr2", "Q_GRID_SEGr3"),
        option_map={1: "Selected"},
        value_range=(0, 1),
        grid_row_labels={
            "Q_GRID_SEGr1": "Segment 1",
            "Q_GRID_SEGr2": "Segment 2",
            "Q_GRID_SEGr3": "Segment 3",
        },
    )


def dataframe_with_grid_segment() -> pd.DataFrame:
    dataframe = load_cross_cut_golden()
    dataframe["Q_GRID_SEGr1"] = [1] * 10 + [0] * 20
    dataframe["Q_GRID_SEGr2"] = [0] * 10 + [1] * 10 + [0] * 10
    dataframe["Q_GRID_SEGr3"] = [0] * 20 + [1] * 10
    return dataframe


def cross_tab_spec(cross_cut_id: str = "CC_TAB") -> CrossCutSpec:
    return CrossCutSpec(
        cross_cut_id=cross_cut_id,
        title="Segment by target",
        analysis_type=AnalysisType.CROSS_TAB,
        source_question_ids=("Q_SEG_1", "Q_TGT_1"),
    )


def segment_profile_spec(
    target_id: str = "Q_TGT_1",
    filter_expr: str = "Q_SEG_1 == 1",
) -> CrossCutSpec:
    return CrossCutSpec(
        cross_cut_id=f"CC_SEG_{target_id}",
        title="Segment profile",
        analysis_type=AnalysisType.SEGMENT_PROFILE,
        source_question_ids=("Q_SEG_1", target_id),
        filter_expr=filter_expr,
    )


def group_comparison_spec() -> CrossCutSpec:
    return CrossCutSpec(
        cross_cut_id="CC_GROUP",
        title="Group comparison",
        analysis_type=AnalysisType.GROUP_COMPARISON,
        source_question_ids=("Q_SEG_1", "Q_NUM_3"),
    )


def expected_vs_realized_spec() -> CrossCutSpec:
    return CrossCutSpec(
        cross_cut_id="CC_EVR",
        title="Expected vs realized",
        analysis_type=AnalysisType.EXPECTED_VS_REALIZED,
        source_question_ids=("Q_EXP_1", "Q_REAL_1"),
    )


class TestCrossCutEngine(unittest.TestCase):
    def test_cross_tab_basic_shape(self) -> None:
        dataframe = load_cross_cut_golden()
        log = CalculationLog()

        results, skips = compute_cross_cuts(
            [cross_tab_spec()],
            make_schema(),
            dataframe,
            log,
        )

        self.assertEqual(skips, [])
        self.assertEqual(len(results), 1)
        result_table = results[0].result_table
        self.assertEqual(result_table["row_question_id"], "Q_SEG_1")
        self.assertEqual(result_table["column_question_id"], "Q_TGT_1")
        self.assertEqual(result_table["grand_total"], 30)

    def test_cross_tab_counts_match_pandas_crosstab(self) -> None:
        dataframe = load_cross_cut_golden()
        log = CalculationLog()

        results, _ = compute_cross_cuts(
            [cross_tab_spec()],
            make_schema(),
            dataframe,
            log,
        )

        self.assertEqual(results[0].result_table["counts"], EXPECTED_CROSS_TAB_COUNTS)
        self.assertEqual(results[0].result_table["row_totals"], EXPECTED_CROSS_TAB_ROW_TOTALS)
        self.assertEqual(
            results[0].result_table["column_totals"],
            EXPECTED_CROSS_TAB_COLUMN_TOTALS,
        )

    def test_cross_tab_row_pct_correct(self) -> None:
        dataframe = load_cross_cut_golden()
        log = CalculationLog()

        results, _ = compute_cross_cuts(
            [cross_tab_spec()],
            make_schema(),
            dataframe,
            log,
        )

        self.assertEqual(results[0].result_table["row_pct"], EXPECTED_CROSS_TAB_ROW_PCT)

    def test_cross_tab_col_pct_correct(self) -> None:
        dataframe = load_cross_cut_golden()
        log = CalculationLog()

        results, _ = compute_cross_cuts(
            [cross_tab_spec()],
            make_schema(),
            dataframe,
            log,
        )

        self.assertEqual(
            results[0].result_table["column_pct"],
            EXPECTED_CROSS_TAB_COL_PCT,
        )

    def test_cross_tab_handles_missing_values(self) -> None:
        dataframe = load_cross_cut_golden()
        dataframe.loc[0, "Q_SEG_1"] = pd.NA
        dataframe.loc[1, "Q_TGT_1"] = pd.NA
        log = CalculationLog()

        results, skips = compute_cross_cuts(
            [cross_tab_spec()],
            make_schema(),
            dataframe,
            log,
        )

        self.assertEqual(skips, [])
        self.assertEqual(results[0].result_table["grand_total"], 28)
        self.assertEqual(log.all_records()[0].missing_n, 2)

    def test_cross_tab_raises_on_non_categorical(self) -> None:
        dataframe = load_cross_cut_golden()
        log = CalculationLog()
        spec = CrossCutSpec(
            cross_cut_id="CC_BAD",
            title="Bad cross tab",
            analysis_type=AnalysisType.CROSS_TAB,
            source_question_ids=("Q_SEG_1", "Q_NUM_3"),
        )

        with self.assertRaises(ValueError):
            _compute_cross_tab(spec, make_schema(), dataframe, log)

    def test_cross_tab_rejects_multi_select_binary(self) -> None:
        dataframe = load_cross_cut_golden()
        log = CalculationLog()
        multi_spec = QuestionSpec(
            question_id="[Q_MS]",
            canonical_id="Q_MS",
            question_text="Multi select",
            question_type=QuestionType.MULTI_SELECT_BINARY,
            raw_columns=("Q_MSr1", "Q_MSr2"),
            option_map={"Q_MSr1": "First", "Q_MSr2": "Second"},
            value_range=(0, 1),
        )
        spec = CrossCutSpec(
            cross_cut_id="CC_MULTI",
            title="Multi cross tab",
            analysis_type=AnalysisType.CROSS_TAB,
            source_question_ids=("Q_MS", "Q_SEG_1"),
        )

        with self.assertRaisesRegex(
            ValueError,
            "CROSS_TAB does not yet support MULTI_SELECT_BINARY",
        ):
            _compute_cross_tab(spec, make_schema_with_extra(multi_spec), dataframe, log)

    def test_cross_tab_with_grid_single_select_on_rows(self) -> None:
        dataframe = dataframe_with_grid_segment()
        log = CalculationLog()
        spec = CrossCutSpec(
            cross_cut_id="CC_GRID_TAB",
            title="Grid dimension by target",
            analysis_type=AnalysisType.CROSS_TAB,
            source_question_ids=("Q_GRID_SEG", "Q_TGT_1"),
        )

        results, skips = compute_cross_cuts(
            [spec],
            make_schema_with_extra(grid_single_select_dimension_spec()),
            dataframe,
            log,
        )

        self.assertEqual(skips, [])
        self.assertEqual(
            results[0].result_table["counts"],
            {
                "Q_GRID_SEGr1": {1: 5, 2: 5, 3: 0, 4: 0},
                "Q_GRID_SEGr2": {1: 3, 2: 3, 3: 2, 4: 2},
                "Q_GRID_SEGr3": {1: 0, 2: 0, 3: 5, 4: 5},
            },
        )
        self.assertEqual(
            results[0].result_table["row_label_map"],
            {
                "Q_GRID_SEGr1": "Segment 1",
                "Q_GRID_SEGr2": "Segment 2",
                "Q_GRID_SEGr3": "Segment 3",
            },
        )

    def test_cross_tab_with_grid_on_columns(self) -> None:
        dataframe = dataframe_with_grid_segment()
        log = CalculationLog()
        spec = CrossCutSpec(
            cross_cut_id="CC_GRID_COLUMNS",
            title="Target by grid dimension",
            analysis_type=AnalysisType.CROSS_TAB,
            source_question_ids=("Q_TGT_1", "Q_GRID_SEG"),
        )

        results, skips = compute_cross_cuts(
            [spec],
            make_schema_with_extra(grid_single_select_dimension_spec()),
            dataframe,
            log,
        )

        self.assertEqual(skips, [])
        self.assertEqual(
            results[0].result_table["counts"],
            {
                1: {"Q_GRID_SEGr1": 5, "Q_GRID_SEGr2": 3, "Q_GRID_SEGr3": 0},
                2: {"Q_GRID_SEGr1": 5, "Q_GRID_SEGr2": 3, "Q_GRID_SEGr3": 0},
                3: {"Q_GRID_SEGr1": 0, "Q_GRID_SEGr2": 2, "Q_GRID_SEGr3": 5},
                4: {"Q_GRID_SEGr1": 0, "Q_GRID_SEGr2": 2, "Q_GRID_SEGr3": 5},
            },
        )
        self.assertEqual(
            results[0].result_table["column_label_map"]["Q_GRID_SEGr1"],
            "Segment 1",
        )

    def test_segment_profile_with_single_select_target(self) -> None:
        dataframe = load_cross_cut_golden()
        log = CalculationLog()

        results, skips = compute_cross_cuts(
            [segment_profile_spec()],
            make_schema(),
            dataframe,
            log,
        )

        self.assertEqual(skips, [])
        target_result = results[0].result_table["target_result"]
        self.assertEqual(results[0].result_table["filter_n"], 10)
        self.assertEqual(target_result["valid_n"], 10)
        self.assertEqual(target_result["distribution"][1]["count"], 5)
        self.assertEqual(target_result["distribution"][1]["rate"], 0.5)
        self.assertEqual(target_result["distribution"][2]["count"], 5)
        self.assertEqual(target_result["distribution"][2]["rate"], 0.5)

    def test_segment_profile_with_numeric_target(self) -> None:
        dataframe = load_cross_cut_golden()
        log = CalculationLog()

        results, skips = compute_cross_cuts(
            [segment_profile_spec(target_id="Q_NUM_3", filter_expr="Q_SEG_1 == 2")],
            make_schema(),
            dataframe,
            log,
        )

        self.assertEqual(skips, [])
        target_result = results[0].result_table["target_result"]
        self.assertEqual(target_result["valid_n"], 10)
        self.assertEqual(target_result["mean"], EXPECTED_SEGMENT_MEANS[2])

    def test_segment_profile_filter_expression_parsed(self) -> None:
        dataframe = load_cross_cut_golden()
        log = CalculationLog()

        results, _ = compute_cross_cuts(
            [segment_profile_spec(filter_expr="Q_SEG_1 == 3")],
            make_schema(),
            dataframe,
            log,
        )

        self.assertEqual(results[0].result_table["filter_n"], 10)
        distribution = results[0].result_table["target_result"]["distribution"]
        self.assertEqual(distribution[3]["count"], 5)
        self.assertEqual(distribution[4]["count"], 5)

    def test_segment_profile_raises_on_invalid_filter(self) -> None:
        dataframe = load_cross_cut_golden()
        log = CalculationLog()
        spec = segment_profile_spec(filter_expr="Q_SEG_1 > 1")

        with self.assertRaises(ValueError):
            _compute_segment_profile(spec, make_schema(), dataframe, log)

    def test_group_comparison_per_segment_means(self) -> None:
        dataframe = load_cross_cut_golden()
        log = CalculationLog()

        results, skips = compute_cross_cuts(
            [group_comparison_spec()],
            make_schema(),
            dataframe,
            log,
        )

        self.assertEqual(skips, [])
        per_segment = results[0].result_table["per_segment"]
        self.assertEqual(per_segment[1]["mean"], EXPECTED_SEGMENT_MEANS[1])
        self.assertEqual(per_segment[2]["mean"], EXPECTED_SEGMENT_MEANS[2])
        self.assertEqual(per_segment[3]["mean"], EXPECTED_SEGMENT_MEANS[3])

    def test_group_comparison_includes_overall(self) -> None:
        dataframe = load_cross_cut_golden()
        log = CalculationLog()

        results, _ = compute_cross_cuts(
            [group_comparison_spec()],
            make_schema(),
            dataframe,
            log,
        )

        overall = results[0].result_table["overall"]
        self.assertEqual(overall["n"], 30)
        self.assertEqual(overall["mean"], EXPECTED_GROUP_OVERALL_MEAN)
        self.assertEqual(overall["std"], EXPECTED_GROUP_OVERALL_STD)

    def test_group_comparison_segment_count_correct(self) -> None:
        dataframe = load_cross_cut_golden()
        log = CalculationLog()

        results, _ = compute_cross_cuts(
            [group_comparison_spec()],
            make_schema(),
            dataframe,
            log,
        )

        per_segment = results[0].result_table["per_segment"]
        self.assertEqual(len(per_segment), 3)
        self.assertEqual(per_segment[1]["n"], 10)
        self.assertEqual(per_segment[2]["n"], 10)
        self.assertEqual(per_segment[3]["n"], 10)

    def test_group_comparison_rejects_allocation_metric(self) -> None:
        dataframe = load_cross_cut_golden()
        log = CalculationLog()
        allocation_spec = QuestionSpec(
            question_id="[Q_ALLOC]",
            canonical_id="Q_ALLOC",
            question_text="Allocation",
            question_type=QuestionType.NUMERIC_ALLOCATION,
            raw_columns=("Q_ALLOCr1", "Q_ALLOCr2"),
            option_map={"Q_ALLOCr1": "First", "Q_ALLOCr2": "Second"},
            value_range=(0, 999),
        )
        spec = CrossCutSpec(
            cross_cut_id="CC_ALLOC",
            title="Allocation group comparison",
            analysis_type=AnalysisType.GROUP_COMPARISON,
            source_question_ids=("Q_SEG_1", "Q_ALLOC"),
        )

        with self.assertRaisesRegex(
            ValueError,
            "GROUP_COMPARISON does not yet support NUMERIC_ALLOCATION metrics",
        ):
            _compute_group_comparison(
                spec,
                make_schema_with_extra(allocation_spec),
                dataframe,
                log,
            )

    def test_group_comparison_with_grid_segment(self) -> None:
        dataframe = dataframe_with_grid_segment()
        log = CalculationLog()
        spec = CrossCutSpec(
            cross_cut_id="CC_GRID_GROUP",
            title="Grid segment comparison",
            analysis_type=AnalysisType.GROUP_COMPARISON,
            source_question_ids=("Q_GRID_SEG", "Q_NUM_3"),
        )

        results, skips = compute_cross_cuts(
            [spec],
            make_schema_with_extra(grid_single_select_dimension_spec()),
            dataframe,
            log,
        )

        self.assertEqual(skips, [])
        per_segment = results[0].result_table["per_segment"]
        self.assertEqual(per_segment["Q_GRID_SEGr1"]["mean"], EXPECTED_SEGMENT_MEANS[1])
        self.assertEqual(per_segment["Q_GRID_SEGr2"]["mean"], EXPECTED_SEGMENT_MEANS[2])
        self.assertEqual(per_segment["Q_GRID_SEGr3"]["mean"], EXPECTED_SEGMENT_MEANS[3])

    def test_expected_vs_realized_paired_n(self) -> None:
        dataframe = load_cross_cut_golden()
        log = CalculationLog()

        results, skips = compute_cross_cuts(
            [expected_vs_realized_spec()],
            make_schema(),
            dataframe,
            log,
        )

        self.assertEqual(skips, [])
        self.assertEqual(results[0].result_table["paired_n"], EXPECTED_PAIRED_N)

    def test_expected_vs_realized_gap_mean(self) -> None:
        dataframe = load_cross_cut_golden()
        log = CalculationLog()

        results, _ = compute_cross_cuts(
            [expected_vs_realized_spec()],
            make_schema(),
            dataframe,
            log,
        )

        gap = results[0].result_table["gap"]
        self.assertEqual(gap["mean"], EXPECTED_GAP_MEAN)
        self.assertEqual(gap["std"], EXPECTED_GAP_STD)

    def test_expected_vs_realized_handles_missing_pairs(self) -> None:
        dataframe = load_cross_cut_golden()
        log = CalculationLog()

        results, _ = compute_cross_cuts(
            [expected_vs_realized_spec()],
            make_schema(),
            dataframe,
            log,
        )

        result_table = results[0].result_table
        self.assertEqual(result_table["expected"]["valid_n"], 15)
        self.assertEqual(result_table["expected"]["mean"], EXPECTED_EXPECTED_MEAN)
        self.assertEqual(result_table["realized"]["valid_n"], 15)
        self.assertEqual(result_table["realized"]["mean"], EXPECTED_REALIZED_MEAN)
        self.assertEqual(result_table["gap"]["missing_n"], 18)

    def test_engine_catches_exception_returns_skip_record(self) -> None:
        dataframe = load_cross_cut_golden()
        log = CalculationLog()
        spec = CrossCutSpec(
            cross_cut_id="CC_BAD",
            title="Bad cross cut",
            analysis_type=AnalysisType.CROSS_TAB,
            source_question_ids=("Q_SEG_1", "Q_NUM_3"),
        )

        results, skips = compute_cross_cuts([spec], make_schema(), dataframe, log)

        self.assertEqual(results, [])
        self.assertEqual(len(skips), 1)
        self.assertEqual(skips[0].skip_reason, "cross_cut_error")
        self.assertIn("ValueError", skips[0].details)

    def test_engine_returns_results_in_spec_order(self) -> None:
        dataframe = load_cross_cut_golden()
        log = CalculationLog()
        specs = [
            cross_tab_spec("CC_FIRST"),
            group_comparison_spec(),
            expected_vs_realized_spec(),
        ]

        results, skips = compute_cross_cuts(specs, make_schema(), dataframe, log)

        self.assertEqual(skips, [])
        self.assertEqual(
            [result.cross_cut_id for result in results],
            ["CC_FIRST", "CC_GROUP", "CC_EVR"],
        )

    def test_engine_audit_records_logged_per_cross_cut(self) -> None:
        dataframe = load_cross_cut_golden()
        log = CalculationLog()

        results, skips = compute_cross_cuts(
            [
                cross_tab_spec(),
                segment_profile_spec(),
                group_comparison_spec(),
                expected_vs_realized_spec(),
            ],
            make_schema(),
            dataframe,
            log,
        )

        self.assertEqual(skips, [])
        self.assertEqual(len(results), 4)
        metric_names = [record.metric_name for record in log.all_records()]
        self.assertIn("cross_tab", metric_names)
        self.assertIn("segment_profile", metric_names)
        self.assertIn("group_comparison", metric_names)
        self.assertIn("expected_vs_realized", metric_names)

    def test_cross_cut_spec_validation(self) -> None:
        with self.assertRaises(ValueError):
            CrossCutSpec(
                cross_cut_id="",
                title="Invalid",
                analysis_type=AnalysisType.SEGMENT_PROFILE,
                source_question_ids=("Q_SEG_1", "Q_TGT_1"),
            )
        with self.assertRaises(ValueError):
            CrossCutSpec(
                cross_cut_id="CC_BAD_TAB",
                title="Bad tab",
                analysis_type=AnalysisType.CROSS_TAB,
                source_question_ids=("Q_SEG_1",),
            )
        with self.assertRaises(ValueError):
            CrossCutSpec(
                cross_cut_id="CC_BAD_EVR",
                title="Bad expected vs realized",
                analysis_type=AnalysisType.EXPECTED_VS_REALIZED,
                source_question_ids=("Q_EXP_1",),
            )

    def test_cross_cut_spec_accepts_valid_display_modes(self) -> None:
        for display_mode in ("counts", "row_pct", "col_pct", "both", "all"):
            spec = CrossCutSpec(
                cross_cut_id=f"CC_MODE_{display_mode}",
                title="Mode test",
                analysis_type=AnalysisType.CROSS_TAB,
                source_question_ids=("Q_SEG_1", "Q_TGT_1"),
                display_mode=display_mode,
            )

            self.assertEqual(spec.display_mode, display_mode)

    def test_cross_cut_spec_rejects_invalid_display_mode(self) -> None:
        with self.assertRaisesRegex(ValueError, "display_mode must be one of"):
            CrossCutSpec(
                cross_cut_id="CC_BAD_MODE",
                title="Bad mode",
                analysis_type=AnalysisType.CROSS_TAB,
                source_question_ids=("Q_SEG_1", "Q_TGT_1"),
                display_mode="invalid",  # type: ignore[arg-type]
            )

    def test_cross_tab_result_carries_display_mode(self) -> None:
        dataframe = load_cross_cut_golden()
        log = CalculationLog()
        spec = CrossCutSpec(
            cross_cut_id="CC_MODE_RESULT",
            title="Mode result",
            analysis_type=AnalysisType.CROSS_TAB,
            source_question_ids=("Q_SEG_1", "Q_TGT_1"),
            display_mode="counts",
        )

        results, skips = compute_cross_cuts([spec], make_schema(), dataframe, log)

        self.assertEqual(skips, [])
        self.assertEqual(results[0].display_mode, "counts")

    def test_segment_profile_spec_requires_filter_expr(self) -> None:
        with self.assertRaisesRegex(ValueError, "SEGMENT_PROFILE requires filter_expr"):
            CrossCutSpec(
                cross_cut_id="CC_BAD_SEGMENT",
                title="Bad segment profile",
                analysis_type=AnalysisType.SEGMENT_PROFILE,
                source_question_ids=("Q_SEG_1", "Q_TGT_1"),
            )

    def test_group_comparison_spec_requires_two_sources(self) -> None:
        with self.assertRaisesRegex(
            ValueError,
            "GROUP_COMPARISON requires exactly 2 source questions",
        ):
            CrossCutSpec(
                cross_cut_id="CC_BAD_GROUP",
                title="Bad group comparison",
                analysis_type=AnalysisType.GROUP_COMPARISON,
                source_question_ids=("Q_SEG_1",),
            )


if __name__ == "__main__":
    unittest.main()
