"""Test the chart recommendation rules engine."""

from __future__ import annotations

import unittest

from src.chart_recommender import ChartType, recommend_chart
from src.models import (
    DenominatorPolicy,
    GridBinaryPivotResult,
    GridBinaryPivotRow,
    GridRatedResult,
    GridRatedRow,
    MultiSelectResult,
    QuestionType,
    RankOrderResult,
    RankOrderRow,
    SingleSelectResult,
)


def make_single_select(n_options: int = 4) -> SingleSelectResult:
    return SingleSelectResult(
        question_id="Q1",
        question_type=QuestionType.SINGLE_SELECT,
        valid_n=100,
        missing_n=0,
        denominator_policy=DenominatorPolicy.VALID_RESPONSES,
        distribution={
            index: {"label": f"Option {index}", "count": 25, "rate": 0.25}
            for index in range(n_options)
        },
    )


class TestChartRecommender(unittest.TestCase):
    def test_single_select_returns_column_stacked(self):
        result = make_single_select(4)
        rec = recommend_chart(result)
        self.assertEqual(rec.chart_type, ChartType.COLUMN_STACKED)
        self.assertEqual(rec.orientation, "vertical")

    def test_multi_select_returns_bar_clustered_horizontal(self):
        result = MultiSelectResult(
            question_id="Q2",
            question_type=QuestionType.MULTI_SELECT_BINARY,
            valid_n=100,
            missing_n=0,
            denominator_policy=DenominatorPolicy.VALID_RESPONSES,
            selections={
                f"opt{index}": {
                    "label": f"Opt {index}",
                    "count": 50,
                    "selection_rate": 0.5,
                }
                for index in range(5)
            },
            respondents_who_answered_any=100,
        )
        rec = recommend_chart(result)
        self.assertEqual(rec.chart_type, ChartType.BAR_CLUSTERED)
        self.assertEqual(rec.orientation, "horizontal")
        self.assertEqual(rec.sort_order, "descending")

    def test_rank_order_returns_bar_stacked(self):
        result = RankOrderResult(
            question_id="Q43",
            question_type=QuestionType.RANK_ORDER,
            valid_n=200,
            missing_n=0,
            denominator_policy=DenominatorPolicy.VALID_RESPONSES,
            question_text="Rank",
            K=3,
            rows=[
                RankOrderRow(
                    option_id="A",
                    option_label="A",
                    counts_per_rank=[50, 30, 20],
                    pcts_per_rank=[0.5, 0.3, 0.2],
                )
            ],
            total_respondents=100,
            total_responses=100,
        )
        rec = recommend_chart(result)
        self.assertEqual(rec.chart_type, ChartType.BAR_STACKED)
        self.assertEqual(rec.orientation, "horizontal")
        self.assertEqual(rec.sort_order, "rank_1_descending")

    def test_grid_rated_2_entities_returns_bar_clustered(self):
        result = GridRatedResult(
            question_id="Q30",
            question_type=QuestionType.GRID_RATED,
            valid_n=200,
            missing_n=0,
            denominator_policy=DenominatorPolicy.VALID_RESPONSES,
            question_text="Rate",
            column_headers=["Winner", "Other"],
            rows=[
                GridRatedRow(
                    row_id="r1",
                    row_label="Criterion 1",
                    means_per_column=[8.5, 7.2],
                    valid_n_per_column=[150, 140],
                    delta=1.3,
                )
            ],
            total_respondents=200,
            total_responses=290,
            show_delta=True,
        )
        rec = recommend_chart(result)
        self.assertEqual(rec.chart_type, ChartType.BAR_CLUSTERED)
        self.assertEqual(rec.sort_order, "delta_descending")
        self.assertIs(rec.show_delta, True)

    def test_grid_rated_3_entities_returns_line(self):
        result = GridRatedResult(
            question_id="Q30",
            question_type=QuestionType.GRID_RATED,
            valid_n=200,
            missing_n=0,
            denominator_policy=DenominatorPolicy.VALID_RESPONSES,
            question_text="Rate",
            column_headers=["A", "B", "C"],
            rows=[
                GridRatedRow(
                    row_id="r1",
                    row_label="Criterion 1",
                    means_per_column=[8.5, 7.2, 6.0],
                    valid_n_per_column=[150, 140, 130],
                    delta=None,
                )
            ],
            total_respondents=200,
            total_responses=420,
            show_delta=False,
        )
        rec = recommend_chart(result)
        self.assertEqual(rec.chart_type, ChartType.LINE)

    def test_grid_binary_pivot_returns_heatmap_table(self):
        result = GridBinaryPivotResult(
            question_id="Q26",
            question_type=QuestionType.GRID_BINARY_SELECT,
            valid_n=300,
            missing_n=0,
            denominator_policy=DenominatorPolicy.VALID_RESPONSES,
            question_text="Roles",
            column_headers=["Blocked", "Scored"],
            rows=[
                GridBinaryPivotRow(
                    row_id="r1",
                    row_label="Future users",
                    counts_per_column=[23, 206],
                    pcts_per_column=[0.08, 0.68],
                )
            ],
            total_respondents=300,
            total_responses=229,
        )
        rec = recommend_chart(result)
        self.assertEqual(rec.chart_type, ChartType.HEATMAP_TABLE)
        self.assertEqual(rec.artifact_type, "formatted_table")
        self.assertEqual(rec.button_label_override, "Generate formatted table")
