"""Test the think-cell table formatter and inline chart renderer."""

from __future__ import annotations

import unittest

from src.chart_recommender import recommend_chart
from src.models import (
    DenominatorPolicy,
    GridBinaryPivotResult,
    GridBinaryPivotRow,
    QuestionType,
    RankOrderResult,
    RankOrderRow,
)
from src.thinkcell_table_formatter import format_for_thinkcell


class TestThinkCellTableFormatter(unittest.TestCase):
    def test_single_select_tsv_has_tab_separated_rows(self):
        from tests.test_chart_recommender import make_single_select

        result = make_single_select(3)
        rec = recommend_chart(result)
        payload = format_for_thinkcell(
            result, rec, question_text="Test Q", survey_name="Test"
        )
        tsv = payload.to_tsv()
        self.assertIn("\t", tsv)
        self.assertIn("Test", payload.source_line)
        self.assertIn("N=100", payload.source_line)


    def test_rank_order_payload_has_rank_rows_and_option_columns(self):
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
                    option_label="Option A",
                    counts_per_rank=[50, 30, 20],
                    pcts_per_rank=[0.5, 0.3, 0.2],
                ),
                RankOrderRow(
                    option_id="B",
                    option_label="Option B",
                    counts_per_rank=[30, 40, 30],
                    pcts_per_rank=[0.3, 0.4, 0.3],
                ),
            ],
            total_respondents=100,
            total_responses=200,
        )
        rec = recommend_chart(result)
        payload = format_for_thinkcell(
            result, rec, question_text="Rank Question", survey_name="Test"
        )
        self.assertEqual(payload.headers, ["Option A", "Option B"])
        self.assertEqual(len(payload.rows), 3)
        self.assertEqual(payload.rows[0][0], "Rank 1")
        self.assertEqual(payload.rows[0][1], 50)
        self.assertEqual(payload.rows[0][2], 30)


    def test_grid_binary_pivot_payload_is_matrix_layout(self):
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
                ),
                GridBinaryPivotRow(
                    row_id="r2",
                    row_label="IT/Engineering",
                    counts_per_column=[109, 435],
                    pcts_per_column=[0.18, 0.73],
                ),
            ],
            total_respondents=300,
            total_responses=773,
        )
        rec = recommend_chart(result)
        payload = format_for_thinkcell(
            result, rec, question_text="Roles", survey_name="Test"
        )
        self.assertEqual(payload.headers, ["Blocked", "Scored"])
        self.assertEqual(len(payload.rows), 2)
        self.assertEqual(payload.rows[0][0], "Future users")


class TestChartRenderer(unittest.TestCase):
    def test_render_chart_returns_plotly_figure_when_plotly_is_available(self):
        try:
            import plotly.graph_objects as go
        except ModuleNotFoundError:
            self.skipTest("Plotly is not installed in this local runtime.")

        from src.chart_renderer import render_chart
        from tests.test_chart_recommender import make_single_select

        result = make_single_select(3)
        rec = recommend_chart(result)
        payload = format_for_thinkcell(
            result, rec, question_text="Test Q", survey_name="Test"
        )
        fig = render_chart(rec, payload)
        self.assertIsInstance(fig, go.Figure)
