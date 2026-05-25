"""Tests for rank-order single-cut calculations."""

from __future__ import annotations

import unittest

import pandas as pd

from src.calculation_log import CalculationLog
from src.models import DenominatorPolicy, QuestionSpec, QuestionType, RankOrderResult
from src.single_cut.rank_order import compute_rank_order


def make_rank_spec(value_range: tuple[int, int] = (1, 2)) -> QuestionSpec:
    return QuestionSpec(
        question_id="[Q22]",
        canonical_id="Q22",
        question_text="Rank the top options",
        question_type=QuestionType.RANK_ORDER,
        raw_columns=("Q22r1", "Q22r2", "Q22r3"),
        option_map={
            "Q22r1": "Option 1",
            "Q22r2": "Option 2",
            "Q22r3": "Option 3",
        },
        value_range=value_range,
        denominator_policy=DenominatorPolicy.VALID_RESPONSES,
    )


class TestRankOrder(unittest.TestCase):
    def test_rank_order_basic(self) -> None:
        df = pd.DataFrame(
            {
                "Q22r1": [1, 1, 1, 2, None],
                "Q22r2": [2, 2, None, 1, 2],
                "Q22r3": [None, None, None, None, None],
            }
        )
        log = CalculationLog()

        result = compute_rank_order(make_rank_spec(), df, log)

        self.assertIsInstance(result, RankOrderResult)
        self.assertEqual(result.total_respondents, 5)
        self.assertEqual(result.total_responses, 8)
        by_id = {row.option_id: row for row in result.rows}
        self.assertEqual(by_id["Q22r1"].counts_per_rank, [3, 1])
        self.assertEqual(by_id["Q22r2"].counts_per_rank, [1, 3])
        self.assertEqual(by_id["Q22r3"].counts_per_rank, [0, 0])
        self.assertEqual(len(log), 6)

    def test_rank_order_with_filter(self) -> None:
        df = pd.DataFrame(
            {
                "Q22r1": [1, 2, 1],
                "Q22r2": [2, 1, None],
                "Q22r3": [None, None, 2],
            }
        )
        mask = pd.Series([True, False, True])

        result = compute_rank_order(make_rank_spec(), df, CalculationLog(), mask)

        by_id = {row.option_id: row for row in result.rows}
        self.assertEqual(result.total_respondents, 2)
        self.assertEqual(by_id["Q22r1"].counts_per_rank, [2, 0])
        self.assertEqual(by_id["Q22r3"].counts_per_rank, [0, 1])

    def test_rank_order_k_from_value_range(self) -> None:
        df = pd.DataFrame({"Q22r1": [5], "Q22r2": [1], "Q22r3": [None]})

        result = compute_rank_order(make_rank_spec((1, 5)), df, CalculationLog())

        self.assertEqual(result.K, 5)
        self.assertEqual(result.rows[0].counts_per_rank, [0, 0, 0, 0, 1])

    def test_rank_order_handles_sparse_blanks(self) -> None:
        df = pd.DataFrame({"Q22r1": [None, 1], "Q22r2": [None, None], "Q22r3": [None, None]})

        result = compute_rank_order(make_rank_spec(), df, CalculationLog())

        self.assertEqual(result.total_respondents, 1)
        self.assertEqual(result.missing_n, 1)

    def test_rank_order_total_responses_matches_non_null_count(self) -> None:
        df = pd.DataFrame({"Q22r1": [1, 2], "Q22r2": [2, None], "Q22r3": [None, 1]})

        result = compute_rank_order(make_rank_spec(), df, CalculationLog())

        self.assertEqual(
            sum(sum(row.counts_per_rank) for row in result.rows),
            result.total_responses,
        )

    def test_applies_conditional_on_filter(self) -> None:
        spec = QuestionSpec(
            question_id="[Q22]",
            canonical_id="Q22",
            question_text="Rank the top options",
            question_type=QuestionType.RANK_ORDER,
            raw_columns=("Q22r1", "Q22r2", "Q22r3"),
            option_map={
                "Q22r1": "Option 1",
                "Q22r2": "Option 2",
                "Q22r3": "Option 3",
            },
            value_range=(1, 2),
            denominator_policy=DenominatorPolicy.VALID_RESPONSES,
            conditional_on="Q_GATE",
        )
        df = pd.DataFrame(
            {
                "Q_GATE": ["Y", "Y", "Y", "Y", "Y", None, None, None, None, None],
                "Q22r1": [1] * 10,
                "Q22r2": [2] * 10,
                "Q22r3": [None] * 10,
            }
        )

        result = compute_rank_order(spec, df, CalculationLog())

        self.assertEqual(result.total_respondents, 5)
        by_id = {row.option_id: row for row in result.rows}
        self.assertEqual(by_id["Q22r1"].counts_per_rank, [5, 0])


if __name__ == "__main__":
    unittest.main()
