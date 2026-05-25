"""Tests for rated grid calculations."""

from __future__ import annotations

import unittest

import pandas as pd

from src.calculation_log import CalculationLog
from src.models import DenominatorPolicy, GridRatedResult, QuestionSpec, QuestionType
from src.single_cut.grid_rated import compute_grid_rated


def make_grid_rated_spec(
    raw_columns: tuple[str, ...] = ("Q30r1c1", "Q30r1c2"),
    option_map: dict[int | str, str] | None = None,
) -> QuestionSpec:
    labels = {column: "Pre-purchase familiarity" for column in raw_columns}
    return QuestionSpec(
        question_id="[Q30]",
        canonical_id="Q30",
        question_text="Rate each vendor from 0 to 10",
        question_type=QuestionType.GRID_RATED,
        raw_columns=raw_columns,
        option_map=option_map or {"1": "Winner", "2": "Other"},
        value_range=(1, 12),
        denominator_policy=DenominatorPolicy.VALID_RESPONSES,
        grid_row_labels=labels,
        possible_role="GRID_RATED",
    )


class TestGridRated(unittest.TestCase):
    def test_grid_rated_basic(self) -> None:
        df = pd.DataFrame({"Q30r1c1": [8, 9, 10], "Q30r1c2": [5, 6, 7]})

        result = compute_grid_rated(make_grid_rated_spec(), df, CalculationLog())

        self.assertIsInstance(result, GridRatedResult)
        self.assertEqual(result.column_headers, ["Winner", "Other"])
        self.assertEqual(result.rows[0].means_per_column, [9.0, 6.0])
        self.assertEqual(result.rows[0].delta, 3.0)

    def test_grid_rated_parses_label_strings(self) -> None:
        df = pd.DataFrame(
            {
                "Q30r1c1": ["10 (extremely high)", "8"],
                "Q30r1c2": ["0 (extremely low)", "6"],
            }
        )

        result = compute_grid_rated(make_grid_rated_spec(), df, CalculationLog())

        self.assertEqual(result.rows[0].means_per_column, [9.0, 3.0])

    def test_grid_rated_skips_missing_value_tokens(self) -> None:
        df = pd.DataFrame({"Q30r1c1": [8, "I don't know"], "Q30r1c2": [5, "N/A"]})

        result = compute_grid_rated(make_grid_rated_spec(), df, CalculationLog())

        self.assertEqual(result.rows[0].valid_n_per_column, [1, 1])
        self.assertEqual(result.total_responses, 2)

    def test_grid_rated_no_delta_for_three_columns(self) -> None:
        spec = make_grid_rated_spec(
            raw_columns=("Q30r1c1", "Q30r1c2", "Q30r1c3"),
            option_map={"1": "Winner", "2": "Other", "3": "Third"},
        )
        df = pd.DataFrame({"Q30r1c1": [8], "Q30r1c2": [6], "Q30r1c3": [7]})

        result = compute_grid_rated(spec, df, CalculationLog())

        self.assertFalse(result.show_delta)
        self.assertIsNone(result.rows[0].delta)

    def test_grid_rated_valid_n_per_column(self) -> None:
        df = pd.DataFrame({"Q30r1c1": [8, None, 10], "Q30r1c2": [5, 6, None]})

        result = compute_grid_rated(make_grid_rated_spec(), df, CalculationLog())

        self.assertEqual(result.rows[0].valid_n_per_column, [2, 2])


if __name__ == "__main__":
    unittest.main()
