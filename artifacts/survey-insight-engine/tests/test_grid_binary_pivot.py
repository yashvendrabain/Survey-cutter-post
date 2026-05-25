"""Tests for binary pivot grid calculations."""

from __future__ import annotations

import unittest

import pandas as pd

from src.calculation_log import CalculationLog
from src.models import (
    DenominatorPolicy,
    GridBinaryPivotResult,
    QuestionSpec,
    QuestionType,
)
from src.single_cut.grid_binary_pivot import compute_grid_binary_pivot


def make_binary_grid_spec() -> QuestionSpec:
    raw_columns = (
        "Q26r1c1",
        "Q26r1c2",
        "Q26r2c1",
        "Q26r2c2",
        "Q26r3c1",
        "Q26r3c2",
    )
    return QuestionSpec(
        question_id="[Q26]",
        canonical_id="Q26",
        question_text="What role did each stakeholder play",
        question_type=QuestionType.GRID_BINARY_SELECT,
        raw_columns=raw_columns,
        option_map={"1": "Blocked Vendors", "2": "Scored Vendors"},
        value_range=(0, 1),
        denominator_policy=DenominatorPolicy.VALID_RESPONSES,
        grid_row_labels={
            "Q26r1c1": "IT",
            "Q26r1c2": "IT",
            "Q26r2c1": "Security",
            "Q26r2c2": "Security",
            "Q26r3c1": "Finance",
            "Q26r3c2": "Finance",
        },
        possible_role="GRID_BINARY_SELECT",
    )


class TestGridBinaryPivot(unittest.TestCase):
    def test_grid_binary_pivot_basic(self) -> None:
        df = pd.DataFrame(
            {
                "Q26r1c1": [1, 0, 1, None, 0],
                "Q26r1c2": [0, 1, 1, None, 0],
                "Q26r2c1": [0, 1, 0, 1, None],
                "Q26r2c2": [1, 1, 0, 0, None],
                "Q26r3c1": [None, 0, 0, 1, 1],
                "Q26r3c2": [None, 1, 0, 1, 0],
            }
        )

        result = compute_grid_binary_pivot(make_binary_grid_spec(), df, CalculationLog())

        self.assertIsInstance(result, GridBinaryPivotResult)
        self.assertEqual(result.column_headers, ["Blocked Vendors", "Scored Vendors"])
        self.assertEqual(len(result.rows), 3)
        self.assertEqual(result.rows[0].counts_per_column, [2, 2])

    def test_grid_binary_pivot_handles_label_values(self) -> None:
        df = pd.DataFrame(
            {
                "Q26r1c1": ["Blocked Vendors", "NO TO: Blocked Vendors", None],
                "Q26r1c2": ["NO TO: Scored Vendors", "Scored Vendors", None],
                "Q26r2c1": [None, None, None],
                "Q26r2c2": [None, None, None],
                "Q26r3c1": [None, None, None],
                "Q26r3c2": [None, None, None],
            }
        )

        result = compute_grid_binary_pivot(make_binary_grid_spec(), df, CalculationLog())

        self.assertEqual(result.rows[0].counts_per_column, [1, 1])
        self.assertEqual(result.rows[0].pcts_per_column, [0.5, 0.5])

    def test_grid_binary_pivot_uses_row_denominator(self) -> None:
        df = pd.DataFrame(
            {
                "Q26r1c1": [1, None, None],
                "Q26r1c2": [0, None, None],
                "Q26r2c1": [1, 1, 0],
                "Q26r2c2": [0, 1, 1],
                "Q26r3c1": [None, None, None],
                "Q26r3c2": [None, None, None],
            }
        )

        result = compute_grid_binary_pivot(make_binary_grid_spec(), df, CalculationLog())

        self.assertEqual(result.rows[0].pcts_per_column, [1.0, 0.0])
        self.assertEqual(result.rows[1].pcts_per_column, [2 / 3, 2 / 3])


if __name__ == "__main__":
    unittest.main()
