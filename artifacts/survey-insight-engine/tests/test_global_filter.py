"""Tests for global filter application."""

from __future__ import annotations

import unittest

import pandas as pd

from src.global_filter import apply_global_filter
from src.models import FilterSpec, GlobalFilterState
from tests.conftest import GOLDEN_30_RESPONDENTS_PATH


def load_global_filter_golden() -> pd.DataFrame:
    dataframe = pd.read_csv(GOLDEN_30_RESPONDENTS_PATH)
    dataframe["Q_REGION"] = [1] * 10 + [2] * 10 + [3] * 10
    dataframe["Q_INDUSTRY"] = [1] * 15 + [2] * 15
    dataframe["Q_FUNCTION"] = ([1] * 5 + [2] * 5) * 3
    return dataframe


class TestGlobalFilter(unittest.TestCase):
    def test_empty_global_filter_returns_dataframe_unchanged(self) -> None:
        dataframe = load_global_filter_golden()
        filtered_df, stats = apply_global_filter(dataframe, GlobalFilterState())

        self.assertIs(filtered_df, dataframe)
        self.assertEqual(len(filtered_df), 30)
        self.assertEqual(
            stats,
            {
                "rows_before": 30,
                "rows_after": 30,
                "rows_removed": 0,
                "filter_description": "(no global filter)",
            },
        )

    def test_single_filter_restricts_dataframe(self) -> None:
        dataframe = load_global_filter_golden()
        filtered_df, stats = apply_global_filter(
            dataframe,
            GlobalFilterState(filters=(FilterSpec("Q_REGION", 1),)),
        )

        self.assertEqual(len(filtered_df), 10)
        self.assertEqual(stats["rows_after"], 10)
        self.assertTrue((filtered_df["Q_REGION"] == 1).all())

    def test_multiple_filters_combine_with_and(self) -> None:
        dataframe = load_global_filter_golden()
        filtered_df, stats = apply_global_filter(
            dataframe,
            GlobalFilterState(
                filters=(
                    FilterSpec("Q_REGION", 2),
                    FilterSpec("Q_INDUSTRY", 1),
                    FilterSpec("Q_FUNCTION", 1),
                )
            ),
        )

        self.assertEqual(len(filtered_df), 5)
        self.assertEqual(stats["rows_before"], 30)
        self.assertEqual(stats["rows_after"], 5)
        self.assertEqual(stats["rows_removed"], 25)

    def test_filter_on_missing_column_raises(self) -> None:
        dataframe = load_global_filter_golden()
        with self.assertRaisesRegex(ValueError, "global filter column"):
            apply_global_filter(
                dataframe,
                GlobalFilterState(filters=(FilterSpec("Q_MISSING", 1),)),
            )

    def test_breakdown_filter_in_global_state_raises(self) -> None:
        with self.assertRaisesRegex(ValueError, "does not allow breakdown"):
            GlobalFilterState(filters=(FilterSpec("Q_REGION"),))

    def test_duplicate_filter_question_id_raises(self) -> None:
        with self.assertRaisesRegex(ValueError, "duplicate global filter"):
            GlobalFilterState(
                filters=(FilterSpec("Q_REGION", 1), FilterSpec("Q_REGION", 2))
            )

    def test_stats_dict_populated_correctly(self) -> None:
        dataframe = load_global_filter_golden()
        _filtered_df, stats = apply_global_filter(
            dataframe,
            GlobalFilterState(filters=(FilterSpec("Q_REGION", 3),)),
        )

        self.assertEqual(stats["rows_before"], 30)
        self.assertEqual(stats["rows_after"], 10)
        self.assertEqual(stats["rows_removed"], 20)
        self.assertEqual(stats["filter_description"], "Q_REGION == 3")

    def test_filtered_dataframe_preserves_index_for_audit(self) -> None:
        dataframe = load_global_filter_golden()
        filtered_df, _stats = apply_global_filter(
            dataframe,
            GlobalFilterState(filters=(FilterSpec("Q_INDUSTRY", 2),)),
        )

        self.assertEqual(list(filtered_df.index), list(range(15, 30)))


if __name__ == "__main__":
    unittest.main()


class TestFilterSpecMultiValue(unittest.TestCase):
    def test_filter_values_isin_semantics(self) -> None:
        dataframe = load_global_filter_golden()
        filtered_df, stats = apply_global_filter(
            dataframe,
            GlobalFilterState(
                filters=(FilterSpec("Q_REGION", filter_values=(1, 3)),)
            ),
        )
        self.assertEqual(len(filtered_df), 20)
        self.assertEqual(set(filtered_df["Q_REGION"].unique()), {1, 3})
        self.assertIn("in [1, 3]", stats["filter_description"])

    def test_filter_values_single_element_matches_scalar(self) -> None:
        dataframe = load_global_filter_golden()
        filtered_a, _ = apply_global_filter(
            dataframe,
            GlobalFilterState(filters=(FilterSpec("Q_REGION", 2),)),
        )
        filtered_b, _ = apply_global_filter(
            dataframe,
            GlobalFilterState(
                filters=(FilterSpec("Q_REGION", filter_values=(2,)),)
            ),
        )
        self.assertEqual(len(filtered_a), len(filtered_b))
        self.assertTrue(filtered_a.equals(filtered_b))

    def test_empty_filter_values_treated_as_breakdown(self) -> None:
        spec = FilterSpec("Q_REGION", filter_values=())
        self.assertTrue(spec.is_breakdown())
        self.assertIsNone(spec.get_effective_values())
        with self.assertRaisesRegex(ValueError, "does not allow breakdown"):
            GlobalFilterState(filters=(spec,))

    def test_filter_value_takes_precedence_over_filter_values(self) -> None:
        spec = FilterSpec("Q_REGION", filter_value=1, filter_values=(2, 3))
        self.assertEqual(spec.get_effective_values(), [1])
        self.assertFalse(spec.is_breakdown())
