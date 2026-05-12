"""Tests for numeric single-cut calculators."""

from __future__ import annotations

import math
import unittest

import pandas as pd

from src.calculation_log import CalculationLog
from src.models import (
    DenominatorPolicy,
    QuestionSpec,
    QuestionType,
)
from src.single_cut._numeric import compute_numeric
from tests.conftest import GOLDEN_30_RESPONDENTS_PATH


Q_NUM_1_MEAN = 5.5
Q_NUM_1_MEDIAN = 5.5
Q_NUM_1_STD = 2.9213837061606083
Q_NUM_1_P25 = 3.0
Q_NUM_1_P50 = 5.5
Q_NUM_1_P75 = 8.0

Q_NUM_2_MEAN = 50.370370370370374
Q_NUM_2_MEDIAN = 50.0
Q_NUM_2_STD = 26.95886916731907
Q_NUM_2_MIN = 10.0
Q_NUM_2_MAX = 100.0
Q_NUM_2_P25 = 30.0
Q_NUM_2_P50 = 50.0
Q_NUM_2_P75 = 70.0

Q_ALLOC_1_PER_OPTION = {
    "Q_ALLOC_1r1": {
        "mean": 48.5,
        "median": 50.0,
        "std": 6.0,
        "min_val": 30.0,
        "max_val": 60.0,
        "valid_n": 30,
        "missing_n": 0,
    },
    "Q_ALLOC_1r2": {
        "mean": 30.9,
        "median": 30.0,
        "std": 3.7,
        "min_val": 30.0,
        "max_val": 50.0,
        "valid_n": 30,
        "missing_n": 0,
    },
    "Q_ALLOC_1r3": {
        "mean": 20.6,
        "median": 20.0,
        "std": 4.9,
        "min_val": 10.0,
        "max_val": 35.0,
        "valid_n": 30,
        "missing_n": 0,
    },
}
Q_ALLOC_1_AGGREGATE = {
    "mean": 33.333333333333336,
    "median": 30.9,
    "std": 14.108271805339353,
    "min": 20.6,
    "max": 48.5,
    "p25": 25.75,
    "p50": 30.9,
    "p75": 39.7,
}
Q_ALLOC_2_AGGREGATE = {
    "mean": 33.333333333333336,
    "median": 30.0,
    "std": 15.275252316519467,
    "min": 20.0,
    "max": 50.0,
    "p25": 25.0,
    "p50": 30.0,
    "p75": 40.0,
}


def load_golden() -> pd.DataFrame:
    return pd.read_csv(GOLDEN_30_RESPONDENTS_PATH)


def make_direct_numeric_spec(canonical_id: str) -> QuestionSpec:
    return QuestionSpec(
        question_id=f"[{canonical_id}]",
        canonical_id=canonical_id,
        question_text=f"{canonical_id} text",
        question_type=QuestionType.DIRECT_NUMERIC,
        raw_columns=(canonical_id,),
        option_map={},
        denominator_policy=DenominatorPolicy.VALID_RESPONSES,
    )


def make_allocation_spec(
    canonical_id: str,
    raw_columns: tuple[str, ...],
) -> QuestionSpec:
    return QuestionSpec(
        question_id=f"[{canonical_id}]",
        canonical_id=canonical_id,
        question_text=f"{canonical_id} text",
        question_type=QuestionType.NUMERIC_ALLOCATION,
        raw_columns=raw_columns,
        option_map={column: column for column in raw_columns},
        value_range=(0, 999),
        denominator_policy=DenominatorPolicy.VALID_RESPONSES,
    )


class TestNumeric(unittest.TestCase):
    def test_direct_numeric_basic(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        spec = make_direct_numeric_spec("Q_NUM_1")

        result = compute_numeric(spec, dataframe, log)

        self.assertEqual(result.valid_n, 30)
        self.assertEqual(result.missing_n, 0)
        self.assertEqual(result.mean, Q_NUM_1_MEAN)
        self.assertEqual(result.median, Q_NUM_1_MEDIAN)
        self.assertEqual(result.std, Q_NUM_1_STD)
        self.assertEqual(result.min_val, 1.0)
        self.assertEqual(result.max_val, 10.0)
        self.assertEqual(result.percentiles[25], Q_NUM_1_P25)
        self.assertEqual(result.percentiles[50], Q_NUM_1_P50)
        self.assertEqual(result.percentiles[75], Q_NUM_1_P75)

    def test_direct_numeric_with_missing(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        spec = make_direct_numeric_spec("Q_NUM_2")

        result = compute_numeric(spec, dataframe, log)

        self.assertEqual(result.valid_n, 27)
        self.assertEqual(result.missing_n, 3)
        self.assertEqual(result.mean, Q_NUM_2_MEAN)
        self.assertEqual(result.median, Q_NUM_2_MEDIAN)
        self.assertEqual(result.std, Q_NUM_2_STD)
        self.assertEqual(result.min_val, Q_NUM_2_MIN)
        self.assertEqual(result.max_val, Q_NUM_2_MAX)
        self.assertEqual(result.percentiles[25], Q_NUM_2_P25)
        self.assertEqual(result.percentiles[50], Q_NUM_2_P50)
        self.assertEqual(result.percentiles[75], Q_NUM_2_P75)

    def test_direct_numeric_filter_mask_applied(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        spec = make_direct_numeric_spec("Q_NUM_1")
        filter_mask = dataframe.index < 10

        result = compute_numeric(
            spec,
            dataframe,
            log,
            filter_mask=filter_mask,
            filter_expr="first ten respondents",
        )

        self.assertEqual(result.valid_n, 10)
        self.assertEqual(result.missing_n, 0)
        self.assertEqual(result.mean, 5.5)
        self.assertEqual(result.audit_records[0].filter_expr, "first ten respondents")

    def test_direct_numeric_all_missing_returns_nan(self) -> None:
        dataframe = pd.DataFrame({"Q_NUM_EMPTY": [pd.NA, pd.NA, pd.NA]})
        log = CalculationLog()
        spec = make_direct_numeric_spec("Q_NUM_EMPTY")

        result = compute_numeric(spec, dataframe, log)

        self.assertEqual(result.valid_n, 0)
        self.assertEqual(result.missing_n, 3)
        self.assertTrue(math.isnan(result.mean))
        self.assertTrue(math.isnan(result.median))
        self.assertTrue(math.isnan(result.std))
        self.assertIn("all values are missing", result.warnings[0])

    def test_direct_numeric_audit_record_logged(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        spec = make_direct_numeric_spec("Q_NUM_1")

        result = compute_numeric(spec, dataframe, log)

        self.assertEqual(len(log), 5)
        self.assertEqual(log.all_records(), result.audit_records)
        self.assertEqual(result.audit_records[0].metric_name, "numeric_summary")
        metric_names = [record.metric_name for record in result.audit_records]
        self.assertEqual(metric_names.count("numeric_std"), 1)
        self.assertEqual(metric_names.count("numeric_p25"), 1)
        self.assertEqual(metric_names.count("numeric_p50"), 1)
        self.assertEqual(metric_names.count("numeric_p75"), 1)

    def test_allocation_basic(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        spec = make_allocation_spec(
            "Q_ALLOC_1",
            ("Q_ALLOC_1r1", "Q_ALLOC_1r2", "Q_ALLOC_1r3"),
        )

        result = compute_numeric(spec, dataframe, log)

        self.assertEqual(result.valid_n, 28)
        self.assertEqual(result.missing_n, 2)
        self.assertEqual(result.allocation_target, 100.0)
        self.assertEqual(result.allocation_tolerance, 2.0)
        self.assertEqual(result.allocation_excluded_n, 2)

    def test_allocation_excludes_out_of_tolerance(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        spec = make_allocation_spec(
            "Q_ALLOC_1",
            ("Q_ALLOC_1r1", "Q_ALLOC_1r2", "Q_ALLOC_1r3"),
        )

        result = compute_numeric(spec, dataframe, log)

        self.assertEqual(result.audit_records[0].valid_n, 28)
        self.assertEqual(result.audit_records[0].missing_n, 2)
        self.assertIn(
            "2 respondents excluded for sum outside tolerance",
            result.warnings,
        )

    def test_allocation_per_option_stats_populated(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        spec = make_allocation_spec(
            "Q_ALLOC_1",
            ("Q_ALLOC_1r1", "Q_ALLOC_1r2", "Q_ALLOC_1r3"),
        )

        result = compute_numeric(spec, dataframe, log)

        self.assertEqual(result.per_option_stats, Q_ALLOC_1_PER_OPTION)

    def test_allocation_with_some_na_respondents(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        spec = make_allocation_spec(
            "Q_ALLOC_2",
            ("Q_ALLOC_2r1", "Q_ALLOC_2r2", "Q_ALLOC_2r3"),
        )

        result = compute_numeric(spec, dataframe, log)

        self.assertEqual(result.valid_n, 25)
        self.assertEqual(result.missing_n, 5)
        self.assertEqual(result.allocation_excluded_n, 0)
        self.assertEqual(
            result.per_option_stats,
            {
                "Q_ALLOC_2r1": {
                    "mean": 50.0,
                    "median": 50.0,
                    "std": 0.0,
                    "min_val": 50.0,
                    "max_val": 50.0,
                    "valid_n": 25,
                    "missing_n": 5,
                },
                "Q_ALLOC_2r2": {
                    "mean": 30.0,
                    "median": 30.0,
                    "std": 0.0,
                    "min_val": 30.0,
                    "max_val": 30.0,
                    "valid_n": 25,
                    "missing_n": 5,
                },
                "Q_ALLOC_2r3": {
                    "mean": 20.0,
                    "median": 20.0,
                    "std": 0.0,
                    "min_val": 20.0,
                    "max_val": 20.0,
                    "valid_n": 25,
                    "missing_n": 5,
                },
            },
        )
        self.assertEqual(result.mean, Q_ALLOC_2_AGGREGATE["mean"])
        self.assertEqual(result.std, Q_ALLOC_2_AGGREGATE["std"])

    def test_numeric_allocation_shows_mean_and_median(self) -> None:
        dataframe = pd.DataFrame(
            {
                "Q33r1": [10, 20, 30],
                "Q33r2": [90, 80, 70],
            }
        )
        log = CalculationLog()
        spec = make_allocation_spec("Q33", ("Q33r1", "Q33r2"))

        result = compute_numeric(spec, dataframe, log)

        self.assertEqual(
            result.per_option_stats,
            {
                "Q33r1": {
                    "mean": 20.0,
                    "median": 20.0,
                    "std": 10.0,
                    "min_val": 10.0,
                    "max_val": 30.0,
                    "valid_n": 3,
                    "missing_n": 0,
                },
                "Q33r2": {
                    "mean": 80.0,
                    "median": 80.0,
                    "std": 10.0,
                    "min_val": 70.0,
                    "max_val": 90.0,
                    "valid_n": 3,
                    "missing_n": 0,
                },
            },
        )
        metric_names = [record.metric_name for record in log.all_records()]
        self.assertEqual(metric_names.count("numeric_allocation_mean"), 2)
        self.assertEqual(metric_names.count("numeric_allocation_median"), 2)

    def test_numeric_allocation_result_exposes_full_stats(self) -> None:
        dataframe = pd.DataFrame({"Q33r1": [10, 20, 30, 40, 50]})
        log = CalculationLog()
        spec = make_allocation_spec("Q33", ("Q33r1",))

        result = compute_numeric(spec, dataframe, log)
        stats = result.per_option_stats["Q33r1"]

        self.assertEqual(stats["mean"], 30.0)
        self.assertEqual(stats["median"], 30.0)
        self.assertAlmostEqual(stats["std"], 15.8)
        self.assertEqual(stats["min_val"], 10.0)
        self.assertEqual(stats["max_val"], 50.0)
        self.assertEqual(stats["valid_n"], 5)
        self.assertEqual(stats["missing_n"], 0)

    def test_direct_numeric_result_exposes_percentiles(self) -> None:
        dataframe = pd.DataFrame({"Q_NUM_100": list(range(1, 101))})
        log = CalculationLog()
        spec = make_direct_numeric_spec("Q_NUM_100")

        result = compute_numeric(spec, dataframe, log)

        self.assertGreater(result.p25, 0.0)
        self.assertGreater(result.p50, 0.0)
        self.assertGreater(result.p75, 0.0)
        self.assertEqual(result.min_val, 1.0)
        self.assertEqual(result.max_val, 100.0)

    def test_allocation_aggregate_stats_correct(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        spec = make_allocation_spec(
            "Q_ALLOC_1",
            ("Q_ALLOC_1r1", "Q_ALLOC_1r2", "Q_ALLOC_1r3"),
        )

        result = compute_numeric(spec, dataframe, log)

        self.assertEqual(result.mean, Q_ALLOC_1_AGGREGATE["mean"])
        self.assertEqual(result.median, Q_ALLOC_1_AGGREGATE["median"])
        self.assertEqual(result.std, Q_ALLOC_1_AGGREGATE["std"])
        self.assertEqual(result.min_val, Q_ALLOC_1_AGGREGATE["min"])
        self.assertEqual(result.max_val, Q_ALLOC_1_AGGREGATE["max"])
        self.assertEqual(result.percentiles[25], Q_ALLOC_1_AGGREGATE["p25"])
        self.assertEqual(result.percentiles[50], Q_ALLOC_1_AGGREGATE["p50"])
        self.assertEqual(result.percentiles[75], Q_ALLOC_1_AGGREGATE["p75"])

    def test_numeric_raises_on_wrong_question_type(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        spec = QuestionSpec(
            question_id="[Q_BAD]",
            canonical_id="Q_BAD",
            question_text="Wrong type",
            question_type=QuestionType.SINGLE_SELECT,
            raw_columns=("Q_BAD",),
            option_map={1: "Yes", 2: "No"},
        )

        with self.assertRaisesRegex(ValueError, "unsupported question_type"):
            compute_numeric(spec, dataframe, log)


if __name__ == "__main__":
    unittest.main()
