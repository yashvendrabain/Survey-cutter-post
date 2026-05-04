"""Golden tests for vectorised calculation primitives."""

from __future__ import annotations

import math
import unittest

import pandas as pd

from src.calc_primitives import (
    allocation_summary,
    count_by_value,
    numeric_summary,
    rate_per_value,
    selection_rate,
)
from tests.conftest import GOLDEN_30_RESPONDENTS_PATH


OUTPUT_SHEET = "Single Cuts"
Q_SS_1_COUNTS = {1: 18, 2: 12}
Q_SS_1_RATES = {
    1: 0.6,
    2: 0.4,
}
Q_SS_2_RATES = {
    1: 0.16666666666666666,
    2: 0.16666666666666666,
    3: 0.3333333333333333,
    4: 0.3333333333333333,
}
Q_SS_3_RATES = {
    1: 0.37037037037037035,
    2: 0.37037037037037035,
    3: 0.25925925925925924,
}
EXPECTED_NUMERIC_STD = 2.9213837061606083
Q_ALLOC_1_MEANS = {
    "Q_ALLOC_1r1": 49.107142857142854,
    "Q_ALLOC_1r2": 30.214285714285715,
    "Q_ALLOC_1r3": 20.678571428571427,
}


def load_golden() -> pd.DataFrame:
    return pd.read_csv(GOLDEN_30_RESPONDENTS_PATH)


class TestCalcPrimitives(unittest.TestCase):
    def test_count_by_value_basic(self) -> None:
        dataframe = load_golden()
        result, audit = count_by_value(
            dataframe["Q_SS_1"], "Q_SS_1", ("Q_SS_1",), OUTPUT_SHEET
        )

        self.assertEqual(result, Q_SS_1_COUNTS)
        self.assertEqual(audit.metric_name, "value_counts")
        self.assertEqual(audit.valid_n, 30)
        self.assertEqual(audit.missing_n, 0)

    def test_count_by_value_with_na(self) -> None:
        dataframe = load_golden()
        result, audit = count_by_value(
            dataframe["Q_SS_3"], "Q_SS_3", ("Q_SS_3",), OUTPUT_SHEET
        )

        self.assertEqual(result, {1: 10, 2: 10, 3: 7})
        self.assertEqual(audit.valid_n, 27)
        self.assertEqual(audit.missing_n, 3)

    def test_rate_per_value_basic(self) -> None:
        dataframe = load_golden()
        result, audit = rate_per_value(
            dataframe["Q_SS_1"], "Q_SS_1", ("Q_SS_1",), OUTPUT_SHEET
        )

        self.assertEqual(result[1]["count"], 18)
        self.assertEqual(result[1]["rate"], Q_SS_1_RATES[1])
        self.assertEqual(result[2]["count"], 12)
        self.assertEqual(result[2]["rate"], Q_SS_1_RATES[2])
        self.assertEqual(audit.denominator, 30)

    def test_rate_per_value_excludes_na_from_denominator(self) -> None:
        dataframe = load_golden()
        result, audit = rate_per_value(
            dataframe["Q_SS_3"], "Q_SS_3", ("Q_SS_3",), OUTPUT_SHEET
        )

        self.assertEqual(result[1]["rate"], Q_SS_3_RATES[1])
        self.assertEqual(result[2]["rate"], Q_SS_3_RATES[2])
        self.assertEqual(result[3]["rate"], Q_SS_3_RATES[3])
        self.assertEqual(audit.denominator, 27)

    def test_selection_rate_valid_responses_policy(self) -> None:
        dataframe = load_golden()
        result, audit = selection_rate(
            {
                "Q_MS_1r1": dataframe["Q_MS_1r1"],
                "Q_MS_1r2": dataframe["Q_MS_1r2"],
                "Q_MS_1r3": dataframe["Q_MS_1r3"],
            },
            "Q_MS_1",
            OUTPUT_SHEET,
            "valid_responses",
        )

        self.assertEqual(result["Q_MS_1r1"]["count"], 12)
        self.assertEqual(result["Q_MS_1r1"]["selection_rate"], 0.4)
        self.assertEqual(result["Q_MS_1r2"]["count"], 6)
        self.assertEqual(result["Q_MS_1r2"]["selection_rate"], 0.2)
        self.assertEqual(result["Q_MS_1r3"]["count"], 0)
        self.assertEqual(result["Q_MS_1r3"]["selection_rate"], 0.0)
        self.assertEqual(audit.valid_n, 30)
        self.assertEqual(audit.missing_n, 0)

    def test_selection_rate_all_respondents_policy(self) -> None:
        dataframe = load_golden()
        result, audit = selection_rate(
            {"Q_MS_1r1": dataframe["Q_MS_1r1"]},
            "Q_MS_1",
            OUTPUT_SHEET,
            "all_respondents",
            all_respondents_n=50,
        )

        self.assertEqual(result["Q_MS_1r1"]["count"], 12)
        self.assertEqual(result["Q_MS_1r1"]["selection_rate"], 0.24)
        self.assertEqual(audit.valid_n, 50)
        self.assertEqual(audit.missing_n, 20)

    def test_selection_rate_handles_1_2_encoding(self) -> None:
        series = pd.Series([1, 2, 0, pd.NA, 2])
        result, audit = selection_rate(
            {"Q_MS_ENCODEDr1": series},
            "Q_MS_ENCODED",
            OUTPUT_SHEET,
            "valid_responses",
        )

        self.assertEqual(result["Q_MS_ENCODEDr1"]["count"], 3)
        self.assertEqual(result["Q_MS_ENCODEDr1"]["selection_rate"], 0.75)
        self.assertEqual(audit.valid_n, 4)
        self.assertEqual(audit.missing_n, 1)

    def test_numeric_summary_basic(self) -> None:
        dataframe = load_golden()
        result, audit = numeric_summary(
            dataframe["Q_NUM_1"], "Q_NUM_1", ("Q_NUM_1",), OUTPUT_SHEET
        )

        self.assertEqual(result["mean"], 5.5)
        self.assertEqual(result["median"], 5.5)
        self.assertEqual(result["std"], EXPECTED_NUMERIC_STD)
        self.assertEqual(result["min"], 1.0)
        self.assertEqual(result["max"], 10.0)
        self.assertEqual(result["p25"], 3.0)
        self.assertEqual(result["p50"], 5.5)
        self.assertEqual(result["p75"], 8.0)
        self.assertEqual(result["valid_n"], 30)
        self.assertEqual(result["missing_n"], 0)
        self.assertEqual(audit.value_raw, 5.5)

    def test_numeric_summary_handles_all_missing(self) -> None:
        series = pd.Series([pd.NA, pd.NA, pd.NA])
        result, audit = numeric_summary(series, "Q_EMPTY", ("Q_EMPTY",), OUTPUT_SHEET)

        self.assertTrue(math.isnan(result["mean"]))
        self.assertTrue(math.isnan(result["median"]))
        self.assertTrue(math.isnan(result["std"]))
        self.assertEqual(result["valid_n"], 0)
        self.assertEqual(result["missing_n"], 3)
        self.assertTrue(math.isnan(audit.value_raw))

    def test_allocation_summary_excludes_out_of_tolerance(self) -> None:
        dataframe = load_golden()
        result, audit = allocation_summary(
            {
                "Q_ALLOC_1r1": dataframe["Q_ALLOC_1r1"],
                "Q_ALLOC_1r2": dataframe["Q_ALLOC_1r2"],
                "Q_ALLOC_1r3": dataframe["Q_ALLOC_1r3"],
            },
            "Q_ALLOC_1",
            OUTPUT_SHEET,
            target_sum=100.0,
            tolerance=2.0,
        )

        self.assertEqual(result["answered_n"], 30)
        self.assertEqual(result["included_n"], 28)
        self.assertEqual(result["excluded_tolerance_n"], 2)
        self.assertEqual(
            result["per_option"]["Q_ALLOC_1r1"]["mean"],
            Q_ALLOC_1_MEANS["Q_ALLOC_1r1"],
        )
        self.assertEqual(
            result["per_option"]["Q_ALLOC_1r2"]["mean"],
            Q_ALLOC_1_MEANS["Q_ALLOC_1r2"],
        )
        self.assertEqual(
            result["per_option"]["Q_ALLOC_1r3"]["mean"],
            Q_ALLOC_1_MEANS["Q_ALLOC_1r3"],
        )
        self.assertEqual(result["per_option"]["Q_ALLOC_1r1"]["median"], 50.0)
        self.assertEqual(result["per_option"]["Q_ALLOC_1r2"]["median"], 30.0)
        self.assertEqual(result["per_option"]["Q_ALLOC_1r3"]["median"], 20.0)
        self.assertEqual(audit.valid_n, 28)
        self.assertEqual(audit.missing_n, 2)

    def test_audit_record_populated_correctly_for_each_function(self) -> None:
        dataframe = load_golden()
        _, count_audit = count_by_value(
            dataframe["Q_SS_1"], "Q_SS_1", ("Q_SS_1",), OUTPUT_SHEET, "segment == A"
        )
        _, rate_audit = rate_per_value(
            dataframe["Q_SS_1"], "Q_SS_1", ("Q_SS_1",), OUTPUT_SHEET
        )
        _, selection_audit = selection_rate(
            {"Q_MS_1r1": dataframe["Q_MS_1r1"]},
            "Q_MS_1",
            OUTPUT_SHEET,
            "valid_responses",
        )
        _, numeric_audit = numeric_summary(
            dataframe["Q_NUM_1"], "Q_NUM_1", ("Q_NUM_1",), OUTPUT_SHEET
        )
        _, allocation_audit = allocation_summary(
            {
                "Q_ALLOC_1r1": dataframe["Q_ALLOC_1r1"],
                "Q_ALLOC_1r2": dataframe["Q_ALLOC_1r2"],
                "Q_ALLOC_1r3": dataframe["Q_ALLOC_1r3"],
            },
            "Q_ALLOC_1",
            OUTPUT_SHEET,
            100.0,
            2.0,
        )

        self.assertEqual(count_audit.metric_name, "value_counts")
        self.assertEqual(count_audit.filter_expr, "segment == A")
        self.assertEqual(rate_audit.metric_name, "rate_per_value")
        self.assertEqual(selection_audit.metric_name, "selection_rate")
        self.assertEqual(numeric_audit.metric_name, "numeric_summary")
        self.assertEqual(allocation_audit.metric_name, "allocation_summary")
        self.assertEqual(allocation_audit.source_columns, (
            "Q_ALLOC_1r1",
            "Q_ALLOC_1r2",
            "Q_ALLOC_1r3",
        ))


if __name__ == "__main__":
    unittest.main()
