"""Tests for the single-select calculator."""

from __future__ import annotations

import unittest

import pandas as pd

from src.calculation_log import CalculationLog
from src.models import DenominatorPolicy, QuestionSpec, QuestionType
from src.single_cut._single_select import compute_single_select
from tests.conftest import GOLDEN_30_RESPONDENTS_PATH


Q_SS_1_RATES = {1: 0.6, 2: 0.4}
Q_SS_3_RATES = {
    1: 0.37037037037037035,
    2: 0.37037037037037035,
    3: 0.25925925925925924,
}


def load_golden() -> pd.DataFrame:
    return pd.read_csv(GOLDEN_30_RESPONDENTS_PATH)


def make_single_select_spec(
    canonical_id: str,
    option_map: dict[int | str, str],
) -> QuestionSpec:
    return QuestionSpec(
        question_id=f"[{canonical_id}]",
        canonical_id=canonical_id,
        question_text=f"{canonical_id} text",
        question_type=QuestionType.SINGLE_SELECT,
        raw_columns=(canonical_id,),
        option_map=option_map,
        value_range=None,
        denominator_policy=DenominatorPolicy.VALID_RESPONSES,
    )


class TestSingleSelect(unittest.TestCase):
    def test_single_select_no_missing(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        spec = make_single_select_spec("Q_SS_1", {1: "Yes", 2: "No"})

        result = compute_single_select(spec, dataframe, log)

        self.assertEqual(result.valid_n, 30)
        self.assertEqual(result.missing_n, 0)
        self.assertEqual(result.distribution[1]["label"], "Yes")
        self.assertEqual(result.distribution[1]["count"], 18)
        self.assertEqual(result.distribution[1]["rate"], Q_SS_1_RATES[1])
        self.assertEqual(result.distribution[2]["label"], "No")
        self.assertEqual(result.distribution[2]["count"], 12)
        self.assertEqual(result.distribution[2]["rate"], Q_SS_1_RATES[2])

    def test_single_select_with_missing(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        spec = make_single_select_spec(
            "Q_SS_3",
            {1: "First", 2: "Second", 3: "Third"},
        )

        result = compute_single_select(spec, dataframe, log)

        self.assertEqual(result.valid_n, 27)
        self.assertEqual(result.missing_n, 3)
        self.assertEqual(result.distribution[1]["rate"], Q_SS_3_RATES[1])
        self.assertEqual(result.distribution[2]["rate"], Q_SS_3_RATES[2])
        self.assertEqual(result.distribution[3]["rate"], Q_SS_3_RATES[3])

    def test_single_select_filter_mask_applied(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        spec = make_single_select_spec("Q_SS_1", {1: "Yes", 2: "No"})
        filter_mask = dataframe.index < 10

        result = compute_single_select(
            spec,
            dataframe,
            log,
            filter_mask=filter_mask,
            filter_expr="first ten respondents",
        )

        self.assertEqual(result.valid_n, 10)
        self.assertEqual(result.missing_n, 0)
        self.assertEqual(result.distribution[1]["count"], 10)
        self.assertEqual(result.distribution[1]["rate"], 1.0)
        self.assertEqual(result.audit_records[0].filter_expr, "first ten respondents")

    def test_single_select_audit_record_logged(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        spec = make_single_select_spec("Q_SS_1", {1: "Yes", 2: "No"})

        result = compute_single_select(spec, dataframe, log)

        self.assertEqual(len(log), 1)
        self.assertEqual(log.all_records(), result.audit_records)
        self.assertEqual(result.audit_records[0].metric_name, "rate_per_value")

    def test_single_select_unmapped_code_warning(self) -> None:
        dataframe = pd.DataFrame({"Q_UNMAPPED": [1, 1, 9, 9]})
        log = CalculationLog()
        spec = make_single_select_spec("Q_UNMAPPED", {1: "Mapped"})

        result = compute_single_select(spec, dataframe, log)

        self.assertEqual(result.distribution[9]["label"], "9")
        self.assertEqual(
            result.warnings,
            ("unmapped option code in raw data: 9",),
        )


if __name__ == "__main__":
    unittest.main()
