"""Tests for the multi-select calculator."""

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
from src.single_cut._multi_select import compute_multi_select
from tests.conftest import GOLDEN_30_RESPONDENTS_PATH


Q_MS_1_RATES = {"Q_MS_1r1": 0.4, "Q_MS_1r2": 0.2, "Q_MS_1r3": 0.0}
Q_MS_2_RATES = {"Q_MS_2r1": 0.6, "Q_MS_2r2": 0.32, "Q_MS_2r3": 0.0}


def load_golden() -> pd.DataFrame:
    return pd.read_csv(GOLDEN_30_RESPONDENTS_PATH)


def make_multi_select_spec(
    canonical_id: str,
    raw_columns: tuple[str, ...],
    option_map: dict[str, str],
    denominator_policy: DenominatorPolicy = DenominatorPolicy.VALID_RESPONSES,
) -> QuestionSpec:
    return QuestionSpec(
        question_id=f"[{canonical_id}]",
        canonical_id=canonical_id,
        question_text=f"{canonical_id} text",
        question_type=QuestionType.MULTI_SELECT_BINARY,
        raw_columns=raw_columns,
        option_map=option_map,
        value_range=(0, 1),
        denominator_policy=denominator_policy,
    )


class TestMultiSelect(unittest.TestCase):
    def test_multi_select_basic(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        spec = make_multi_select_spec(
            "Q_MS_1",
            ("Q_MS_1r1", "Q_MS_1r2", "Q_MS_1r3"),
            {
                "Q_MS_1r1": "First",
                "Q_MS_1r2": "Second",
                "Q_MS_1r3": "Third",
            },
        )

        result = compute_multi_select(spec, dataframe, log)

        self.assertEqual(result.valid_n, 30)
        self.assertEqual(result.missing_n, 0)
        self.assertEqual(result.respondents_who_answered_any, 30)
        self.assertEqual(result.selections["Q_MS_1r1"]["count"], 12)
        self.assertEqual(
            result.selections["Q_MS_1r1"]["selection_rate"],
            Q_MS_1_RATES["Q_MS_1r1"],
        )
        self.assertEqual(result.selections["Q_MS_1r2"]["count"], 6)
        self.assertEqual(
            result.selections["Q_MS_1r2"]["selection_rate"],
            Q_MS_1_RATES["Q_MS_1r2"],
        )
        self.assertEqual(result.selections["Q_MS_1r3"]["count"], 0)
        self.assertEqual(
            result.selections["Q_MS_1r3"]["selection_rate"],
            Q_MS_1_RATES["Q_MS_1r3"],
        )

    def test_multi_select_valid_responses_denominator(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        spec = make_multi_select_spec(
            "Q_MS_2",
            ("Q_MS_2r1", "Q_MS_2r2", "Q_MS_2r3"),
            {
                "Q_MS_2r1": "First",
                "Q_MS_2r2": "Second",
                "Q_MS_2r3": "Third",
            },
        )

        result = compute_multi_select(spec, dataframe, log)

        self.assertEqual(result.valid_n, 25)
        self.assertEqual(result.missing_n, 5)
        self.assertEqual(result.respondents_who_answered_any, 25)
        self.assertEqual(result.selections["Q_MS_2r1"]["count"], 15)
        self.assertEqual(
            result.selections["Q_MS_2r1"]["selection_rate"],
            Q_MS_2_RATES["Q_MS_2r1"],
        )
        self.assertEqual(result.selections["Q_MS_2r2"]["count"], 8)
        self.assertEqual(
            result.selections["Q_MS_2r2"]["selection_rate"],
            Q_MS_2_RATES["Q_MS_2r2"],
        )
        self.assertEqual(result.selections["Q_MS_2r3"]["count"], 0)
        self.assertEqual(
            result.selections["Q_MS_2r3"]["selection_rate"],
            Q_MS_2_RATES["Q_MS_2r3"],
        )

    def test_multi_select_with_missing_respondents(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        spec = make_multi_select_spec(
            "Q_MS_2",
            ("Q_MS_2r1", "Q_MS_2r2", "Q_MS_2r3"),
            {
                "Q_MS_2r1": "First",
                "Q_MS_2r2": "Second",
                "Q_MS_2r3": "Third",
            },
        )

        result = compute_multi_select(spec, dataframe, log)

        self.assertEqual(result.audit_records[0].denominator, 25)
        self.assertEqual(result.audit_records[0].missing_n, 5)

    def test_multi_select_filter_mask_applied(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        spec = make_multi_select_spec(
            "Q_MS_1",
            ("Q_MS_1r1", "Q_MS_1r2", "Q_MS_1r3"),
            {
                "Q_MS_1r1": "First",
                "Q_MS_1r2": "Second",
                "Q_MS_1r3": "Third",
            },
        )
        filter_mask = dataframe.index < 10

        result = compute_multi_select(
            spec,
            dataframe,
            log,
            filter_mask=filter_mask,
            filter_expr="first ten respondents",
        )

        self.assertEqual(result.valid_n, 10)
        self.assertEqual(result.selections["Q_MS_1r1"]["count"], 10)
        self.assertEqual(result.selections["Q_MS_1r1"]["selection_rate"], 1.0)
        self.assertEqual(result.selections["Q_MS_1r2"]["count"], 6)
        self.assertEqual(result.selections["Q_MS_1r2"]["selection_rate"], 0.6)
        self.assertEqual(result.audit_records[0].filter_expr, "first ten respondents")

    def test_multi_select_handles_all_na_subcolumns(self) -> None:
        dataframe = pd.DataFrame(
            {
                "Q_MS_EMPTYr1": [pd.NA, pd.NA, pd.NA],
                "Q_MS_EMPTYr2": [pd.NA, pd.NA, pd.NA],
            }
        )
        log = CalculationLog()
        spec = make_multi_select_spec(
            "Q_MS_EMPTY",
            ("Q_MS_EMPTYr1", "Q_MS_EMPTYr2"),
            {
                "Q_MS_EMPTYr1": "First",
                "Q_MS_EMPTYr2": "Second",
            },
        )

        result = compute_multi_select(spec, dataframe, log)

        self.assertEqual(result.valid_n, 0)
        self.assertEqual(result.missing_n, 3)
        self.assertEqual(result.respondents_who_answered_any, 0)
        self.assertEqual(result.selections["Q_MS_EMPTYr1"]["count"], 0)
        self.assertTrue(
            math.isnan(result.selections["Q_MS_EMPTYr1"]["selection_rate"])
        )
        self.assertIn("all sub-columns are 100% missing", result.warnings)

    def test_multi_select_unmapped_label_warning(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        spec = make_multi_select_spec(
            "Q_MS_1",
            ("Q_MS_1r1", "Q_MS_1r2", "Q_MS_1r3"),
            {"Q_MS_1r1": "First"},
        )

        result = compute_multi_select(spec, dataframe, log)

        self.assertEqual(result.selections["Q_MS_1r2"]["label"], "Q_MS_1r2")
        self.assertIn(
            "no label for Q_MS_1r2; using id as label",
            result.warnings,
        )

    def test_multi_select_audit_record_logged(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        spec = make_multi_select_spec(
            "Q_MS_1",
            ("Q_MS_1r1", "Q_MS_1r2", "Q_MS_1r3"),
            {
                "Q_MS_1r1": "First",
                "Q_MS_1r2": "Second",
                "Q_MS_1r3": "Third",
            },
        )

        result = compute_multi_select(spec, dataframe, log)

        self.assertEqual(len(log), 1)
        self.assertEqual(log.all_records(), result.audit_records)
        self.assertEqual(result.audit_records[0].metric_name, "selection_rate")

    def test_multi_select_raises_on_wrong_question_type(self) -> None:
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
            compute_multi_select(spec, dataframe, log)

    def test_multi_select_handles_missing_subcolumn(self) -> None:
        dataframe = pd.DataFrame({"Q_MS_PARTIALr1": [1, 0, 1]})
        log = CalculationLog()
        spec = make_multi_select_spec(
            "Q_MS_PARTIAL",
            ("Q_MS_PARTIALr1", "Q_MS_PARTIALr2"),
            {
                "Q_MS_PARTIALr1": "First",
                "Q_MS_PARTIALr2": "Second",
            },
        )

        result = compute_multi_select(spec, dataframe, log)

        self.assertEqual(tuple(result.selections), ("Q_MS_PARTIALr1",))
        self.assertEqual(result.selections["Q_MS_PARTIALr1"]["count"], 2)
        self.assertIn(
            "sub-column Q_MS_PARTIALr2 not found in data",
            result.warnings,
        )


if __name__ == "__main__":
    unittest.main()
