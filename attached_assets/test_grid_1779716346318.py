"""Tests for the grid single-select calculator."""

from __future__ import annotations

import unittest

import pandas as pd

from src.calculation_log import CalculationLog
from src.models import (
    DenominatorPolicy,
    GridSingleSelectResult,
    QuestionSpec,
    QuestionType,
    SingleSelectResult,
)
from src.single_cut._grid import compute_grid
from tests.conftest import GOLDEN_30_RESPONDENTS_PATH


GRID_OPTION_MAP = {
    1: "Strongly disagree",
    2: "Disagree",
    3: "Agree",
    4: "Strongly agree",
}
GRID_ROW_LABELS = {
    "Q_GRID_1r1": "First grid row",
    "Q_GRID_1r2": "Second grid row",
    "Q_GRID_1r3": "Third grid row",
}


def load_golden() -> pd.DataFrame:
    return pd.read_csv(GOLDEN_30_RESPONDENTS_PATH)


def make_grid_spec(
    canonical_id: str = "Q_GRID_1",
    raw_columns: tuple[str, ...] = ("Q_GRID_1r1", "Q_GRID_1r2", "Q_GRID_1r3"),
    grid_row_labels: dict[str, str] | None = None,
) -> QuestionSpec:
    labels = GRID_ROW_LABELS if grid_row_labels is None else grid_row_labels
    return QuestionSpec(
        question_id=f"[{canonical_id}]",
        canonical_id=canonical_id,
        question_text="Grid question",
        question_type=QuestionType.GRID_SINGLE_SELECT,
        raw_columns=raw_columns,
        option_map=GRID_OPTION_MAP,
        value_range=(1, 4),
        denominator_policy=DenominatorPolicy.VALID_RESPONSES,
        grid_row_labels=labels,
    )


class TestGrid(unittest.TestCase):
    def test_grid_basic_three_rows(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        spec = make_grid_spec()

        result = compute_grid(spec, dataframe, log)

        self.assertEqual(tuple(result.rows), tuple(GRID_ROW_LABELS))
        self.assertEqual(result.rows["Q_GRID_1r1"].distribution[1]["count"], 5)
        self.assertEqual(result.rows["Q_GRID_1r1"].distribution[1]["rate"], 5 / 30)
        self.assertEqual(result.rows["Q_GRID_1r1"].distribution[2]["count"], 10)
        self.assertEqual(result.rows["Q_GRID_1r1"].distribution[2]["rate"], 10 / 30)
        self.assertEqual(result.rows["Q_GRID_1r1"].distribution[3]["count"], 10)
        self.assertEqual(result.rows["Q_GRID_1r1"].distribution[3]["rate"], 10 / 30)
        self.assertEqual(result.rows["Q_GRID_1r1"].distribution[4]["count"], 5)
        self.assertEqual(result.rows["Q_GRID_1r1"].distribution[4]["rate"], 5 / 30)

    def test_grid_row_with_missing(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        spec = make_grid_spec()

        result = compute_grid(spec, dataframe, log)
        row = result.rows["Q_GRID_1r3"]

        self.assertEqual(row.valid_n, 25)
        self.assertEqual(row.missing_n, 5)
        self.assertEqual(row.distribution[1]["rate"], 3 / 25)
        self.assertEqual(row.distribution[2]["rate"], 7 / 25)
        self.assertEqual(row.distribution[3]["rate"], 10 / 25)
        self.assertEqual(row.distribution[4]["rate"], 5 / 25)

    def test_grid_overall_valid_n_correct(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        spec = make_grid_spec()

        result = compute_grid(spec, dataframe, log)

        self.assertEqual(result.valid_n, 30)
        self.assertEqual(result.overall_valid_n, 30)
        self.assertEqual(result.missing_n, 0)
        self.assertEqual(result.audit_records[0].denominator, 30)

    def test_grid_per_row_results_are_single_select_results(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        spec = make_grid_spec()

        result = compute_grid(spec, dataframe, log)

        self.assertIsInstance(result, GridSingleSelectResult)
        for row_result in result.rows.values():
            self.assertIsInstance(row_result, SingleSelectResult)
            self.assertIs(row_result.question_type, QuestionType.SINGLE_SELECT)

    def test_grid_filter_mask_applied(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        spec = make_grid_spec()
        filter_mask = dataframe.index < 10

        result = compute_grid(
            spec,
            dataframe,
            log,
            filter_mask=filter_mask,
            filter_expr="first ten respondents",
        )

        self.assertEqual(result.valid_n, 10)
        self.assertEqual(result.missing_n, 0)
        self.assertEqual(result.rows["Q_GRID_1r1"].valid_n, 10)
        self.assertEqual(result.rows["Q_GRID_1r1"].distribution[1]["count"], 5)
        self.assertEqual(result.rows["Q_GRID_1r1"].distribution[1]["rate"], 0.5)
        self.assertEqual(result.audit_records[0].filter_expr, "first ten respondents")

    def test_grid_handles_missing_subcolumn(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        labels = {
            "Q_GRID_1r1": "First grid row",
            "Q_GRID_1r_missing": "Missing grid row",
        }
        spec = make_grid_spec(
            raw_columns=("Q_GRID_1r1", "Q_GRID_1r_missing"),
            grid_row_labels=labels,
        )

        result = compute_grid(spec, dataframe, log)

        self.assertEqual(tuple(result.rows), ("Q_GRID_1r1",))
        self.assertEqual(result.valid_n, 30)
        self.assertIn(
            "row Q_GRID_1r_missing (Missing grid row) not in raw data; skipped",
            result.warnings,
        )

    def test_grid_handles_all_rows_missing(self) -> None:
        dataframe = pd.DataFrame({"other": [1, 2, 3]})
        log = CalculationLog()
        labels = {
            "Q_GRID_EMPTYr1": "First grid row",
            "Q_GRID_EMPTYr2": "Second grid row",
        }
        spec = make_grid_spec(
            canonical_id="Q_GRID_EMPTY",
            raw_columns=("Q_GRID_EMPTYr1", "Q_GRID_EMPTYr2"),
            grid_row_labels=labels,
        )

        result = compute_grid(spec, dataframe, log)

        self.assertEqual(result.rows, {})
        self.assertEqual(result.valid_n, 0)
        self.assertEqual(result.missing_n, 3)
        self.assertEqual(result.overall_valid_n, 0)
        self.assertIn("no grid rows present in raw data", result.warnings)

    def test_grid_single_select_excludes_unchecked_rows(self) -> None:
        dataframe = pd.DataFrame(
            {
                "Q_GRID_BINr1": [1, 0, 0],
                "Q_GRID_BINr2": [0, 0, 0],
                "Q_GRID_BINr3": [0, 1, 0],
                "Q_GRID_BINr4": [0, 0, 0],
            }
        )
        labels = {
            "Q_GRID_BINr1": "Selected first",
            "Q_GRID_BINr2": "Never selected second",
            "Q_GRID_BINr3": "Selected third",
            "Q_GRID_BINr4": "Never selected fourth",
        }
        spec = QuestionSpec(
            question_id="[Q_GRID_BIN]",
            canonical_id="Q_GRID_BIN",
            question_text="Binary grid",
            question_type=QuestionType.GRID_SINGLE_SELECT,
            raw_columns=tuple(labels),
            option_map={0: "Unchecked", 1: "Checked"},
            value_range=(0, 1),
            denominator_policy=DenominatorPolicy.VALID_RESPONSES,
            grid_row_labels=labels,
        )
        log = CalculationLog()

        result = compute_grid(spec, dataframe, log)

        self.assertEqual(tuple(result.rows), ("Q_GRID_BINr1", "Q_GRID_BINr3"))
        for row_result in result.rows.values():
            self.assertEqual(tuple(row_result.distribution), (1,))
            self.assertGreater(row_result.distribution[1]["count"], 0)

    def test_grid_audit_records_logged_per_row_plus_parent(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        spec = make_grid_spec()

        result = compute_grid(spec, dataframe, log)

        self.assertEqual(len(log), 4)
        self.assertEqual(log.all_records()[-1], result.audit_records[0])
        self.assertEqual(result.audit_records[0].metric_name, "grid_overall")

    def test_grid_raises_on_wrong_question_type(self) -> None:
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
            compute_grid(spec, dataframe, log)

    def test_grid_raises_on_missing_grid_row_labels(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        spec = make_grid_spec()
        object.__setattr__(spec, "grid_row_labels", None)

        with self.assertRaisesRegex(ValueError, "grid_row_labels"):
            compute_grid(spec, dataframe, log)


if __name__ == "__main__":
    unittest.main()
