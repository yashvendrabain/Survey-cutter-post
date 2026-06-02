from __future__ import annotations

from pathlib import Path
import uuid
import unittest

import pandas as pd

from src.raw_decoder import (
    _has_categorical_grid_row_columns,
    _has_double_colon_grid_columns,
    decode_raw_data,
)


Q27_FIELD_SALES_COLUMN = "Q27_Field_sales_direct_face_to_"
Q27_INSIDE_SALES_COLUMN = "Q27_Inside_sales_direct_sales_r"
Q27_FIELD_SALES_LABEL = "Field sales (direct, face-to-face outside sales reps)"
Q27_INSIDE_SALES_LABEL = "Inside sales (direct sales reps)"
Q14_PLANNED_REVENUE_COLUMN = "Q14_Planned_2024_Revenue_Growth"
Q14_ACTUAL_REVENUE_COLUMN = "Q14_Actual_2024_Revenue_Growth"
Q14_PLANNED_REVENUE_LABEL = "Planned 2024 :: Revenue Growth"
Q14_ACTUAL_REVENUE_LABEL = "Actual 2024 :: Revenue Growth"


def _write_raw_csv(rows: list[dict[str, object | None]]) -> Path:
    scratch_root = Path.cwd() / "outputs" / "test_tmp"
    scratch_root.mkdir(parents=True, exist_ok=True)
    path = scratch_root / f"raw_grid_real_{uuid.uuid4().hex}.csv"
    pd.DataFrame(rows).to_csv(path, index=False)
    return path


def _safe_unlink(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)
    except PermissionError:
        pass


def _data_map(question: dict) -> dict:
    return {
        "questions": [question],
        "source_path": "datamap.xlsx",
        "sheet_name": "Sheet1",
        "total_rows_in_sheet": 1,
        "parser_warnings": [],
    }


def _q27_question() -> dict:
    return {
        "canonical_id": "Q27",
        "raw_id": "Q27",
        "question_text": "How has usage changed by channel?",
        "type_hint": None,
        "value_range": (1, 5),
        "options": [
            (1, "More often"),
            (2, "Less often"),
            (3, "Same"),
            (4, "Not applicable"),
            (5, "Don't know"),
        ],
        "sub_columns": [
            (Q27_FIELD_SALES_COLUMN, Q27_FIELD_SALES_LABEL),
            (Q27_INSIDE_SALES_COLUMN, Q27_INSIDE_SALES_LABEL),
        ],
        "parent_canonical_id": None,
        "source_row": 1,
        "warnings": [],
    }


class TestRawDecoderGridPreservationReal(unittest.TestCase):
    def _decode(
        self,
        rows: list[dict[str, object | None]],
        question: dict,
    ) -> pd.DataFrame:
        raw_path = _write_raw_csv(rows)
        try:
            decoded, _report = decode_raw_data(str(raw_path), _data_map(question))
            return decoded
        finally:
            _safe_unlink(raw_path)

    def test_q27_style_columns_preserved_numeric(self) -> None:
        decoded = self._decode(
            [
                {
                    "respondent_id": "R1",
                    Q27_FIELD_SALES_COLUMN: 3,
                    Q27_INSIDE_SALES_COLUMN: 2,
                },
                {
                    "respondent_id": "R2",
                    Q27_FIELD_SALES_COLUMN: 1,
                    Q27_INSIDE_SALES_COLUMN: 5,
                },
            ],
            _q27_question(),
        )

        self.assertEqual(decoded[Q27_FIELD_SALES_COLUMN].tolist(), [3.0, 1.0])
        self.assertEqual(decoded[Q27_INSIDE_SALES_COLUMN].tolist(), [2.0, 5.0])
        self.assertNotIn("Selected", decoded[Q27_FIELD_SALES_COLUMN].tolist())
        self.assertNotIn("Same", decoded[Q27_FIELD_SALES_COLUMN].tolist())

    def test_q27_categorical_grid_detects_from_sub_column_labels(self) -> None:
        question = _q27_question()
        raw_columns = (
            "respondent_id",
            Q27_FIELD_SALES_COLUMN,
            Q27_INSIDE_SALES_COLUMN,
        )

        self.assertTrue(_has_categorical_grid_row_columns(question, raw_columns))

    def test_q14_style_double_colon_preserved_float(self) -> None:
        question = {
            "canonical_id": "Q14",
            "raw_id": "Q14",
            "question_text": "Actual percentages by metric",
            "type_hint": "open_numeric",
            "value_range": (0, 100),
            "options": [(1, "Selected"), (2, "Not selected")],
            "sub_columns": [
                (Q14_PLANNED_REVENUE_COLUMN, Q14_PLANNED_REVENUE_LABEL),
                (Q14_ACTUAL_REVENUE_COLUMN, Q14_ACTUAL_REVENUE_LABEL),
            ],
            "parent_canonical_id": None,
            "source_row": 1,
            "warnings": [],
        }
        decoded = self._decode(
            [
                {
                    "respondent_id": "R1",
                    Q14_PLANNED_REVENUE_COLUMN: 13.5,
                    Q14_ACTUAL_REVENUE_COLUMN: 10,
                },
                {
                    "respondent_id": "R2",
                    Q14_PLANNED_REVENUE_COLUMN: 15.0,
                    Q14_ACTUAL_REVENUE_COLUMN: 12.5,
                },
            ],
            question,
        )

        self.assertEqual(decoded[Q14_PLANNED_REVENUE_COLUMN].tolist(), [13.5, 15.0])
        self.assertEqual(decoded[Q14_ACTUAL_REVENUE_COLUMN].tolist(), [10.0, 12.5])
        self.assertNotIn("Selected", decoded[Q14_ACTUAL_REVENUE_COLUMN].tolist())
        self.assertTrue(
            pd.api.types.is_numeric_dtype(decoded[Q14_PLANNED_REVENUE_COLUMN])
        )

    def test_q14_double_colon_detects_from_sub_column_labels(self) -> None:
        question = {
            "canonical_id": "Q14",
            "raw_id": "Q14",
            "question_text": "Actual percentages by metric",
            "type_hint": "open_numeric",
            "value_range": (0, 100),
            "options": [(1, "Selected"), (2, "Not selected")],
            "sub_columns": [
                (Q14_PLANNED_REVENUE_COLUMN, Q14_PLANNED_REVENUE_LABEL),
                (Q14_ACTUAL_REVENUE_COLUMN, Q14_ACTUAL_REVENUE_LABEL),
            ],
            "parent_canonical_id": None,
            "source_row": 1,
            "warnings": [],
        }
        raw_columns = (
            "respondent_id",
            Q14_PLANNED_REVENUE_COLUMN,
            Q14_ACTUAL_REVENUE_COLUMN,
        )

        self.assertTrue(_has_double_colon_grid_columns(question, raw_columns))

    def test_q4_employee_band_labels_decoded_to_labels(self) -> None:
        question = {
            "canonical_id": "Q4",
            "raw_id": "Q4",
            "question_text": "How many employees does your company have?",
            "type_hint": "values_range",
            "value_range": (1, 7),
            "options": [
                (1, "Fewer than 250 employees"),
                (2, "250 to 499 employees"),
                (3, "500 to 999 employees"),
                (4, "1,000 to 1,999 employees"),
                (5, "2,000 to 4,999 employees"),
                (6, "5,000 to 9,999 employees"),
                (7, "More than 250,000 employees"),
            ],
            "sub_columns": [],
            "parent_canonical_id": None,
            "source_row": 1,
            "warnings": [],
        }
        decoded = self._decode(
            [
                {"respondent_id": "R1", "Q4": 5},
                {"respondent_id": "R2", "Q4": 4},
                {"respondent_id": "R3", "Q4": 7},
            ],
            question,
        )

        self.assertEqual(
            decoded["Q4"].tolist(),
            [
                "2,000 to 4,999 employees",
                "1,000 to 1,999 employees",
                "More than 250,000 employees",
            ],
        )

    def test_q95_range_labels_decoded_to_labels(self) -> None:
        question = {
            "canonical_id": "Q95",
            "raw_id": "Q95",
            "question_text": "How many opportunities?",
            "type_hint": "values_range",
            "value_range": (1, 5),
            "options": [
                (1, "0"),
                (2, "1-2"),
                (3, "3-5"),
                (4, "6-9"),
                (5, "More than 9"),
            ],
            "sub_columns": [],
            "parent_canonical_id": None,
            "source_row": 1,
            "warnings": [],
        }
        decoded = self._decode(
            [
                {"respondent_id": "R1", "Q95": 3},
                {"respondent_id": "R2", "Q95": 2},
                {"respondent_id": "R3", "Q95": 1},
            ],
            question,
        )

        self.assertEqual(decoded["Q95"].tolist(), ["3-5", "1-2", "0"])

    def test_regular_single_select_still_option_mapped(self) -> None:
        question = {
            "canonical_id": "Q3",
            "raw_id": "Q3",
            "question_text": "Employment status",
            "type_hint": "values_range",
            "value_range": (1, 2),
            "options": [(1, "Full-time"), (2, "Part-time")],
            "sub_columns": [],
            "parent_canonical_id": None,
            "source_row": 1,
            "warnings": [],
        }
        decoded = self._decode(
            [
                {"respondent_id": "R1", "Q3": 1},
                {"respondent_id": "R2", "Q3": 2},
            ],
            question,
        )

        self.assertEqual(decoded["Q3"].tolist(), ["Full-time", "Part-time"])

    def test_regular_multi_select_binary_still_works(self) -> None:
        question = {
            "canonical_id": "Q3",
            "raw_id": "Q3",
            "question_text": "Which segments do you serve?",
            "type_hint": "values_range",
            "value_range": (0, 1),
            "options": [(1, "B2B"), (2, "B2C")],
            "sub_columns": [
                ("Q3: B2B", "B2B"),
                ("Q3: B2C", "B2C"),
            ],
            "parent_canonical_id": None,
            "source_row": 1,
            "warnings": [],
        }
        decoded = self._decode(
            [
                {"respondent_id": "R1", "Q3: B2B": 2, "Q3: B2C": None},
                {"respondent_id": "R2", "Q3: B2B": None, "Q3: B2C": 2},
                {"respondent_id": "R3", "Q3: B2B": 2, "Q3: B2C": None},
            ],
            question,
        )

        self.assertEqual(decoded["Q3: B2B"].dropna().tolist(), [1, 1])
        self.assertEqual(decoded["Q3: B2C"].dropna().tolist(), [1])

    def test_q27_real_winvslag_simulation(self) -> None:
        decoded = self._decode(
            [
                {
                    "respondent_id": "R1",
                    Q27_FIELD_SALES_COLUMN: 3,
                    Q27_INSIDE_SALES_COLUMN: 2,
                },
                {
                    "respondent_id": "R2",
                    Q27_FIELD_SALES_COLUMN: 3,
                    Q27_INSIDE_SALES_COLUMN: 4,
                },
                {
                    "respondent_id": "R3",
                    Q27_FIELD_SALES_COLUMN: None,
                    Q27_INSIDE_SALES_COLUMN: 5,
                },
            ],
            _q27_question(),
        )

        self.assertEqual(decoded[Q27_FIELD_SALES_COLUMN].dropna().tolist(), [3.0, 3.0])
        self.assertEqual(
            decoded[Q27_INSIDE_SALES_COLUMN].dropna().tolist(),
            [2.0, 4.0, 5.0],
        )
        self.assertNotIn(
            "Selected",
            decoded[Q27_FIELD_SALES_COLUMN].astype(str).tolist(),
        )
        self.assertNotIn(
            "More often",
            decoded[Q27_FIELD_SALES_COLUMN].astype(str).tolist(),
        )


if __name__ == "__main__":
    unittest.main()
