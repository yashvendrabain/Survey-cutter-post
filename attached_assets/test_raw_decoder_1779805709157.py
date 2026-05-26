"""Tests for raw survey data decoding."""

from __future__ import annotations

from pathlib import Path
import tempfile
import unittest

import pandas as pd

from src.raw_decoder import decode_raw_data
from tests.conftest import (
    RAW_DECODER_CSV_PATH,
    RAW_DECODER_NO_ID_CSV_PATH,
    RAW_DECODER_XLSX_PATH,
)


RAW_DECODER_DATA_MAP = {
    "questions": [
        {
            "canonical_id": "record",
            "raw_id": "record",
            "question_text": "Respondent ID",
            "type_hint": None,
            "value_range": None,
            "options": [],
            "sub_columns": [],
            "parent_canonical_id": None,
            "source_row": 1,
            "warnings": [],
        },
        {
            "canonical_id": "Q3",
            "raw_id": "[Q3]",
            "question_text": "Are you currently in a full-time position",
            "type_hint": "values_range",
            "value_range": (1, 2),
            "options": [(1, "Yes"), (2, "No")],
            "sub_columns": [],
            "parent_canonical_id": None,
            "source_row": 2,
            "warnings": [],
        },
        {
            "canonical_id": "Q53",
            "raw_id": "Q53",
            "question_text": "Which of the following challenges...",
            "type_hint": "values_range",
            "value_range": (0, 1),
            "options": [],
            "sub_columns": [("Q53r1", "Knowledge gap"), ("Q53r2", "Alignment gap")],
            "parent_canonical_id": None,
            "source_row": 10,
            "warnings": [],
        },
        {
            "canonical_id": "Q70",
            "raw_id": "[Q70]",
            "question_text": "What % of your pipeline...",
            "type_hint": "values_range",
            "value_range": (0, 100),
            "options": [],
            "sub_columns": [],
            "parent_canonical_id": None,
            "source_row": 27,
            "warnings": [],
        },
        {
            "canonical_id": "Q4r98oe",
            "raw_id": "[Q4r98oe]",
            "question_text": "Other industry",
            "type_hint": "open_text",
            "value_range": None,
            "options": [],
            "sub_columns": [],
            "parent_canonical_id": "Q4",
            "source_row": 35,
            "warnings": [],
        },
        {
            "canonical_id": "vStatus",
            "raw_id": "vStatus",
            "question_text": "Status metadata",
            "type_hint": None,
            "value_range": None,
            "options": [],
            "sub_columns": [],
            "parent_canonical_id": None,
            "source_row": 50,
            "warnings": [],
        },
    ],
    "source_path": "datamap.xlsx",
    "sheet_name": "Sheet1",
    "total_rows_in_sheet": 50,
    "parser_warnings": [],
}


class TestRawDecoder(unittest.TestCase):
    def decode_dataframe(self, dataframe: pd.DataFrame, data_map: dict) -> pd.DataFrame:
        with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as temp_file:
            path = Path(temp_file.name)
        try:
            dataframe.to_csv(path, index=False)
            decoded, _ = decode_raw_data(str(path), data_map)
            return decoded
        finally:
            if path.exists():
                path.unlink()

    def grid_data_map(self) -> dict:
        return {
            "questions": [
                {
                    "canonical_id": "Q38",
                    "raw_id": "[Q38]",
                    "question_text": "Grid rejection test",
                    "type_hint": "values_range",
                    "value_range": (1, 2),
                    "options": [(1, "Selected"), (2, "Not selected")],
                    "sub_columns": [("Q38r1", "Fallback A"), ("Q38r2", "Fallback B")],
                    "parent_canonical_id": None,
                    "source_row": 1,
                    "warnings": [],
                }
            ],
            "source_path": "datamap.xlsx",
            "sheet_name": "Sheet1",
            "total_rows_in_sheet": 1,
            "parser_warnings": [],
        }

    def test_grid_rejection_prefixes_decode_to_binary_selection(self) -> None:
        data_map = self.grid_data_map()
        decoded = self.decode_dataframe(
            pd.DataFrame(
                {
                    "Q38r1": ["Option A", "NO TO: Option A", ""],
                    "Q38r2": ["NO - Option B", "Option B", None],
                }
            ),
            data_map,
        )

        self.assertEqual(decoded["Q38r1"].dropna().tolist(), [1, 0])
        self.assertEqual(decoded["Q38r2"].dropna().tolist(), [0, 1])
        self.assertEqual(data_map["questions"][0]["sub_columns"][0][1], "Option A")
        self.assertEqual(data_map["questions"][0]["sub_columns"][1][1], "Option B")

    def test_grid_all_rejected_column_has_zero_selected_full_denominator(self) -> None:
        decoded = self.decode_dataframe(
            pd.DataFrame(
                {
                    "Q38r1": [
                        "NO TO: Option A",
                        "NOT SELECTED: Option A",
                        "Not selected: Option A",
                    ]
                }
            ),
            self.grid_data_map(),
        )

        self.assertEqual(int(decoded["Q38r1"].sum()), 0)
        self.assertEqual(int(decoded["Q38r1"].notna().sum()), 3)

    def test_grid_template_placeholders_are_excluded(self) -> None:
        decoded = self.decode_dataframe(
            pd.DataFrame(
                {
                    "Q38r1": ["Option A", "$ {Q38.r901.open}", "${Q38.r902.open}"],
                }
            ),
            self.grid_data_map(),
        )

        self.assertEqual(decoded["Q38r1"].dropna().tolist(), [1])

    def test_loads_csv_file(self) -> None:
        dataframe, report = decode_raw_data(
            str(RAW_DECODER_CSV_PATH), RAW_DECODER_DATA_MAP
        )

        self.assertEqual(dataframe.shape, (20, 8))
        self.assertEqual(report.total_rows, 20)
        self.assertEqual(report.total_columns, 8)
        self.assertEqual(report.columns_in_datamap, 7)

    def test_loads_xlsx_file(self) -> None:
        dataframe, report = decode_raw_data(
            str(RAW_DECODER_XLSX_PATH), RAW_DECODER_DATA_MAP
        )

        self.assertEqual(dataframe.shape, (20, 8))
        self.assertEqual(report.total_rows, 20)
        self.assertIn(
            "data loaded from sheet 'Raw' (not the first sheet)",
            report.warnings,
        )

    def test_missing_tokens_replaced_with_na(self) -> None:
        dataframe, _ = decode_raw_data(str(RAW_DECODER_CSV_PATH), RAW_DECODER_DATA_MAP)

        self.assertTrue(pd.isna(dataframe.loc[3, "Q3"]))
        self.assertTrue(pd.isna(dataframe.loc[5, "Q4r98oe"]))
        self.assertTrue(pd.isna(dataframe.loc[6, "Q4r98oe"]))
        self.assertTrue(pd.isna(dataframe.loc[7, "Q4r98oe"]))

    def test_decode_raw_data_returns_labels_not_codes(self) -> None:
        dataframe, _ = decode_raw_data(str(RAW_DECODER_CSV_PATH), RAW_DECODER_DATA_MAP)

        values = dataframe["Q3"].dropna().unique().tolist()
        self.assertIn("Yes", values)
        self.assertIn("No", values)
        self.assertNotIn(1, values)
        self.assertNotIn(2, values)
        self.assertTrue(all(isinstance(value, str) for value in values))

    def test_respondent_id_found_when_present(self) -> None:
        dataframe, report = decode_raw_data(
            str(RAW_DECODER_CSV_PATH), RAW_DECODER_DATA_MAP
        )

        self.assertIn("record", dataframe.columns)
        self.assertEqual(dataframe.loc[0, "record"], "R001")
        self.assertNotIn(
            "no respondent ID column found; generated sequential IDs",
            report.warnings,
        )

    def test_respondent_id_generated_when_absent(self) -> None:
        dataframe, report = decode_raw_data(
            str(RAW_DECODER_NO_ID_CSV_PATH), RAW_DECODER_DATA_MAP
        )

        self.assertIn("respondent_id", dataframe.columns)
        self.assertEqual(int(dataframe.loc[0, "respondent_id"]), 1)
        self.assertEqual(int(dataframe.loc[19, "respondent_id"]), 20)
        self.assertIn(
            "no respondent ID column found; generated sequential IDs",
            report.warnings,
        )

    def test_dtype_coercion_numeric_column(self) -> None:
        dataframe, _ = decode_raw_data(str(RAW_DECODER_CSV_PATH), RAW_DECODER_DATA_MAP)

        self.assertTrue(pd.api.types.is_numeric_dtype(dataframe["Q70"]))
        self.assertEqual(float(dataframe.loc[0, "Q70"]), 10.0)

    def test_coercion_log_populated_when_values_coerced(self) -> None:
        _, report = decode_raw_data(str(RAW_DECODER_CSV_PATH), RAW_DECODER_DATA_MAP)

        q3_logs = [entry for entry in report.coercion_log if entry["column"] == "Q3"]
        self.assertEqual(
            q3_logs,
            [
                {
                    "column": "Q3",
                    "from_type": "string",
                    "to_type": "numeric",
                    "values_coerced": ["abc"],
                    "rows_affected": 1,
                }
            ],
        )

    def test_open_text_column_not_coerced(self) -> None:
        dataframe, _ = decode_raw_data(str(RAW_DECODER_CSV_PATH), RAW_DECODER_DATA_MAP)

        self.assertEqual(dataframe.loc[0, "Q4r98oe"], "note 1")
        self.assertFalse(pd.api.types.is_numeric_dtype(dataframe["Q4r98oe"]))

    def test_out_of_range_pct_computed_correctly(self) -> None:
        _, report = decode_raw_data(str(RAW_DECODER_CSV_PATH), RAW_DECODER_DATA_MAP)

        self.assertEqual(report.per_column_out_of_range_pct["Q70"], 0.1)
        self.assertIn("column Q70 has 10.0% out-of-range values", report.warnings)

    def test_columns_not_in_datamap_identified(self) -> None:
        _, report = decode_raw_data(str(RAW_DECODER_CSV_PATH), RAW_DECODER_DATA_MAP)

        self.assertEqual(report.columns_not_in_datamap, ("extra_col",))

    def test_high_missing_warning_emitted(self) -> None:
        _, report = decode_raw_data(str(RAW_DECODER_CSV_PATH), RAW_DECODER_DATA_MAP)

        self.assertEqual(report.per_column_missing_pct["Q4r98oe"], 0.65)
        self.assertIn("column Q4r98oe has 65.0% missing values", report.warnings)


if __name__ == "__main__":
    unittest.main()
