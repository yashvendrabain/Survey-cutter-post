from __future__ import annotations

from pathlib import Path
import uuid
import unittest

import pandas as pd

from src.raw_decoder import decode_raw_data


def _write_raw_csv(rows: list[dict[str, object | None]]) -> Path:
    scratch_root = Path.cwd() / "outputs" / "test_tmp"
    scratch_root.mkdir(parents=True, exist_ok=True)
    path = scratch_root / f"raw_grid_{uuid.uuid4().hex}.csv"
    pd.DataFrame(rows).to_csv(path, index=False)
    return path


def _safe_unlink(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)
    except PermissionError:
        pass


def _grid_single_select_data_map() -> dict:
    return {
        "questions": [
            {
                "canonical_id": "Q27",
                "raw_id": "Q27",
                "question_text": "How has usage changed?",
                "type_hint": "values_range",
                "value_range": (1, 5),
                "options": [
                    (1, "More often"),
                    (2, "Less often"),
                    (3, "Same"),
                    (4, "Not applicable"),
                    (5, "Don't know"),
                ],
                "sub_columns": [
                    ("Q27r1", "International offices"),
                    ("Q27r2", "General venture partnerships"),
                ],
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


def _grid_rated_data_map() -> dict:
    return {
        "questions": [
            {
                "canonical_id": "Q30",
                "raw_id": "Q30",
                "question_text": "Rate the following",
                "type_hint": "values_range",
                "value_range": (0, 10),
                "options": [],
                "sub_columns": [("Q30r1c1", "Winner")],
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


def _multi_select_data_map() -> dict:
    return {
        "questions": [
            {
                "canonical_id": "Q53",
                "raw_id": "Q53",
                "question_text": "Which challenges apply?",
                "type_hint": "values_range",
                "value_range": (0, 1),
                "options": [],
                "sub_columns": [
                    ("Q53r1", "Knowledge gap"),
                    ("Q53r2", "Alignment gap"),
                ],
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


class TestRawDecoderGridPreservation(unittest.TestCase):
    def test_grid_single_select_numeric_value_preserved(self) -> None:
        raw_path = _write_raw_csv(
            [
                {
                    "respondent_id": "R1",
                    "Q27r1": 3,
                    "Q27r2": 2,
                }
            ]
        )
        try:
            decoded, _report = decode_raw_data(
                str(raw_path),
                _grid_single_select_data_map(),
            )
        finally:
            _safe_unlink(raw_path)

        self.assertEqual(float(decoded.loc[0, "Q27r1"]), 3.0)
        self.assertEqual(float(decoded.loc[0, "Q27r2"]), 2.0)
        self.assertNotEqual(decoded.loc[0, "Q27r1"], "Selected")
        self.assertNotEqual(decoded.loc[0, "Q27r1"], "Same")

    def test_grid_rated_numeric_value_preserved(self) -> None:
        raw_path = _write_raw_csv(
            [
                {
                    "respondent_id": "R1",
                    "Q30r1c1": 7.5,
                }
            ]
        )
        try:
            decoded, _report = decode_raw_data(str(raw_path), _grid_rated_data_map())
        finally:
            _safe_unlink(raw_path)

        self.assertEqual(float(decoded.loc[0, "Q30r1c1"]), 7.5)

    def test_multi_select_binary_still_works(self) -> None:
        raw_path = _write_raw_csv(
            [
                {"respondent_id": "R1", "Q53r1": "Choice A", "Q53r2": None},
                {"respondent_id": "R2", "Q53r1": None, "Q53r2": "Choice B"},
                {"respondent_id": "R3", "Q53r1": "Choice A", "Q53r2": None},
            ]
        )
        try:
            decoded, _report = decode_raw_data(str(raw_path), _multi_select_data_map())
        finally:
            _safe_unlink(raw_path)

        self.assertEqual(int(decoded["Q53r1"].notna().sum()), 2)
        self.assertEqual(int(decoded["Q53r2"].notna().sum()), 1)

    def test_blank_cells_remain_none(self) -> None:
        raw_path = _write_raw_csv(
            [
                {
                    "respondent_id": "R1",
                    "Q27r1": None,
                    "Q27r2": "",
                }
            ]
        )
        try:
            decoded, _report = decode_raw_data(
                str(raw_path),
                _grid_single_select_data_map(),
            )
        finally:
            _safe_unlink(raw_path)

        self.assertTrue(pd.isna(decoded.loc[0, "Q27r1"]))
        self.assertTrue(pd.isna(decoded.loc[0, "Q27r2"]))


if __name__ == "__main__":
    unittest.main()
