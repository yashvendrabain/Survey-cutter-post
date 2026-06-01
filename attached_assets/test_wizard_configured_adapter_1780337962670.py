"""Tests for the wizard-configured parser adapter."""

from __future__ import annotations

import unittest

import pandas as pd
from openpyxl import Workbook

from src.adapters.wizard_configured import WizardConfig, WizardConfiguredAdapter


def _config(**overrides) -> WizardConfig:
    values = {
        "raw_data_sheet_name": "Responses",
        "data_map_sheet_name": "Data Map",
        "respondent_id_column": "uuid",
        "question_id_pattern": r"^Q\d+",
        "sub_column_separator": "r",
        "option_code_position": "column_b",
        "section_prefixes": ("Q",),
        "config_name": None,
        "helper_columns": tuple(),
    }
    values.update(overrides)
    return WizardConfig(**values)


def _workbook(rows: list[list[object]]) -> Workbook:
    workbook = Workbook()
    raw = workbook.active
    raw.title = "Responses"
    raw.append(["uuid", "Q1", "Q2r1", "Q2r2", "helper_flag"])
    raw.append(["a", 1, 1, 0, "x"])
    sheet = workbook.create_sheet("Data Map")
    for row in rows:
        sheet.append(row)
    return workbook


class TestWizardConfiguredAdapter(unittest.TestCase):
    def test_wizard_config_constructs_with_minimal_fields(self) -> None:
        cfg = _config()

        self.assertEqual(cfg.raw_data_sheet_name, "Responses")
        self.assertEqual(cfg.section_prefixes, ("Q",))

    def test_wizard_config_validates_required_fields(self) -> None:
        with self.assertRaisesRegex(ValueError, "raw_data_sheet_name"):
            _config(raw_data_sheet_name="")

    def test_wizard_adapter_parses_with_explicit_config(self) -> None:
        workbook = _workbook([["Q1", "Employment"], [1, "Full-time"], [2, "Part-time"]])
        raw_df = pd.DataFrame(columns=["uuid", "Q1"])

        parsed = WizardConfiguredAdapter(_config()).parse(workbook, raw_df)

        self.assertEqual(parsed["questions"][0]["canonical_id"], "Q1")
        self.assertEqual(parsed["questions"][0]["options"], [(1, "Full-time"), (2, "Part-time")])

    def test_wizard_adapter_handles_custom_question_id_pattern(self) -> None:
        workbook = _workbook([["Item_1", "Age"], [1, "18-34"], [2, "35+"]])
        raw_df = pd.DataFrame(columns=["uuid", "Item_1"])

        parsed = WizardConfiguredAdapter(
            _config(question_id_pattern=r"^Item_\d+", sub_column_separator="none")
        ).parse(workbook, raw_df)

        self.assertEqual(parsed["questions"][0]["canonical_id"], "Item_1")

    def test_wizard_adapter_handles_custom_subcolumn_separator(self) -> None:
        workbook = _workbook([["Item_1", "Rate each item"], [1, "Yes"], [2, "No"]])
        raw_df = pd.DataFrame(columns=["uuid", "Item_1_sub_1", "Item_1_sub_2"])

        parsed = WizardConfiguredAdapter(
            _config(question_id_pattern=r"^Item_\d+", sub_column_separator="_sub_")
        ).parse(workbook, raw_df)

        self.assertEqual(parsed["questions"][0]["sub_columns"][0], ("Item_1_sub_1", "1"))

    def test_wizard_adapter_skips_configured_helper_columns(self) -> None:
        workbook = _workbook([["Q2", "Pick all"], [1, "Selected"], [0, "Not selected"]])
        raw_df = pd.DataFrame(columns=["uuid", "Q2r1", "Q2r2"])

        parsed = WizardConfiguredAdapter(_config(helper_columns=("Q2r2",))).parse(workbook, raw_df)

        self.assertEqual(parsed["questions"][0]["sub_columns"], [("Q2r1", "1")])

    def test_wizard_adapter_supports_section_prefixes(self) -> None:
        workbook = _workbook([["D1", "Region"], [1, "NA"], [2, "EU"]])
        raw_df = pd.DataFrame(columns=["uuid", "D1"])

        parsed = WizardConfiguredAdapter(
            _config(question_id_pattern=r"^Q\d+", section_prefixes=("Q", "D"))
        ).parse(workbook, raw_df)

        self.assertEqual(parsed["questions"][0]["canonical_id"], "D1")

    def test_wizard_adapter_returns_normalized_data_map_shape(self) -> None:
        workbook = _workbook([["Q1", "Employment"], [1, "Full-time"]])
        raw_df = pd.DataFrame(columns=["uuid", "Q1"])

        parsed = WizardConfiguredAdapter(_config()).parse(workbook, raw_df)

        self.assertEqual(
            set(parsed),
            {"questions", "source_path", "sheet_name", "total_rows_in_sheet", "parser_warnings"},
        )


if __name__ == "__main__":
    unittest.main()
