"""Tests for wizard diagnostic text."""

from __future__ import annotations

import importlib.util
from pathlib import Path
import unittest


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"


def load_app_module():
    spec = importlib.util.spec_from_file_location("survey_app_wizard_diag", APP_PATH)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def _diagnostic() -> str:
    app_module = load_app_module()
    return app_module._build_wizard_diagnostic_panel_text(
        {
            "raw_data_sheet_name": "Responses",
            "data_map_sheet_name": "Datamap_Sheet1",
            "question_id_pattern": r"^Q\d+",
            "sub_column_separator": "r",
            "helper_columns": ("qc_flag", "straightline_count"),
        },
        "0 questions parsed",
        raw_columns=["Item_1", "Item_2_sub_1", "qc_flag"],
        first_column_values=["Item_1: What is your age?", "Item_2: What is your gender?"],
    )


class TestWizardDiagnosticPanel(unittest.TestCase):
    def test_diagnostic_panel_shows_question_id_pattern_mismatch(self) -> None:
        diagnostic = _diagnostic()

        self.assertIn("0 rows matching question ID pattern: ^Q\\d+", diagnostic)
        self.assertIn("Item_1: What is your age?", diagnostic)

    def test_diagnostic_panel_shows_subcolumn_pattern_mismatch(self) -> None:
        diagnostic = _diagnostic()

        self.assertIn("Sub-column pattern: r - no matches found", diagnostic)
        self.assertIn("Item_2_sub_1", diagnostic)

    def test_diagnostic_panel_suggests_alternative_patterns(self) -> None:
        diagnostic = _diagnostic()

        self.assertIn("Suggested pattern: ^Item_\\d+", diagnostic)
        self.assertIn("Suggested separator: '_sub_'", diagnostic)

    def test_diagnostic_panel_shows_helper_column_skips(self) -> None:
        diagnostic = _diagnostic()

        self.assertIn("Helper columns skipped: qc_flag, straightline_count", diagnostic)


if __name__ == "__main__":
    unittest.main()
