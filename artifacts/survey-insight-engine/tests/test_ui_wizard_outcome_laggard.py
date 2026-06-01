from __future__ import annotations

from pathlib import Path
import unittest


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"


class TestOutcomeLaggardUi(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app_text = APP_PATH.read_text(encoding="utf-8")

    def test_laggard_cohort_ui_copy_is_present(self) -> None:
        self.assertIn("Laggard cohort", self.app_text)
        self.assertIn("Laggard option values", self.app_text)
        self.assertIn("Laggard threshold", self.app_text)

    def test_override_outcome_checkbox_and_dropdown_are_present(self) -> None:
        self.assertIn("Override outcome question for laggards", self.app_text)
        self.assertIn("laggard_outcome_variable_selector", self.app_text)

    def test_grid_override_subquestion_dropdown_is_present(self) -> None:
        self.assertIn("Laggard grid sub-question", self.app_text)
        self.assertIn("laggard_outcome_sub_question_selector", self.app_text)

    def test_outcome_segmented_download_button_is_present(self) -> None:
        self.assertIn("Outcome Segmented Workbook (Winners vs Laggards)", self.app_text)
        self.assertIn("export_winners_vs_laggards_workbook", self.app_text)

    def test_protective_app_container_and_run_pipeline_import_position(self) -> None:
        self.assertIn("with app.container():", self.app_text)
        run_pipeline = self.app_text.index("def _run_pipeline(")
        body = self.app_text[run_pipeline:self.app_text.index("status.update", run_pipeline)]
        self.assertIn("app = _require_streamlit()", body)


if __name__ == "__main__":
    unittest.main()
