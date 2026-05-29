from __future__ import annotations

from pathlib import Path
import unittest


APP_PATH = Path(__file__).resolve().parents[1] / "artifacts" / "survey-insight-engine" / "app.py"


class TestManualCohortUi(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app_text = APP_PATH.read_text(encoding="utf-8")

    def test_manual_mode_hides_outcome_variable_picker(self) -> None:
        self.assertIn("Manual UUID mode hides the outcome variable picker.", self.app_text)
        self.assertIn('app.session_state["seg_mode_radio"] = "manual_uuid"', self.app_text)

    def test_manual_mode_shows_summary_panel(self) -> None:
        self.assertIn("Manual cohort detected:", self.app_text)
        self.assertIn("Manual cohort definition", self.app_text)
        self.assertIn("Overlap (in both winner and laggard lists)", self.app_text)

    def test_manual_mode_validation_blocks_export_on_overlap(self) -> None:
        self.assertIn("manual_cohort_overlap_blocked", self.app_text)
        self.assertIn("Manual cohort overlap detected", self.app_text)
        self.assertIn("Resolve manual cohort overlap before export", self.app_text)

    def test_manual_upload_widget_present(self) -> None:
        self.assertIn("Upload Winners & Laggards xlsx", self.app_text)
        self.assertIn("manual_winners_laggards_upload", self.app_text)

    def test_fallback_info_message_surfaced_in_app(self) -> None:
        self.assertIn("Manual cohort matched against", self.app_text)
        self.assertIn("but uploaded values are", self.app_text)
        self.assertIn("manual_cohort_id_column", self.app_text)


if __name__ == "__main__":
    unittest.main()
