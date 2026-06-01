"""Text-introspection checks for the survey format wizard UI."""

from __future__ import annotations

from pathlib import Path
import unittest


APP_PATH = Path(__file__).resolve().parents[1] / "artifacts" / "survey-insight-engine" / "app.py"


class TestWizardAppUi(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app_text = APP_PATH.read_text(encoding="utf-8")

    def test_wizard_section_appears_in_app_py(self) -> None:
        self.assertIn("Survey format wizard", self.app_text)
        self.assertIn("We couldn't automatically detect this survey's format", self.app_text)

    def test_wizard_has_7_step_structure(self) -> None:
        for label in (
            "Identify sheets",
            "Respondent ID column",
            "Question ID format",
            "Sub-column separator",
            "Option code location",
            "Section prefixes",
            "Helper / metadata columns",
        ):
            self.assertIn(label, self.app_text)

    def test_wizard_preview_section_present(self) -> None:
        self.assertIn("Wizard configured. Detected:", self.app_text)
        self.assertIn("Proceed to analysis", self.app_text)

    def test_wizard_diagnostic_section_present(self) -> None:
        self.assertIn("Could not parse this survey with the wizard configuration", self.app_text)
        self.assertIn("Adjust wizard and retry", self.app_text)

    def test_wizard_save_draft_path_present(self) -> None:
        self.assertIn("/tmp/draft_adapters", self.app_text)
        self.assertIn("Save as draft adapter", self.app_text)


if __name__ == "__main__":
    unittest.main()
