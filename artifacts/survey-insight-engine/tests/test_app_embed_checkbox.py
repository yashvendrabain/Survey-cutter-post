from __future__ import annotations

from pathlib import Path
import unittest


APP_PATH = Path(__file__).resolve().parents[1] / "artifacts" / "survey-insight-engine" / "app.py"


class TestAppEmbedCheckbox(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app_text = APP_PATH.read_text(encoding="utf-8")

    def test_embed_checkbox_default_off(self) -> None:
        self.assertIn(
            "Include original raw data and data map as sheets in exported workbooks",
            self.app_text,
        )
        self.assertIn('key="wizard_embed_input_files"', self.app_text)
        self.assertIn("value=False", self.app_text)

    def test_embed_checkbox_state_persists(self) -> None:
        self.assertIn('"wizard_embed_input_files": False', self.app_text)
        self.assertIn('app.session_state["input_file_embed_sources"]', self.app_text)
        self.assertIn("embed_input_files=app.session_state.get(\"wizard_embed_input_files\", False)", self.app_text)
        self.assertIn("input_file_sources=app.session_state.get(\"input_file_embed_sources\")", self.app_text)


if __name__ == "__main__":
    unittest.main()
