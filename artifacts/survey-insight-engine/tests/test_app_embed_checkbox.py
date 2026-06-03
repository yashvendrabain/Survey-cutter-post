from __future__ import annotations

from pathlib import Path
import unittest


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"


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

    def test_ai_enhancements_checkbox_default_on(self) -> None:
        self.assertIn("Generate AI themes & labels (adds ~30s)", self.app_text)
        self.assertIn('"skip_ai_enhancements": False', self.app_text)
        self.assertIn(
            'app.session_state["skip_ai_enhancements"] = not generate_ai_enhancements',
            self.app_text,
        )

    def test_stage4_ai_calls_parallel_with_separate_caches(self) -> None:
        run_pipeline_start = self.app_text.index("def _run_pipeline(")
        stage4_start = self.app_text.index("def _run_stage4_ai_enhancements", run_pipeline_start)
        stage5_start = self.app_text.index("workbook_custom_filter_count", stage4_start)
        stage4_text = self.app_text[stage4_start:stage5_start]

        self.assertIn("ThreadPoolExecutor(max_workers=3)", stage4_text)
        self.assertIn("theme_cache: dict[str, Any] = {}", stage4_text)
        self.assertIn("demo_cache: dict[str, Any] = {}", stage4_text)
        self.assertIn("label_cache: dict[str, Any] = {}", stage4_text)
        self.assertIn("themes_future = executor.submit", stage4_text)
        self.assertIn("demo_future = executor.submit", stage4_text)
        self.assertIn("labels_future = executor.submit", stage4_text)
        self.assertNotIn("cache=_INSIGHT_CACHE", stage4_text)

    def test_stage4_ai_skip_uses_deterministic_fallbacks(self) -> None:
        self.assertIn('app.session_state.get("skip_ai_enhancements", False)', self.app_text)
        self.assertIn("_fallback_stage4_themes(", self.app_text)
        self.assertIn("_fallback_stage4_demo_priority(", self.app_text)
        self.assertIn("_fallback_stage4_short_labels(", self.app_text)


if __name__ == "__main__":
    unittest.main()
