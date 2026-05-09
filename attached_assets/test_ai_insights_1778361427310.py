"""Tests for the AI insight layer."""

from __future__ import annotations

import os
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from config import PORTKEY_PREMIUM_MODEL
from src.ai_insights import generate_insight
from src.models import InsightResult


def _single_cut_payload() -> dict:
    return {
        "table_kind": "single_cut",
        "question_id": "Q1",
        "question_text": "Revenue growth in 2023 vs 2022",
        "valid_n": 100,
        "missing_n": 0,
        "filters_applied": [],
        "rows": [
            {"label": "Low growth", "count": 25, "rate": 0.25},
            {"label": "Moderate growth", "count": 50, "rate": 0.50},
            {"label": "High growth", "count": 25, "rate": 0.25},
        ],
        "summary": {},
    }


def _fake_response(content: str, total_tokens: int = 21) -> SimpleNamespace:
    return SimpleNamespace(
        choices=[
            SimpleNamespace(
                message=SimpleNamespace(content=content),
            )
        ],
        usage=SimpleNamespace(total_tokens=total_tokens),
    )


class TestAiInsights(unittest.TestCase):
    def test_template_fallback_when_api_key_missing(self) -> None:
        with patch.dict(os.environ, {}, clear=True):
            result = generate_insight(
                _single_cut_payload(),
                table_kind="single_cut",
            )

        self.assertTrue(result.was_template)
        self.assertTrue(result.title)
        self.assertTrue(result.insight)
        self.assertIn("PORTKEY_API_KEY", result.error_message)

    def test_template_fallback_for_single_cut_picks_top_option(self) -> None:
        with patch.dict(os.environ, {}, clear=True):
            result = generate_insight(
                _single_cut_payload(),
                table_kind="single_cut",
            )

        self.assertTrue(result.was_template)
        self.assertIn("Moderate growth", result.insight)
        self.assertIn("50", result.insight)

    def test_template_fallback_with_empty_rows(self) -> None:
        payload = _single_cut_payload()
        payload["rows"] = []

        result = generate_insight(payload, table_kind="single_cut")

        self.assertTrue(result.was_template)
        self.assertIn("No respondents", result.insight)

    def test_template_fallback_includes_filter_description(self) -> None:
        payload = _single_cut_payload()
        payload["filters_applied"] = ["Region == APAC"]

        with patch.dict(os.environ, {}, clear=True):
            result = generate_insight(payload, table_kind="single_cut")

        self.assertIn("Region == APAC", result.insight)

    def test_insight_result_validation_rejects_empty_title(self) -> None:
        with self.assertRaises(ValueError):
            InsightResult(title="", insight="Useful insight")

    def test_insight_result_validation_rejects_empty_insight(self) -> None:
        with self.assertRaises(ValueError):
            InsightResult(title="Useful title", insight="")

    def test_generate_insight_handles_api_exception(self) -> None:
        with patch.dict(os.environ, {"PORTKEY_API_KEY": "test-key"}):
            with patch("src.ai_insights.OpenAI", side_effect=RuntimeError("boom")):
                result = generate_insight(
                    _single_cut_payload(),
                    table_kind="single_cut",
                )

        self.assertTrue(result.was_template)
        self.assertIn("API call failed", result.error_message)
        self.assertIn("RuntimeError", result.error_message)

    def test_generate_insight_handles_invalid_json_response(self) -> None:
        client = Mock()
        client.chat.completions.create.return_value = _fake_response("not json")

        with patch.dict(os.environ, {"PORTKEY_API_KEY": "test-key"}):
            with patch("src.ai_insights.OpenAI", return_value=client):
                result = generate_insight(
                    _single_cut_payload(),
                    table_kind="single_cut",
                )

        self.assertTrue(result.was_template)
        self.assertIn("Response parse failed", result.error_message)

    def test_generate_insight_handles_missing_title_or_insight(self) -> None:
        client = Mock()
        client.chat.completions.create.return_value = _fake_response(
            '{"title": "", "insight": "Something happened."}'
        )

        with patch.dict(os.environ, {"PORTKEY_API_KEY": "test-key"}):
            with patch("src.ai_insights.OpenAI", return_value=client):
                result = generate_insight(
                    _single_cut_payload(),
                    table_kind="single_cut",
                )

        self.assertTrue(result.was_template)
        self.assertIn("title or insight empty", result.error_message)

    def test_payload_question_text_used_for_title_hint(self) -> None:
        with patch.dict(os.environ, {}, clear=True):
            result = generate_insight(
                _single_cut_payload(),
                table_kind="single_cut",
            )

        self.assertEqual(result.title, "Revenue growth in 2023 vs 2022")

    def test_premium_flag_selects_gpt_4o_model(self) -> None:
        client = Mock()
        client.chat.completions.create.return_value = _fake_response(
            '{"title": "Revenue growth pattern", "insight": "The table shows moderate growth as the largest response."}'
        )

        with patch.dict(os.environ, {"PORTKEY_API_KEY": "test-key"}):
            with patch("src.ai_insights.OpenAI", return_value=client):
                result = generate_insight(
                    _single_cut_payload(),
                    table_kind="single_cut",
                    use_premium=True,
                )

        self.assertFalse(result.was_template)
        self.assertEqual(result.model_used, PORTKEY_PREMIUM_MODEL)
        self.assertEqual(
            client.chat.completions.create.call_args.kwargs["model"],
            PORTKEY_PREMIUM_MODEL,
        )


if __name__ == "__main__":
    unittest.main()
