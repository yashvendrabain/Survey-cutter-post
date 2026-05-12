"""Tests for PPT-ready AI insight headline generation."""

from __future__ import annotations

import json
import os
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from src.ai_insights import (
    TEMPLATE_INSIGHT,
    _extract_headline,
    _format_differentiator_payload,
    _format_single_cut_payload,
    _format_winner_profile_payload,
    _payload_hash,
    generate_insight,
)
from src.models import InsightResult


def _single_cut_payload() -> dict:
    return {
        "question_text": "Revenue growth in 2023 vs 2022",
        "question_type": "single_select",
        "valid_n": 100,
        "distribution": {
            "Low growth": {"count": 25, "rate": 0.25},
            "Moderate growth": {"count": 50, "rate": 0.50},
            "High growth": {"count": 25, "rate": 0.25},
        },
        "top_option": "Moderate growth",
        "top_option_pct": "50.0%",
    }


def _differentiator_payload() -> dict:
    return {
        "question_text": "Investment in GTM technology",
        "top_option": "High investment",
        "winner_rate": 0.673,
        "loser_rate": 0.291,
        "lift": 2.31,
        "cramers_v": 0.292,
        "winner_n": 75,
        "loser_n": 72,
    }


def _winner_profile_payload() -> dict:
    return {
        "winner_label": "Winner",
        "loser_label": "Laggard",
        "winner_n": 75,
        "loser_n": 72,
        "traits": [
            {
                "question_id": "Q_GTM",
                "option_label": "High investment",
                "winner_rate": 0.673,
                "loser_rate": 0.291,
                "lift": 2.31,
                "rate_gap": 0.382,
            }
        ],
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


def _mock_openai_client(headline: str) -> Mock:
    client = Mock()
    client.chat.completions.create.return_value = _fake_response(headline)
    return client


class TestAiInsights(unittest.TestCase):
    def test_generate_insight_returns_insight_result(self) -> None:
        with patch.dict(os.environ, {}, clear=True):
            result = generate_insight(_single_cut_payload(), table_kind="single_cut")

        self.assertIsInstance(result, InsightResult)
        self.assertEqual(result.title, "")

    def test_generate_insight_has_non_empty_insight_string(self) -> None:
        with patch.dict(os.environ, {}, clear=True):
            result = generate_insight(_single_cut_payload(), table_kind="single_cut")

        self.assertIsInstance(result.insight, str)
        self.assertTrue(result.insight)

    def test_generate_insight_template_when_api_key_unavailable(self) -> None:
        with patch.dict(os.environ, {}, clear=True):
            result = generate_insight(_single_cut_payload(), table_kind="single_cut")

        self.assertTrue(result.was_template)
        self.assertEqual(result.title, "")
        self.assertEqual(result.insight, TEMPLATE_INSIGHT)

    def test_generate_insight_valid_api_response_is_not_template(self) -> None:
        headline = "Respondents report 50% moderate growth in 2023"
        client = _mock_openai_client(headline)

        with patch.dict(os.environ, {"PORTKEY_API_KEY": "test-key"}):
            with patch("src.ai_insights.OpenAI", return_value=client):
                result = generate_insight(_single_cut_payload(), table_kind="single_cut")

        self.assertFalse(result.was_template)
        self.assertEqual(result.title, "")
        self.assertEqual(result.insight, headline)
        self.assertEqual(client.chat.completions.create.call_count, 1)

    def test_cache_hit_returns_same_object_without_second_api_call(self) -> None:
        headline = "Respondents report 50% moderate growth in 2023"
        client = _mock_openai_client(headline)
        cache: dict = {}

        with patch.dict(os.environ, {"PORTKEY_API_KEY": "test-key"}):
            with patch("src.ai_insights.OpenAI", return_value=client):
                first = generate_insight(
                    _single_cut_payload(),
                    table_kind="single_cut",
                    cache=cache,
                )
                second = generate_insight(
                    _single_cut_payload(),
                    table_kind="single_cut",
                    cache=cache,
                )

        self.assertIs(first, second)
        self.assertEqual(client.chat.completions.create.call_count, 1)

    def test_payload_hash_is_deterministic(self) -> None:
        payload = _single_cut_payload()

        first = _payload_hash("single_cut", payload)
        second = _payload_hash("single_cut", dict(reversed(payload.items())))

        self.assertEqual(first, second)

    def test_format_differentiator_payload_returns_expected_json_keys(self) -> None:
        formatted = _format_differentiator_payload(_differentiator_payload())
        parsed = json.loads(formatted)

        self.assertEqual(parsed["winner_selection_rate"], "67.3%")
        self.assertEqual(parsed["laggard_selection_rate"], "29.1%")
        self.assertEqual(parsed["lift_ratio"], "2.31x")
        self.assertEqual(parsed["association_strength"], "Cramers V = 0.292")

    def test_format_winner_profile_payload_returns_defining_traits(self) -> None:
        formatted = _format_winner_profile_payload(_winner_profile_payload())
        parsed = json.loads(formatted)

        self.assertEqual(parsed["analysis_type"], "winner_profile_summary")
        self.assertIsInstance(parsed["defining_traits"], list)
        self.assertEqual(parsed["defining_traits"][0]["question"], "Q_GTM")

    def test_format_single_cut_payload_returns_valid_json(self) -> None:
        formatted = _format_single_cut_payload(_single_cut_payload())
        parsed = json.loads(formatted)

        self.assertEqual(parsed["analysis_type"], "single_question_distribution")
        self.assertEqual(parsed["question"], "Revenue growth in 2023 vs 2022")
        self.assertEqual(parsed["top_option_pct"], "50.0%")

    def test_template_fallback_returns_template_constant_sentence(self) -> None:
        with patch.dict(os.environ, {}, clear=True):
            result = generate_insight(_single_cut_payload(), table_kind="single_cut")

        self.assertEqual(result.insight, TEMPLATE_INSIGHT)
        self.assertEqual(result.insight.count("."), 1)

    def test_extract_headline_strips_quotes_and_whitespace(self) -> None:
        headline = _extract_headline('  "Winners report 67% adoption"  ')

        self.assertEqual(headline, "Winners report 67% adoption")

    def test_generate_insight_without_cache_calls_api_directly_each_time(self) -> None:
        headline = "Respondents report 50% moderate growth in 2023"
        client = _mock_openai_client(headline)

        with patch.dict(os.environ, {"PORTKEY_API_KEY": "test-key"}):
            with patch("src.ai_insights.OpenAI", return_value=client):
                first = generate_insight(
                    _single_cut_payload(),
                    table_kind="single_cut",
                    cache=None,
                )
                second = generate_insight(
                    _single_cut_payload(),
                    table_kind="single_cut",
                    cache=None,
                )

        self.assertIsNot(first, second)
        self.assertEqual(client.chat.completions.create.call_count, 2)


if __name__ == "__main__":
    unittest.main()
