"""Tests for PPT-ready AI insight headline generation."""

from __future__ import annotations

import json
import os
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from src.ai_insights import (
    TEMPLATE_INSIGHT,
    _build_allowed_numbers,
    _extract_headline,
    _extract_numbers_from_text,
    _format_differentiator_payload,
    _format_differentiator_table_payload,
    _format_single_cut_payload,
    _format_winner_profile_payload,
    _format_winner_profile_trait_payload,
    _payload_hash,
    _validate_numbers,
    categorize_demographic_questions,
    categorize_questions_into_themes,
    generate_insight,
    generate_outlier_insight,
    generate_short_labels,
    generate_table_insight,
)
from src.models import InsightResult


def _single_cut_payload() -> dict:
    return {
        "question_text": "Revenue growth outcome",
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


def _differentiator_table_payload() -> dict:
    return {
        "outcome_question_id": "Q_OUTCOME",
        "winner_label": "Winner",
        "loser_label": "Laggard",
        "winner_n": 75,
        "loser_n": 72,
        "differentiators": [
            {
                "question_text": "Investment in GTM technology",
                "top_option_label": "High investment",
                "winner_rate": 0.673,
                "loser_rate": 0.291,
                "lift": 2.31,
                "cramers_v": 0.292,
            },
            {
                "question_text": "Inside sales usage",
                "top_option_label": "Uses inside sales",
                "winner_rate": 0.81,
                "loser_rate": 0.12,
                "lift": 6.75,
                "cramers_v": 0.41,
            },
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
        headline = "Respondents report 50% moderate growth"
        client = _mock_openai_client(headline)

        with patch.dict(os.environ, {"PORTKEY_API_KEY": "test-key"}):
            with patch("src.ai_insights.OpenAI", return_value=client):
                result = generate_insight(_single_cut_payload(), table_kind="single_cut")

        self.assertFalse(result.was_template)
        self.assertEqual(result.title, "")
        self.assertEqual(result.insight, headline)
        self.assertEqual(client.chat.completions.create.call_count, 1)

    def test_cache_hit_returns_same_object_without_second_api_call(self) -> None:
        headline = "Respondents report 50% moderate growth"
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
        self.assertEqual(parsed["lift_ratio_rounded"], "2.3x")
        self.assertIn("allowed_numbers", parsed)

    def test_format_winner_profile_payload_returns_defining_traits(self) -> None:
        formatted = _format_winner_profile_payload(_winner_profile_payload())
        parsed = json.loads(formatted)

        self.assertEqual(parsed["analysis_type"], "winner_profile_summary")
        self.assertIsInstance(parsed["defining_traits"], list)
        self.assertEqual(parsed["defining_traits"][0]["question"], "Q_GTM")
        self.assertIn("allowed_numbers", parsed)

    def test_format_winner_profile_trait_payload(self) -> None:
        payload = {
            "question_text": "Revenue change",
            "option_label": "3-5% growth",
            "winner_rate": 0.414,
            "loser_rate": 0.167,
            "lift": 2.48,
            "rate_gap": 0.247,
            "winner_label": "Winner",
            "laggard_label": "Laggard",
            "laggard_top_option_label": "Flat/no growth",
            "laggard_top_option_winner_rate": 0.08,
            "laggard_top_option_loser_rate": 0.35,
        }

        formatted = _format_winner_profile_trait_payload(payload)
        parsed = json.loads(formatted)

        self.assertEqual(parsed["winners_top_option"], "3-5% growth")
        self.assertEqual(parsed["laggards_top_option"], "Flat/no growth")
        self.assertIn("selection_share_contract", parsed)
        self.assertIn("proportion", parsed["selection_share_contract"])
        self.assertIn("forbidden_framing", parsed)
        self.assertIn("allocated", parsed["forbidden_framing"])
        self.assertIn(41, parsed["allowed_numbers"])
        self.assertIn(35, parsed["allowed_numbers"])

    def test_format_single_cut_payload_returns_valid_json(self) -> None:
        formatted = _format_single_cut_payload(_single_cut_payload())
        parsed = json.loads(formatted)

        self.assertEqual(parsed["analysis_type"], "single_question_distribution")
        self.assertEqual(parsed["question"], "Revenue growth outcome")
        self.assertEqual(parsed["top_option_pct"], "50.0%")
        self.assertIn(50, parsed["allowed_numbers"])

    def test_template_fallback_returns_template_constant_sentence(self) -> None:
        with patch.dict(os.environ, {}, clear=True):
            result = generate_insight(_single_cut_payload(), table_kind="single_cut")

        self.assertEqual(result.insight, TEMPLATE_INSIGHT)
        self.assertEqual(result.insight.count("."), 1)

    def test_extract_headline_strips_quotes_and_whitespace(self) -> None:
        headline = _extract_headline('  "Winners report 67% adoption"  ')

        self.assertEqual(headline, "Winners report 67% adoption")

    def test_generate_insight_without_cache_calls_api_directly_each_time(self) -> None:
        headline = "Respondents report 50% moderate growth"
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

    def test_validate_numbers_pass(self) -> None:
        is_valid, invalid = _validate_numbers("Winners report 67% adoption", [67])

        self.assertTrue(is_valid)
        self.assertEqual(invalid, [])

    def test_validate_numbers_fail(self) -> None:
        is_valid, invalid = _validate_numbers("Winners report 68% adoption", [67])

        self.assertFalse(is_valid)
        self.assertEqual(invalid, [68.0])

    def test_validate_numbers_tolerance(self) -> None:
        exact_valid, _invalid_exact = _validate_numbers("Winners report 67.3% adoption", [67.3])
        rounded_valid, _invalid_rounded = _validate_numbers("Winners report 67% adoption", [67.3])

        self.assertTrue(exact_valid)
        self.assertTrue(rounded_valid)

    def test_extract_numbers(self) -> None:
        self.assertEqual(
            _extract_numbers_from_text("67% of Winners vs 2.3x Laggards"),
            [67.0, 2.3],
        )

    def test_build_allowed_numbers_deduplication(self) -> None:
        self.assertEqual(_build_allowed_numbers([67.34, 67.34]), [67, 67.3, 67.34])

    def test_call_api_falls_back_on_hallucination(self) -> None:
        client = _mock_openai_client("Respondents report 51% moderate growth")

        with patch.dict(os.environ, {"PORTKEY_API_KEY": "test-key"}):
            with patch("src.ai_insights.OpenAI", return_value=client):
                result = generate_insight(_single_cut_payload(), table_kind="single_cut")

        self.assertTrue(result.was_template)
        self.assertEqual(result.insight, TEMPLATE_INSIGHT)
        self.assertIn("Hallucinated numbers detected", result.error_message)

    def test_generate_table_insight(self) -> None:
        headline = "Winners outpace Laggards across 2 differentiators, led by 67% vs 29%"
        client = _mock_openai_client(headline)

        with patch.dict(os.environ, {"PORTKEY_API_KEY": "test-key"}):
            with patch("src.ai_insights.OpenAI", return_value=client):
                result = generate_table_insight(
                    _differentiator_table_payload(),
                    table_kind="differentiator_table",
                )

        self.assertIsInstance(result, InsightResult)
        self.assertFalse(result.was_template)
        self.assertEqual(result.insight, headline)

    def test_generate_outlier_insight(self) -> None:
        headline = "Outlier: Inside sales reaches 81% of Winners vs 12% of Laggards"
        client = _mock_openai_client(headline)

        with patch.dict(os.environ, {"PORTKEY_API_KEY": "test-key"}):
            with patch("src.ai_insights.OpenAI", return_value=client):
                result = generate_outlier_insight(
                    _differentiator_table_payload(),
                    table_kind="outlier",
                )

        self.assertIsInstance(result, InsightResult)
        self.assertFalse(result.was_template)
        self.assertEqual(result.insight, headline)

    def test_table_payload_includes_allowed_numbers(self) -> None:
        formatted = _format_differentiator_table_payload(_differentiator_table_payload())
        parsed = json.loads(formatted)

        self.assertIn("allowed_numbers", parsed)
        self.assertIn(67, parsed["allowed_numbers"])
        self.assertIn(6.75, parsed["allowed_numbers"])

    def test_categorize_questions_template_fallback_when_no_api_key(self) -> None:
        questions = [
            {
                "question_id": "Q1",
                "question_text": "What is your industry?",
                "is_demographic": True,
            },
            {
                "question_id": "Q2",
                "question_text": "Revenue grew?",
                "is_demographic": False,
            },
        ]

        with patch.dict(os.environ, {}, clear=True):
            result = categorize_questions_into_themes(questions)

        self.assertTrue(result["was_template"])
        theme_names = [theme["name"] for theme in result["themes"]]
        self.assertIn("Demographics", theme_names)
        self.assertIn("All Questions", theme_names)

    def test_categorize_questions_validates_all_ids_covered(self) -> None:
        questions = [
            {"question_id": "Q1", "question_text": "A", "is_demographic": False},
            {"question_id": "Q2", "question_text": "B", "is_demographic": False},
            {"question_id": "Q3", "question_text": "C", "is_demographic": False},
        ]
        mock_response = json.dumps(
            {
                "themes": [
                    {"name": "T1", "question_ids": ["Q1"]},
                    {"name": "T2", "question_ids": ["Q2"]},
                ]
            }
        )
        client = _mock_openai_client(mock_response)

        with patch.dict(os.environ, {"PORTKEY_API_KEY": "test-key"}):
            with patch("src.ai_insights.OpenAI", return_value=client):
                result = categorize_questions_into_themes(questions)

        all_qids = set()
        for theme in result["themes"]:
            all_qids.update(theme["question_ids"])
        self.assertIn("Q3", all_qids)
        self.assertFalse(result["was_template"])

    def test_generate_short_labels_fallback_truncates_when_no_api_key(self) -> None:
        questions = [
            {
                "question_id": "Q1",
                "question_text": "This is a very long question text " * 5,
            }
        ]

        with patch.dict(os.environ, {}, clear=True):
            labels = generate_short_labels(questions)

        self.assertIn("Q1", labels)
        self.assertLessEqual(len(labels["Q1"]), 50)

    def test_generate_short_labels_covers_all_questions(self) -> None:
        questions = [
            {"question_id": "Q1", "question_text": "Revenue growth question"},
            {"question_id": "Q2", "question_text": "Customer strategy question"},
        ]
        client = _mock_openai_client(json.dumps({"labels": {"Q1": "Revenue growth"}}))

        with patch.dict(os.environ, {"PORTKEY_API_KEY": "test-key"}):
            with patch("src.ai_insights.OpenAI", return_value=client):
                labels = generate_short_labels(questions)

        self.assertEqual(labels["Q1"], "Revenue growth")
        self.assertIn("Q2", labels)
        self.assertEqual(labels["Q2"], "Customer strategy question")

    def test_bulk_ai_enhancement_calls_use_25_second_timeout(self) -> None:
        questions = [
            {
                "question_id": "Q1",
                "question_text": "Revenue growth question",
                "question_type": "single_select",
                "is_demographic": False,
            }
        ]

        theme_client = _mock_openai_client(
            json.dumps({"themes": [{"name": "Growth", "question_ids": ["Q1"]}]})
        )
        with patch.dict(os.environ, {"PORTKEY_API_KEY": "test-key"}):
            with patch("src.ai_insights.OpenAI", return_value=theme_client):
                categorize_questions_into_themes(questions)
        self.assertEqual(
            theme_client.chat.completions.create.call_args.kwargs["timeout"],
            25,
        )

        labels_client = _mock_openai_client(json.dumps({"labels": {"Q1": "Growth"}}))
        with patch.dict(os.environ, {"PORTKEY_API_KEY": "test-key"}):
            with patch("src.ai_insights.OpenAI", return_value=labels_client):
                generate_short_labels(questions)
        self.assertEqual(
            labels_client.chat.completions.create.call_args.kwargs["timeout"],
            25,
        )

        demo_client = _mock_openai_client(
            json.dumps(
                {
                    "matches": [{"question_id": "Q1", "category": "industry", "tier": 2}],
                    "other_demographics": [],
                }
            )
        )
        with patch.dict(os.environ, {"PORTKEY_API_KEY": "test-key"}):
            with patch("src.ai_insights.OpenAI", return_value=demo_client):
                categorize_demographic_questions(questions)
        self.assertEqual(
            demo_client.chat.completions.create.call_args.kwargs["timeout"],
            25,
        )

    def test_categorize_demographic_questions_orders_by_priority_tier(self) -> None:
        questions = [
            {
                "question_id": "Q_INDUSTRY",
                "question_text": "What industry is your company in?",
                "question_type": "single_select",
            },
            {
                "question_id": "Q_COUNTRY",
                "question_text": "What country are you based in?",
                "question_type": "single_select",
            },
            {
                "question_id": "Q_OTHER",
                "question_text": "What procurement cohort are you in?",
                "question_type": "single_select",
            },
        ]
        mock_response = json.dumps(
            {
                "matches": [
                    {"question_id": "Q_INDUSTRY", "category": "industry", "tier": 2},
                    {"question_id": "Q_COUNTRY", "category": "country", "tier": 1},
                ],
                "other_demographics": ["Q_OTHER"],
            }
        )
        client = _mock_openai_client(mock_response)

        with patch.dict(os.environ, {"PORTKEY_API_KEY": "test-key"}):
            with patch("src.ai_insights.OpenAI", return_value=client):
                result = categorize_demographic_questions(questions)

        self.assertEqual(
            result["priority_ordered"],
            ["Q_COUNTRY", "Q_INDUSTRY", "Q_OTHER"],
        )
        self.assertEqual(result["categories"]["Q_COUNTRY"], "country")
        self.assertFalse(result["was_template"])


if __name__ == "__main__":
    unittest.main()
