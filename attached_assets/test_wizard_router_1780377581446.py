"""Tests for adapter-router wizard fallback decisions."""

from __future__ import annotations

import unittest

from src.adapters.base import AdapterDetectionResult, DataMapAdapter
from src.adapters.registry import AdapterRouter, NoAdapterError


class _FakeAdapter(DataMapAdapter):
    def __init__(
        self,
        name: str,
        confidence: float,
        reason: str = "reason",
        question_count: int = 2,
    ) -> None:
        self.name = name
        self.confidence = confidence
        self.reason = reason
        self.question_count = question_count

    def detect(self, workbook, raw_df=None) -> AdapterDetectionResult:
        del workbook, raw_df
        return AdapterDetectionResult(self.confidence, self.reason)

    def parse(self, workbook, raw_df=None) -> dict:
        del workbook, raw_df
        return {
            "questions": [
                {"canonical_id": f"Q{index}"}
                for index in range(1, self.question_count + 1)
            ]
        }


class TestWizardRouter(unittest.TestCase):
    def test_router_signals_needs_wizard_below_threshold(self) -> None:
        needs_wizard, _scores = AdapterRouter([_FakeAdapter("weak", 0.25)]).needs_wizard(object())

        self.assertTrue(needs_wizard)

    def test_router_does_not_signal_wizard_above_threshold(self) -> None:
        needs_wizard, _scores = AdapterRouter([_FakeAdapter("strong", 0.6)]).needs_wizard(object())

        self.assertFalse(needs_wizard)

    def test_router_returns_all_adapter_scores_in_decision(self) -> None:
        needs_wizard, scores = AdapterRouter(
            [_FakeAdapter("a", 0.2, "low"), _FakeAdapter("b", 0.8, "high")]
        ).needs_wizard(object())

        self.assertFalse(needs_wizard)
        self.assertEqual(scores, [("b", 0.8, "high"), ("a", 0.2, "low")])

    def test_threshold_at_exact_03_does_not_trigger_wizard(self) -> None:
        needs_wizard, _scores = AdapterRouter([_FakeAdapter("exact", 0.3)]).needs_wizard(object())

        self.assertFalse(needs_wizard)

    def test_threshold_at_04_does_not_trigger_wizard(self) -> None:
        needs_wizard, _scores = AdapterRouter([_FakeAdapter("warning_zone", 0.4)]).needs_wizard(object())

        self.assertFalse(needs_wizard)

    def test_router_falls_back_to_wizard_when_parse_returns_zero_questions(self) -> None:
        router = AdapterRouter([_FakeAdapter("false_positive", 0.9, question_count=0)])

        with self.assertRaisesRegex(NoAdapterError, "parsed only 0 questions"):
            router.parse(object())

    def test_router_falls_back_when_parse_returns_one_question(self) -> None:
        router = AdapterRouter([_FakeAdapter("false_positive", 0.9, question_count=1)])

        with self.assertRaisesRegex(NoAdapterError, "parsed only 1 questions"):
            router.parse(object())

    def test_router_succeeds_when_parse_returns_two_or_more_questions(self) -> None:
        router = AdapterRouter([_FakeAdapter("valid", 0.9, question_count=2)])

        parsed = router.parse(object())

        self.assertEqual(len(parsed["questions"]), 2)


if __name__ == "__main__":
    unittest.main()
