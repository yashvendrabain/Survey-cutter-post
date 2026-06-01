"""Tests for adapter-router wizard fallback decisions."""

from __future__ import annotations

import unittest

from src.adapters.base import AdapterDetectionResult, DataMapAdapter
from src.adapters.registry import AdapterRouter


class _FakeAdapter(DataMapAdapter):
    def __init__(self, name: str, confidence: float, reason: str = "reason") -> None:
        self.name = name
        self.confidence = confidence
        self.reason = reason

    def detect(self, workbook, raw_df=None) -> AdapterDetectionResult:
        del workbook, raw_df
        return AdapterDetectionResult(self.confidence, self.reason)

    def parse(self, workbook, raw_df=None) -> dict:
        del workbook, raw_df
        return {"questions": []}


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


if __name__ == "__main__":
    unittest.main()
