"""Tests for data-map adapter primitives."""

from __future__ import annotations

from dataclasses import FrozenInstanceError
import unittest

from src.adapters.base import AdapterDetectionResult, DataMapAdapter
from src.adapters.registry import (
    CONFIDENCE_USE_THRESHOLD,
    AdapterRouter,
    NoAdapterError,
)


class _FakeAdapter(DataMapAdapter):
    def __init__(self, name: str, confidence: float, payload: dict | None = None):
        self.name = name
        self.confidence = confidence
        self.payload = payload or {"questions": []}

    def detect(self, workbook, raw_df=None) -> AdapterDetectionResult:
        del workbook, raw_df
        return AdapterDetectionResult(self.confidence, f"{self.name} reason")

    def parse(self, workbook, raw_df=None) -> dict:
        del workbook, raw_df
        return self.payload


class TestAdapterBase(unittest.TestCase):
    def test_detection_result_is_frozen(self) -> None:
        result = AdapterDetectionResult(0.8, "matched")
        with self.assertRaises(FrozenInstanceError):
            result.confidence = 0.1  # type: ignore[misc]

    def test_router_picks_highest_confidence_adapter(self) -> None:
        router = AdapterRouter(
            [
                _FakeAdapter("low", 0.4),
                _FakeAdapter(
                    "high",
                    0.9,
                    {"winner": True},
                ),
            ]
        )

        adapter, result = router.pick_adapter(object())

        self.assertEqual(adapter.name, "high")
        self.assertEqual(result.confidence, 0.9)
        self.assertEqual(
            router.parse(object(), min_questions=0),
            {"winner": True},
        )

    def test_router_raises_when_no_adapter_registered(self) -> None:
        router = AdapterRouter([])

        with self.assertRaisesRegex(NoAdapterError, "No adapters registered"):
            router.pick_adapter(object())

    def test_router_raises_when_all_adapters_below_reject_threshold(self) -> None:
        router = AdapterRouter([_FakeAdapter("weak", 0.1)])

        with self.assertRaisesRegex(NoAdapterError, "No adapter matched"):
            router.pick_adapter(object())

    def test_router_flags_ambiguous_high_confidence_tie(self) -> None:
        router = AdapterRouter(
            [
                _FakeAdapter("first", 0.71),
                _FakeAdapter("second", 0.69),
            ]
        )

        with self.assertRaisesRegex(NoAdapterError, "Ambiguous format"):
            router.pick_adapter(object())

    def test_warning_zone_adapter_can_still_be_selected(self) -> None:
        router = AdapterRouter([_FakeAdapter("warning", 0.4)])

        adapter, result = router.pick_adapter(object())

        self.assertEqual(adapter.name, "warning")
        self.assertLess(result.confidence, CONFIDENCE_USE_THRESHOLD)


if __name__ == "__main__":
    unittest.main()
