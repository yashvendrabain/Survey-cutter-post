"""Adapter router for survey data-map parsing."""

from __future__ import annotations

from typing import Sequence

from src.adapters.base import AdapterDetectionResult, DataMapAdapter


CONFIDENCE_USE_THRESHOLD = 0.5
CONFIDENCE_REJECT_THRESHOLD = 0.3
AMBIGUOUS_CONFIDENCE_THRESHOLD = 0.5
AMBIGUOUS_TIE_DELTA = 0.05


class NoAdapterError(Exception):
    """Raised when no data-map adapter can confidently parse a workbook."""


class AdapterRouter:
    def __init__(self, adapters: Sequence[DataMapAdapter]):
        self.adapters = list(adapters)

    def detect_all(
        self,
        workbook,
        raw_df=None,
    ) -> list[tuple[DataMapAdapter, AdapterDetectionResult]]:
        """Run detect() on every adapter, sorted by confidence descending."""

        results = [(adapter, adapter.detect(workbook, raw_df)) for adapter in self.adapters]
        results.sort(key=lambda item: item[1].confidence, reverse=True)
        return results

    def pick_adapter(self, workbook, raw_df=None) -> tuple[DataMapAdapter, AdapterDetectionResult]:
        """Pick the best adapter or raise NoAdapterError."""

        results = self.detect_all(workbook, raw_df)
        if not results:
            raise NoAdapterError("No adapters registered")

        best_adapter, best_result = results[0]
        if best_result.confidence < CONFIDENCE_REJECT_THRESHOLD:
            detail = ", ".join(
                f"{adapter.name}={result.confidence:.2f}"
                for adapter, result in results
            )
            raise NoAdapterError(f"No adapter matched. Detected: {detail}")

        if len(results) > 1:
            _second_adapter, second_result = results[1]
            if (
                best_result.confidence >= AMBIGUOUS_CONFIDENCE_THRESHOLD
                and second_result.confidence >= AMBIGUOUS_CONFIDENCE_THRESHOLD
                and abs(best_result.confidence - second_result.confidence)
                <= AMBIGUOUS_TIE_DELTA
            ):
                raise NoAdapterError(
                    "Ambiguous format detected. Manual override required."
                )

        return best_adapter, best_result

    def needs_wizard(self, workbook, raw_df=None) -> tuple[bool, list[tuple[str, float, str]]]:
        """Return whether parser configuration should fall back to the wizard."""

        results = self.detect_all(workbook, raw_df)
        all_scores = [
            (adapter.name, result.confidence, result.reason)
            for adapter, result in results
        ]
        if not results:
            return True, all_scores
        best_score = results[0][1].confidence
        return best_score < CONFIDENCE_REJECT_THRESHOLD, all_scores

    def parse(self, workbook, raw_df=None, *, min_questions: int = 2) -> dict:
        """Parse workbook via the best-matching adapter.

        Args:
            min_questions: Minimum questions required from the parse to accept
                the adapter. The production default rejects false-positive
                matches that produce near-empty results.
        """

        adapter, result = self.pick_adapter(workbook, raw_df)
        parsed = adapter.parse(workbook, raw_df)
        questions = parsed.get("questions", [])
        if len(questions) < min_questions:
            raise NoAdapterError(
                f"Adapter {adapter.name!r} (confidence {result.confidence:.2f}) "
                f"parsed only {len(questions)} questions"
            )
        return parsed


_DEFAULT_REGISTRY: AdapterRouter | None = None


def get_default_registry() -> AdapterRouter:
    global _DEFAULT_REGISTRY
    if _DEFAULT_REGISTRY is None:
        from src.adapters.bcn_multicolumn import BcnMulticolumnAdapter
        from src.adapters.compact_two_column import CompactTwoColumnAdapter

        _DEFAULT_REGISTRY = AdapterRouter(
            [
                BcnMulticolumnAdapter(),
                CompactTwoColumnAdapter(),
            ]
        )
    return _DEFAULT_REGISTRY
