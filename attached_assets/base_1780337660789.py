"""Base interfaces for data-map parser adapters."""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True, slots=True)
class AdapterDetectionResult:
    """Detection score and explanation returned by adapter.detect()."""

    confidence: float
    reason: str
    is_certain: bool = False

    def __post_init__(self) -> None:
        object.__setattr__(self, "is_certain", self.confidence >= 0.5)


class DataMapAdapter(ABC):
    """Base class for survey input parsers."""

    name: str = "abstract_base"

    @abstractmethod
    def detect(self, workbook: Any, raw_df: Any | None = None) -> AdapterDetectionResult:
        """Inspect the workbook and return a confidence score."""
        raise NotImplementedError

    @abstractmethod
    def parse(self, workbook: Any, raw_df: Any | None = None) -> dict:
        """Parse the workbook into the normalized DataMap shape."""
        raise NotImplementedError
