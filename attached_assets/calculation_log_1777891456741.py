"""Append-only calculation audit log for analysis runs."""

from __future__ import annotations

from threading import Lock

from src.models import AuditRecord


class CalculationLog:
    """Thread-safe append-only audit log of every computed metric."""

    def __init__(self) -> None:
        self._records: list[AuditRecord] = []
        self._lock = Lock()

    def record(self, audit: AuditRecord) -> None:
        """Append a single audit record."""
        with self._lock:
            self._records.append(audit)

    def all_records(self) -> tuple[AuditRecord, ...]:
        """Return all records in insertion order."""
        with self._lock:
            return tuple(self._records)

    def __len__(self) -> int:
        with self._lock:
            return len(self._records)
