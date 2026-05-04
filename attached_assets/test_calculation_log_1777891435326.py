"""Tests for the calculation audit log."""

from __future__ import annotations

import threading
import unittest
from datetime import datetime, timezone

from src.calculation_log import CalculationLog
from src.models import AuditRecord


def make_audit(metric_name: str = "metric") -> AuditRecord:
    return AuditRecord(
        output_sheet="Sheet",
        metric_name=metric_name,
        source_question_id="Q1",
        source_columns=("Q1",),
        filter_expr=None,
        numerator=None,
        denominator=1,
        formula="test formula",
        value_raw=1.0,
        valid_n=1,
        missing_n=0,
        timestamp=datetime.now(timezone.utc),
    )


class TestCalculationLog(unittest.TestCase):
    def test_record_appends(self) -> None:
        log = CalculationLog()
        audit = make_audit()

        log.record(audit)

        self.assertEqual(len(log), 1)
        self.assertEqual(log.all_records(), (audit,))

    def test_all_records_returns_tuple_in_order(self) -> None:
        log = CalculationLog()
        first = make_audit("first")
        second = make_audit("second")

        log.record(first)
        log.record(second)

        self.assertEqual(log.all_records(), (first, second))

    def test_thread_safety_smoke_test(self) -> None:
        log = CalculationLog()

        def append_records(thread_id: int) -> None:
            for index in range(100):
                log.record(make_audit(f"thread-{thread_id}-{index}"))

        threads = [
            threading.Thread(target=append_records, args=(thread_id,))
            for thread_id in range(10)
        ]
        for thread in threads:
            thread.start()
        for thread in threads:
            thread.join()

        self.assertEqual(len(log), 1000)


if __name__ == "__main__":
    unittest.main()
