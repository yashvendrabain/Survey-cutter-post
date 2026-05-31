"""Tests for generic respondent ID column detection."""

from __future__ import annotations

import unittest

from src.raw_decoder import _find_respondent_id_column


class TestRespondentIdDetection(unittest.TestCase):
    def test_respondent_column_named_Respondent(self) -> None:
        self.assertEqual(
            _find_respondent_id_column(["Respondent", "Q1", "Q2"]),
            "Respondent",
        )

    def test_respondent_column_named_ResponseID(self) -> None:
        self.assertEqual(
            _find_respondent_id_column(["responseid", "Q1", "Q2"]),
            "responseid",
        )

    def test_respondent_column_named_record_legacy(self) -> None:
        self.assertEqual(
            _find_respondent_id_column(["record", "Q1", "Q2"]),
            "record",
        )

    def test_respondent_column_falls_through_to_first_column_when_none_match(self) -> None:
        self.assertEqual(
            _find_respondent_id_column(["first_col", "Q1", "Q2"]),
            "first_col",
        )

    def test_respondent_column_priority_Respondent_over_record(self) -> None:
        self.assertEqual(
            _find_respondent_id_column(["record", "Respondent", "Q1"]),
            "Respondent",
        )


if __name__ == "__main__":
    unittest.main()
