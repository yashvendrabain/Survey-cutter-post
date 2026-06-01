"""Tests for compact two-column data map parsing."""

from __future__ import annotations

import unittest

from openpyxl import Workbook

from src.datamap_parser import (
    _detect_datamap_format,
    _is_question_id,
    _parse_compact_datamap,
)


def _worksheet_from_rows(rows: list[list[object]]) -> object:
    workbook = Workbook()
    worksheet = workbook.active
    for row in rows:
        worksheet.append(row)
    return worksheet


class TestCompactDatamapParser(unittest.TestCase):
    def test_detect_compact_format_two_column_only(self) -> None:
        worksheet = _worksheet_from_rows(
            [
                ["Q1", "Employment status"],
                ["1", "Full-time"],
                ["2", "Part-time"],
                [None, None],
                ["Q2", "Industry"],
                ["1", "Technology"],
                ["2", "Retail"],
            ]
        )

        self.assertEqual(_detect_datamap_format(worksheet), "compact_two_column")

    def test_detect_bcn_format_with_explicit_headers(self) -> None:
        worksheet = _worksheet_from_rows(
            [["Question ID", "Question text", "Type"], ["Q1", "Text", "Values"]]
        )

        self.assertEqual(_detect_datamap_format(worksheet), "bcn_multicolumn")

    def test_detect_unknown_format_returns_unknown(self) -> None:
        worksheet = _worksheet_from_rows([["hello", "world"], ["not", "a map"]])

        self.assertEqual(_detect_datamap_format(worksheet), "unknown")

    def test_compact_parser_extracts_question_and_options(self) -> None:
        worksheet = _worksheet_from_rows(
            [["Q1", "Employment status"], ["1", "Full-time"], ["2", "Part-time"], ["3", "Contract"], ["4", "Other"]]
        )

        questions = _parse_compact_datamap(worksheet)

        self.assertEqual(len(questions), 1)
        self.assertEqual(questions[0]["canonical_id"], "Q1")
        self.assertEqual(questions[0]["question_text"], "Employment status")
        self.assertEqual(
            questions[0]["options"],
            [(1, "Full-time"), (2, "Part-time"), (3, "Contract"), (4, "Other")],
        )

    def test_compact_parser_blank_row_terminates_question(self) -> None:
        worksheet = _worksheet_from_rows(
            [
                ["Q1", "Employment status"],
                ["1", "Full-time"],
                [None, None],
                ["Q2", "Industry"],
                ["1", "Technology"],
            ]
        )

        questions = _parse_compact_datamap(worksheet)

        self.assertEqual([q["canonical_id"] for q in questions], ["Q1", "Q2"])
        self.assertEqual(questions[0]["options"], [(1, "Full-time")])
        self.assertEqual(questions[1]["options"], [(1, "Technology")])

    def test_compact_parser_handles_alphabetic_option_codes(self) -> None:
        worksheet = _worksheet_from_rows(
            [["Q1", "Employment status"], ["a", "Full-time"], ["b", "Part-time"], ["c", "Contract"]]
        )

        questions = _parse_compact_datamap(worksheet)

        self.assertEqual(
            questions[0]["options"],
            [("a", "Full-time"), ("b", "Part-time"), ("c", "Contract")],
        )

    def test_compact_parser_multiple_questions_all_extracted(self) -> None:
        rows: list[list[object]] = []
        for index in range(1, 6):
            rows.extend(
                [
                    [f"Q{index}", f"Question {index}"],
                    ["1", "Yes"],
                    ["2", "No"],
                    [None, None],
                ]
            )
        worksheet = _worksheet_from_rows(rows)

        questions = _parse_compact_datamap(worksheet)

        self.assertEqual([q["canonical_id"] for q in questions], ["Q1", "Q2", "Q3", "Q4", "Q5"])
        self.assertTrue(all(q["options"] == [(1, "Yes"), (2, "No")] for q in questions))

    def test_question_id_detector_accepts_Q1(self) -> None:
        self.assertTrue(_is_question_id("Q1"))

    def test_question_id_detector_accepts_Q33r1(self) -> None:
        self.assertTrue(_is_question_id("Q33r1"))

    def test_question_id_detector_rejects_plain_integers(self) -> None:
        self.assertFalse(_is_question_id("1"))

    def test_question_id_detector_rejects_text(self) -> None:
        self.assertFalse(_is_question_id("Full-time"))


if __name__ == "__main__":
    unittest.main()
