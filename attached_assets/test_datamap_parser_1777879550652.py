"""Tests for the data map parser state machine."""

from __future__ import annotations

import unittest

from src.datamap_parser import parse_datamap
from tests.conftest import DATAMAP_FIXTURE_PATH, MISSING_SHEET_FIXTURE_PATH


EXPECTED_QUESTIONS = [
    {
        "canonical_id": "Q3",
        "raw_id": "[Q3]",
        "question_text": "Are you currently in a full-time position",
        "type_hint": "values_range",
        "value_range": (1, 2),
        "options": [(1, "Yes"), (2, "No")],
        "sub_columns": [],
        "parent_canonical_id": None,
        "source_row": 2,
        "warnings": [],
    },
    {
        "canonical_id": "Q53",
        "raw_id": "Q53",
        "question_text": "Which of the following challenges...",
        "type_hint": "values_range",
        "value_range": (0, 1),
        "options": [(0, "Unchecked"), (1, "Checked")],
        "sub_columns": [
            ("Q53r1", "Knowledge gap"),
            ("Q53r2", "Sales marketing alignment gap"),
        ],
        "parent_canonical_id": None,
        "source_row": 10,
        "warnings": [],
    },
    {
        "canonical_id": "Q15",
        "raw_id": "Q15",
        "question_text": "What best describes your involvement...",
        "type_hint": "values_range",
        "value_range": (1, 4),
        "options": [
            (1, "Directly involved in decision making AND budget"),
            (2, "Directly involved in decision making OR budget"),
            (3, "Indirectly involved"),
            (4, "Not involved"),
        ],
        "sub_columns": [
            ("Q15r1", "Overall company strategy"),
            ("Q15r2", "Go-to-market strategy"),
            ("Q15r3", "Partner strategy"),
        ],
        "parent_canonical_id": None,
        "source_row": 17,
        "warnings": [],
    },
    {
        "canonical_id": "Q70",
        "raw_id": "[Q70]",
        "question_text": "What % of your pipeline...",
        "type_hint": "values_range",
        "value_range": (0, 100),
        "options": [],
        "sub_columns": [],
        "parent_canonical_id": None,
        "source_row": 27,
        "warnings": [],
    },
    {
        "canonical_id": "Q33",
        "raw_id": "Q33",
        "question_text": "Allocate 100 points...",
        "type_hint": "values_range",
        "value_range": (0, 999),
        "options": [],
        "sub_columns": [
            ("Q33r1", "Managing pricing"),
            ("Q33r2", "Customer-led shift"),
        ],
        "parent_canonical_id": None,
        "source_row": 30,
        "warnings": [],
    },
    {
        "canonical_id": "Q4r98oe",
        "raw_id": "[Q4r98oe]",
        "question_text": (
            "Which of the following industries... - Other (please specify)"
        ),
        "type_hint": "open_text",
        "value_range": None,
        "options": [],
        "sub_columns": [],
        "parent_canonical_id": "Q4",
        "source_row": 35,
        "warnings": [],
    },
    {
        "canonical_id": "QMissingType",
        "raw_id": "QMissingType",
        "question_text": "Header but missing type hint",
        "type_hint": None,
        "value_range": None,
        "options": [],
        "sub_columns": [],
        "parent_canonical_id": None,
        "source_row": 38,
        "warnings": ["header followed by blank row, no type hint"],
    },
    {
        "canonical_id": "vQTIME_MINUTES",
        "raw_id": "[vQTIME_MINUTES]",
        "question_text": "Survey length in minutes",
        "type_hint": "values_range",
        "value_range": (-99999999999999, 999999999999999),
        "options": [],
        "sub_columns": [],
        "parent_canonical_id": None,
        "source_row": 41,
        "warnings": [],
    },
]


def question_by_id(parsed: dict, canonical_id: str) -> dict:
    for question in parsed["questions"]:
        if question["canonical_id"] == canonical_id:
            return question
    raise AssertionError(f"question not found: {canonical_id}")


class TestDataMapParser(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.parsed = parse_datamap(str(DATAMAP_FIXTURE_PATH))

    def test_parse_simple_single_select(self) -> None:
        self.assertEqual(question_by_id(self.parsed, "Q3"), EXPECTED_QUESTIONS[0])

    def test_parse_multi_select_binary(self) -> None:
        self.assertEqual(question_by_id(self.parsed, "Q53"), EXPECTED_QUESTIONS[1])

    def test_parse_grid_pattern(self) -> None:
        self.assertEqual(question_by_id(self.parsed, "Q15"), EXPECTED_QUESTIONS[2])

    def test_parse_direct_numeric_no_options(self) -> None:
        self.assertEqual(question_by_id(self.parsed, "Q70"), EXPECTED_QUESTIONS[3])

    def test_parse_numeric_allocation(self) -> None:
        self.assertEqual(question_by_id(self.parsed, "Q33"), EXPECTED_QUESTIONS[4])

    def test_parse_open_text_follow_up(self) -> None:
        self.assertEqual(
            question_by_id(self.parsed, "Q4r98oe"), EXPECTED_QUESTIONS[5]
        )

    def test_parse_open_numeric_wide_range(self) -> None:
        self.assertEqual(
            question_by_id(self.parsed, "vQTIME_MINUTES"), EXPECTED_QUESTIONS[7]
        )

    def test_parse_skips_index_sheet(self) -> None:
        parsed_ids = [question["canonical_id"] for question in self.parsed["questions"]]

        self.assertEqual(self.parsed["sheet_name"], "Sheet1")
        self.assertNotIn("QIndex", parsed_ids)

    def test_parse_raises_on_missing_sheet1(self) -> None:
        with self.assertRaisesRegex(
            ValueError, "'Sheet1' sheet not found. Available sheets: IndexOnly"
        ):
            parse_datamap(str(MISSING_SHEET_FIXTURE_PATH))

    def test_parse_handles_consecutive_blank_rows(self) -> None:
        expected = {
            "questions": EXPECTED_QUESTIONS,
            "source_path": str(DATAMAP_FIXTURE_PATH),
            "sheet_name": "Sheet1",
            "total_rows_in_sheet": 42,
            "parser_warnings": ["orphan row at row 8: 'This is workbook junk'"],
        }

        self.assertEqual(self.parsed, expected)

    def test_parse_warns_on_orphan_row_between_blocks(self) -> None:
        self.assertEqual(
            self.parsed["parser_warnings"],
            ["orphan row at row 8: 'This is workbook junk'"],
        )

    def test_parse_warns_on_header_with_no_type_hint(self) -> None:
        self.assertEqual(
            question_by_id(self.parsed, "QMissingType"),
            EXPECTED_QUESTIONS[6],
        )

    def test_parse_strips_brackets_from_subcolumn_ids(self) -> None:
        q53 = question_by_id(self.parsed, "Q53")
        q15 = question_by_id(self.parsed, "Q15")
        q33 = question_by_id(self.parsed, "Q33")

        self.assertEqual([sub_col[0] for sub_col in q53["sub_columns"]], ["Q53r1", "Q53r2"])
        self.assertEqual(
            [sub_col[0] for sub_col in q15["sub_columns"]],
            ["Q15r1", "Q15r2", "Q15r3"],
        )
        self.assertEqual([sub_col[0] for sub_col in q33["sub_columns"]], ["Q33r1", "Q33r2"])

    def test_parse_preserves_brackets_in_raw_id(self) -> None:
        self.assertEqual(question_by_id(self.parsed, "Q3")["raw_id"], "[Q3]")
        self.assertEqual(question_by_id(self.parsed, "Q53")["raw_id"], "Q53")
        self.assertEqual(
            question_by_id(self.parsed, "vQTIME_MINUTES")["raw_id"],
            "[vQTIME_MINUTES]",
        )


if __name__ == "__main__":
    unittest.main()
