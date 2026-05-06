"""Tests for Word survey parsing."""

from __future__ import annotations

import unittest

from src.word_survey_parser import parse_word_survey
from tests.conftest import FORMAT_A_DOCX_PATH, FORMAT_B_DOCX_PATH


REQUIRED_QUESTION_KEYS = {
    "canonical_id",
    "raw_id",
    "question_text",
    "type_hint",
    "value_range",
    "options",
    "sub_columns",
    "parent_canonical_id",
    "source_row",
    "warnings",
}
REQUIRED_DATAMAP_KEYS = {
    "questions",
    "source_path",
    "sheet_name",
    "total_rows_in_sheet",
    "parser_warnings",
}


def question_by_id(data_map: dict, canonical_id: str) -> dict:
    return next(
        question
        for question in data_map["questions"]
        if question["canonical_id"] == canonical_id
    )


class TestWordSurveyParser(unittest.TestCase):
    def test_detects_format_a_correctly(self) -> None:
        data_map = parse_word_survey(str(FORMAT_A_DOCX_PATH))

        canonical_ids = [question["canonical_id"] for question in data_map["questions"]]
        self.assertIn("Q1", canonical_ids)
        self.assertIn("H_Q4", canonical_ids)
        self.assertIn("S_OpenEnd_JobTitle", canonical_ids)

    def test_detects_format_b_correctly(self) -> None:
        data_map = parse_word_survey(str(FORMAT_B_DOCX_PATH))

        self.assertEqual(
            [question["canonical_id"] for question in data_map["questions"]],
            ["Q01", "Q02", "Q03"],
        )

    def test_format_a_parses_question_ids(self) -> None:
        data_map = parse_word_survey(str(FORMAT_A_DOCX_PATH))

        self.assertEqual(question_by_id(data_map, "Q1")["raw_id"], "Q1")
        self.assertEqual(question_by_id(data_map, "H_Q4")["raw_id"], "H_Q4")

    def test_format_a_parses_single_select_type_hint(self) -> None:
        question = question_by_id(parse_word_survey(str(FORMAT_A_DOCX_PATH)), "Q1")

        self.assertEqual(question["type_hint"], "values_range")
        self.assertEqual(question["value_range"], (1, 3))

    def test_format_a_parses_multi_select_type_hint(self) -> None:
        question = question_by_id(parse_word_survey(str(FORMAT_A_DOCX_PATH)), "Q2")

        self.assertEqual(question["type_hint"], "values_range")
        self.assertEqual(question["value_range"], (0, 1))

    def test_format_a_parses_open_text_type_hint(self) -> None:
        question = question_by_id(parse_word_survey(str(FORMAT_A_DOCX_PATH)), "Q3")

        self.assertEqual(question["type_hint"], "open_text")
        self.assertIsNone(question["value_range"])

    def test_format_a_strips_programmer_instructions_from_labels(self) -> None:
        question = question_by_id(parse_word_survey(str(FORMAT_A_DOCX_PATH)), "Q1")

        labels = [label for _code, label in question["options"]]
        self.assertEqual(
            labels,
            ["Employed full time", "Employed part-time", "Retired"],
        )

    def test_format_a_preserves_annotations_in_warnings(self) -> None:
        question = question_by_id(parse_word_survey(str(FORMAT_A_DOCX_PATH)), "Q2")

        self.assertIn("ANNOTATION: [MULTI-SELECT; RANDOM ORDER]", question["warnings"])
        self.assertIn("ANNOTATION: [anchor]", question["warnings"])
        self.assertIn("ANNOTATION: [exclusive]", question["warnings"])

    def test_format_a_detects_parent_canonical_id_for_oe(self) -> None:
        question = question_by_id(parse_word_survey(str(FORMAT_A_DOCX_PATH)), "Q2oe")

        self.assertEqual(question["type_hint"], "open_text")
        self.assertEqual(question["parent_canonical_id"], "Q2")

    def test_format_a_infers_sequential_option_codes(self) -> None:
        question = question_by_id(parse_word_survey(str(FORMAT_A_DOCX_PATH)), "Q1")

        self.assertEqual([code for code, _label in question["options"]], [1, 2, 3])

    def test_format_a_respects_explicit_option_codes_901_997(self) -> None:
        question = question_by_id(parse_word_survey(str(FORMAT_A_DOCX_PATH)), "Q2")

        self.assertIn((901, "Other open-end anchor"), question["options"])
        self.assertIn((997, "None of the above"), question["options"])

    def test_format_b_auto_generates_question_ids(self) -> None:
        data_map = parse_word_survey(str(FORMAT_B_DOCX_PATH))

        self.assertEqual(data_map["questions"][0]["raw_id"], "Q01")
        self.assertEqual(data_map["questions"][1]["raw_id"], "Q02")

    def test_format_b_parses_type_hint_from_type_line(self) -> None:
        data_map = parse_word_survey(str(FORMAT_B_DOCX_PATH))

        self.assertEqual(question_by_id(data_map, "Q01")["value_range"], (1, 4))
        self.assertEqual(question_by_id(data_map, "Q02")["value_range"], (0, 1))

    def test_format_b_strips_tag_annotations_from_labels(self) -> None:
        question = question_by_id(parse_word_survey(str(FORMAT_B_DOCX_PATH)), "Q02")

        self.assertIn((1, "Finance"), question["options"])
        self.assertIn("ANNOTATION: [TAG: FS]", question["warnings"])

    def test_format_b_open_text_question(self) -> None:
        question = question_by_id(parse_word_survey(str(FORMAT_B_DOCX_PATH)), "Q03")

        self.assertEqual(question["type_hint"], "open_text")
        self.assertEqual(question["options"], [])

    def test_output_is_valid_datamap_typedict(self) -> None:
        for path in (FORMAT_A_DOCX_PATH, FORMAT_B_DOCX_PATH):
            data_map = parse_word_survey(str(path))

            self.assertEqual(set(data_map.keys()), REQUIRED_DATAMAP_KEYS)
            self.assertEqual(data_map["sheet_name"], "word_document")
            self.assertIsInstance(data_map["questions"], list)
            self.assertIsInstance(data_map["parser_warnings"], list)
            for question in data_map["questions"]:
                self.assertEqual(set(question.keys()), REQUIRED_QUESTION_KEYS)
                self.assertIsInstance(question["options"], list)
                self.assertIsInstance(question["sub_columns"], list)
                self.assertIsInstance(question["warnings"], list)


if __name__ == "__main__":
    unittest.main()
