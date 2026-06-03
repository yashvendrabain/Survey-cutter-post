"""Tests for raw sub-column headers that embed option labels."""

from __future__ import annotations

import unittest

from src.models import QuestionType
from src.adapters.label_pattern_subcolumn import infer_label_pattern_sub_columns
from src.question_classifier import (
    classify_questions,
)


def _question(
    canonical_id: str = "Q3",
    options: list[tuple[int | str, str]] | None = None,
    sub_columns: list[tuple[str, str]] | None = None,
    value_range: tuple[int, int] | None = (1, 6),
) -> dict:
    return {
        "canonical_id": canonical_id,
        "raw_id": canonical_id,
        "question_text": "Select all that apply",
        "type_hint": "values_range",
        "value_range": value_range,
        "options": options
        if options is not None
        else [
            (1, "We sell directly to other businesses (B2B)"),
            (2, "We sell directly to consumers (B2C)"),
            (3, "We sell to intermediaries who sell to consumers (B2B2C)"),
            (4, "None of the above"),
        ],
        "sub_columns": sub_columns if sub_columns is not None else [],
        "parent_canonical_id": None,
        "source_row": 1,
        "warnings": [],
    }


def _schema_for(question: dict, raw_columns: list[str]):
    return classify_questions(
        {
            "questions": [question],
            "source_path": "test-map.xlsx",
            "sheet_name": "Sheet1",
            "total_rows_in_sheet": 1,
            "parser_warnings": [],
        },
        raw_columns,
        respondent_id_column="Respondent",
        total_respondents=10,
    )


class TestSubcolumnLabelPattern(unittest.TestCase):
    def test_detect_colon_separated_subcolumns(self) -> None:
        raw_columns = [
            "Respondent",
            "Q3: We sell directly to other businesses (B2B)",
            "Q3: We sell directly to consumers (B2C)",
        ]

        spec = _schema_for(_question(), raw_columns).get_question("Q3")

        self.assertEqual(spec.question_type, QuestionType.MULTI_SELECT_BINARY)
        self.assertEqual(spec.raw_columns, tuple(raw_columns[1:]))

    def test_detect_hyphen_separated_subcolumns(self) -> None:
        raw_columns = [
            "Respondent",
            "Q3 - We sell directly to other businesses (B2B)",
            "Q3 - We sell directly to consumers (B2C)",
        ]

        spec = _schema_for(_question(), raw_columns).get_question("Q3")

        self.assertEqual(spec.question_type, QuestionType.MULTI_SELECT_BINARY)
        self.assertEqual(spec.raw_columns, tuple(raw_columns[1:]))

    def test_exact_label_match_preferred_over_substring(self) -> None:
        raw_column = "Q3: Direct sales"
        spec = _schema_for(
            _question(options=[(1, "Direct"), (2, "Direct sales")]),
            ["Respondent", raw_column, "Q3: Direct"],
        ).get_question("Q3")

        self.assertEqual(spec.option_map[raw_column], "Direct sales")

    def test_case_insensitive_label_match(self) -> None:
        raw_column = "Q3: We Sell Directly"
        spec = _schema_for(
            _question(options=[(1, "we sell directly"), (2, "Other")]),
            ["Respondent", raw_column, "Q3: Other"],
        ).get_question("Q3")

        self.assertEqual(spec.option_map[raw_column], "we sell directly")

    def test_punctuation_normalized_match(self) -> None:
        raw_column = "Q3: We sell directly to other businesses"
        spec = _schema_for(
            _question(
                options=[
                    (1, "We sell directly to other businesses (B2B)"),
                    (2, "None of the above"),
                ]
            ),
            ["Respondent", raw_column, "Q3: None of the above"],
        ).get_question("Q3")

        self.assertEqual(
            spec.option_map[raw_column],
            "We sell directly to other businesses (B2B)",
        )

    def test_substring_fallback_match(self) -> None:
        inferred, warnings = infer_label_pattern_sub_columns(
            _question(),
            ("Q3: other businesses",),
        )

        self.assertEqual(
            inferred,
            [("Q3: other businesses", "We sell directly to other businesses (B2B)")],
        )
        self.assertIn("Accepted substring match", warnings[0])

    def test_unmatched_column_surfaces_warning(self) -> None:
        inferred, warnings = infer_label_pattern_sub_columns(
            _question(),
            ("Q3: Something not in the data map",),
        )

        self.assertEqual(inferred, [])
        self.assertIn("Could not match column", warnings[0])

    def test_unmatched_column_uses_schema_warning_channel(self) -> None:
        raw_columns = [
            "Respondent",
            "Q3: We sell directly to other businesses (B2B)",
            "Q3: None of the above",
            "Q3: Something not in the data map",
        ]

        spec = _schema_for(_question(), raw_columns).get_question("Q3")

        self.assertTrue(spec.classification_confidence_low)
        self.assertTrue(
            any("Could not match column" in warning for warning in spec.warnings)
        )

    def test_question_type_promoted_to_multi_select_binary(self) -> None:
        raw_columns = [
            "Respondent",
            "Q3: We sell directly to other businesses (B2B)",
            "Q3: None of the above",
        ]

        spec = _schema_for(_question(), raw_columns).get_question("Q3")

        self.assertEqual(spec.question_type, QuestionType.MULTI_SELECT_BINARY)
        self.assertEqual(spec.value_range, (0, 1))

    def test_bcn_pattern_takes_priority(self) -> None:
        question = _question(
            options=[],
            sub_columns=[("Q3r1", "B2B"), ("Q3r2", "B2C")],
            value_range=(0, 1),
        )
        raw_columns = ["Respondent", "Q3r1", "Q3r2", "Q3: B2B"]

        spec = _schema_for(question, raw_columns).get_question("Q3")

        self.assertEqual(spec.raw_columns, ("Q3r1", "Q3r2"))
        self.assertNotIn("Q3: B2B", spec.raw_columns)

    def test_label_pattern_does_not_affect_questions_without_options(self) -> None:
        question = _question("Q70", options=[], value_range=(0, 100))
        raw_columns = ["Respondent", "Q70", "Q70: Something"]

        spec = _schema_for(question, raw_columns).get_question("Q70")

        self.assertEqual(spec.question_type, QuestionType.DIRECT_NUMERIC)
        self.assertEqual(spec.raw_columns, ("Q70",))


if __name__ == "__main__":
    unittest.main()
