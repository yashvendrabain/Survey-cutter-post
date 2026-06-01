"""Tests for grid rows embedded in raw column labels."""

from __future__ import annotations

import unittest

from src.adapters.grid_categorical_row import (
    apply_grid_categorical_row_matching,
    grid_categorical_row_warnings,
)
from src.models import QuestionType
from src.question_classifier import classify_questions


def _question(
    canonical_id: str = "Q27",
    options: list[tuple[int | str, str]] | None = None,
    sub_columns: list[tuple[str, str]] | None = None,
) -> dict:
    return {
        "canonical_id": canonical_id,
        "raw_id": canonical_id,
        "question_text": "How have these sales channels changed?",
        "type_hint": "values_range",
        "value_range": (1, 5),
        "options": options
        if options is not None
        else [
            (1, "Increased"),
            (2, "Same"),
            (3, "Decreased"),
            (4, "N/A"),
            (5, "Don't know"),
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


class TestGridCategoricalRowAdapter(unittest.TestCase):
    def test_fires_when_rating_scale_options_and_label_pattern_columns(self) -> None:
        promoted = apply_grid_categorical_row_matching(
            _question(),
            ("Q27: Field sales", "Q27: Inside sales"),
        )

        self.assertEqual(
            promoted["sub_columns"],
            [("Q27: Field sales", "Field sales"), ("Q27: Inside sales", "Inside sales")],
        )

    def test_does_not_fire_when_options_are_not_rating_scale(self) -> None:
        question = _question(
            options=[
                (1, "B2B"),
                (2, "B2C"),
                (3, "Channel partners"),
            ]
        )

        promoted = apply_grid_categorical_row_matching(
            question,
            ("Q27: Field sales", "Q27: Inside sales"),
        )

        self.assertEqual(promoted["sub_columns"], [])

    def test_does_not_fire_when_explicit_sub_columns_already_present(self) -> None:
        question = _question(sub_columns=[("Q27r1", "Field sales")])

        promoted = apply_grid_categorical_row_matching(
            question,
            ("Q27: Field sales", "Q27: Inside sales"),
        )

        self.assertEqual(promoted["sub_columns"], [("Q27r1", "Field sales")])

    def test_does_not_fire_when_label_pattern_already_promoted(self) -> None:
        question = _question(
            sub_columns=[
                ("Q27: Increased", "Increased"),
                ("Q27: Decreased", "Decreased"),
            ]
        )

        promoted = apply_grid_categorical_row_matching(
            question,
            ("Q27: Field sales", "Q27: Inside sales"),
        )

        self.assertEqual(
            promoted["sub_columns"],
            [("Q27: Increased", "Increased"), ("Q27: Decreased", "Decreased")],
        )

    def test_extracts_row_labels_correctly(self) -> None:
        promoted = apply_grid_categorical_row_matching(
            _question(),
            (
                "Q27: Field sales (direct, face-to-face outside sales reps)",
                "Q27: Specialist sellers",
            ),
        )

        self.assertEqual(
            promoted["sub_columns"],
            [
                (
                    "Q27: Field sales (direct, face-to-face outside sales reps)",
                    "Field sales (direct, face-to-face outside sales reps)",
                ),
                ("Q27: Specialist sellers", "Specialist sellers"),
            ],
        )

    def test_promotes_to_grid_single_select_type(self) -> None:
        spec = _schema_for(
            _question(),
            ["Respondent", "Q27: Field sales", "Q27: Inside sales"],
        ).get_question("Q27")

        self.assertEqual(spec.question_type, QuestionType.GRID_SINGLE_SELECT)

    def test_preserves_original_options_as_grid_scale(self) -> None:
        question = _question()

        promoted = apply_grid_categorical_row_matching(
            question,
            ("Q27: Field sales", "Q27: Inside sales"),
        )

        self.assertEqual(promoted["options"], question["options"])

    def test_handles_em_dash_separator(self) -> None:
        promoted = apply_grid_categorical_row_matching(
            _question(),
            ("Q27 \u2014 Field sales", "Q27 \u2014 Inside sales"),
        )

        self.assertEqual(
            promoted["sub_columns"],
            [("Q27 \u2014 Field sales", "Field sales"), ("Q27 \u2014 Inside sales", "Inside sales")],
        )

    def test_logs_warnings_for_unmatched_columns(self) -> None:
        promoted = apply_grid_categorical_row_matching(
            _question(),
            ("Q27: Field sales", "Q27: Inside sales", "Q27: "),
        )

        self.assertEqual(len(grid_categorical_row_warnings(promoted)), 1)
        self.assertIn("unmatched column", grid_categorical_row_warnings(promoted)[0])

    def test_rating_scale_detection_increased_decreased_same(self) -> None:
        question = _question(
            options=[
                (1, "Increased"),
                (2, "About the same"),
                (3, "Decreased"),
            ]
        )

        promoted = apply_grid_categorical_row_matching(
            question,
            ("Q27: Field sales", "Q27: Inside sales"),
        )

        self.assertEqual(len(promoted["sub_columns"]), 2)

    def test_rating_scale_detection_agree_disagree(self) -> None:
        question = _question(
            options=[
                (1, "Strongly agree"),
                (2, "Somewhat agree"),
                (3, "Disagree"),
            ]
        )

        promoted = apply_grid_categorical_row_matching(
            question,
            ("Q27: Field sales", "Q27: Inside sales"),
        )

        self.assertEqual(len(promoted["sub_columns"]), 2)


if __name__ == "__main__":
    unittest.main()
