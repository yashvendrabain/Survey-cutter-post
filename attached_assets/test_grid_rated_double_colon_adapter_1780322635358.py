"""Tests for rated grids embedded in double-colon raw headers."""

from __future__ import annotations

import unittest

from src.adapters.grid_rated_double_colon import (
    apply_grid_rated_double_colon_matching,
    grid_rated_double_colon_warnings,
)
from src.models import QuestionType
from src.question_classifier import classify_questions


def _question(
    canonical_id: str = "Q14",
    options: list[tuple[int | str, str]] | None = None,
    sub_columns: list[tuple[str, str]] | None = None,
    type_hint: str = "values_range",
    value_range: tuple[int, int] | None = (0, 100),
) -> dict:
    return {
        "canonical_id": canonical_id,
        "raw_id": canonical_id,
        "question_text": "Enter growth metrics.",
        "type_hint": type_hint,
        "value_range": value_range,
        "options": options if options is not None else [],
        "sub_columns": sub_columns if sub_columns is not None else [],
        "parent_canonical_id": None,
        "source_row": 1,
        "warnings": [],
    }


def _time_metric_columns() -> tuple[str, ...]:
    return (
        "Q14: Planned 2024 :: Revenue_Growth",
        "Q14: Planned 2024 :: Gross margin growth",
        "Q14: Actual 2024 :: Revenue_Growth",
        "Q14: Actual 2024 :: Gross margin growth",
    )


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


class TestGridRatedDoubleColonAdapter(unittest.TestCase):
    def test_fires_when_columns_have_double_colon_separator(self) -> None:
        promoted = apply_grid_rated_double_colon_matching(
            _question(),
            _time_metric_columns(),
        )

        self.assertEqual(len(promoted["sub_columns"]), 4)

    def test_does_not_fire_when_no_double_colon_in_columns(self) -> None:
        promoted = apply_grid_rated_double_colon_matching(
            _question(),
            ("Q14: Planned 2024 - Revenue_Growth", "Q14: Actual 2024 - Revenue_Growth"),
        )

        self.assertEqual(promoted["sub_columns"], [])

    def test_does_not_fire_when_question_has_full_categorical_options(self) -> None:
        question = _question(
            options=[
                (1, "Low"),
                (2, "Medium"),
                (3, "High"),
            ]
        )

        promoted = apply_grid_rated_double_colon_matching(
            question,
            _time_metric_columns(),
        )

        self.assertEqual(promoted["sub_columns"], [])

    def test_extracts_dim1_and_dim2_labels_correctly(self) -> None:
        promoted = apply_grid_rated_double_colon_matching(
            _question(),
            _time_metric_columns(),
        )

        self.assertEqual(
            promoted["sub_columns"][0],
            (
                "Q14: Planned 2024 :: Revenue_Growth",
                "Planned 2024 :: Revenue_Growth",
            ),
        )

    def test_promotes_to_grid_rated_type(self) -> None:
        spec = _schema_for(
            _question(),
            ["Respondent", *_time_metric_columns()],
        ).get_question("Q14")

        self.assertEqual(spec.question_type, QuestionType.GRID_RATED)

    def test_preserves_numeric_value_range(self) -> None:
        promoted = apply_grid_rated_double_colon_matching(
            _question(value_range=(-100, 100)),
            _time_metric_columns(),
        )

        self.assertEqual(promoted["value_range"], (-100, 100))

    def test_handles_partial_grid_with_warning(self) -> None:
        promoted = apply_grid_rated_double_colon_matching(
            _question(),
            (
                "Q14: Planned 2024 :: Revenue_Growth",
                "Q14: Planned 2024 :: Gross margin growth",
                "Q14: Actual 2024 :: Revenue_Growth",
            ),
        )

        warnings = grid_rated_double_colon_warnings(promoted)
        self.assertEqual(len(warnings), 1)
        self.assertIn("partial grid", warnings[0])

    def test_2d_grid_detection_time_x_metric(self) -> None:
        promoted = apply_grid_rated_double_colon_matching(
            _question(),
            _time_metric_columns(),
        )

        self.assertEqual(grid_rated_double_colon_warnings(promoted), [])
        self.assertEqual(len(promoted["sub_columns"]), 4)

    def test_logs_warnings_for_non_grid_pattern(self) -> None:
        promoted = apply_grid_rated_double_colon_matching(
            _question(),
            (
                "Q14: Planned 2024 :: Revenue_Growth",
                "Q14: Planned 2024 :: Gross margin growth",
            ),
        )

        warnings = grid_rated_double_colon_warnings(promoted)
        self.assertEqual(len(warnings), 1)
        self.assertIn("non-grid pattern", warnings[0])

    def test_does_not_consume_columns_already_in_sub_columns(self) -> None:
        question = _question(sub_columns=[("Q14r1", "Revenue")])

        promoted = apply_grid_rated_double_colon_matching(
            question,
            _time_metric_columns(),
        )

        self.assertEqual(promoted["sub_columns"], [("Q14r1", "Revenue")])

    def test_handles_blank_dim1(self) -> None:
        raw_columns = (
            "Q46:  :: Sales plays 1",
            "Q46:  :: Sales plays 2",
            "Q46:  :: Sales plays 3",
        )

        promoted = apply_grid_rated_double_colon_matching(
            _question("Q46"),
            raw_columns,
        )
        spec = _schema_for(
            _question("Q46"),
            ["Respondent", *raw_columns],
        ).get_question("Q46")

        self.assertEqual(
            promoted["sub_columns"],
            [
                ("Q46:  :: Sales plays 1", "Sales plays 1"),
                ("Q46:  :: Sales plays 2", "Sales plays 2"),
                ("Q46:  :: Sales plays 3", "Sales plays 3"),
            ],
        )
        self.assertEqual(grid_rated_double_colon_warnings(promoted), [])
        self.assertEqual(spec.question_type, QuestionType.GRID_RATED)

    def test_handles_dim1_with_only_help_keywords(self) -> None:
        promoted = apply_grid_rated_double_colon_matching(
            _question("Q46"),
            (
                "Q46:  :: Sales plays 1",
                "Q46:  :: Sales plays 2",
                "Q46: I don't know",
                "Q46: I don't know :: Sales plays helper",
            ),
        )

        self.assertEqual(
            promoted["sub_columns"],
            [
                ("Q46:  :: Sales plays 1", "Sales plays 1"),
                ("Q46:  :: Sales plays 2", "Sales plays 2"),
            ],
        )


if __name__ == "__main__":
    unittest.main()
