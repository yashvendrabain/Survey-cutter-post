from __future__ import annotations

import unittest

from src.question_classifier import classify_questions
from src.models import QuestionType


def _question(
    qid: str,
    text: str,
    *,
    value_range: tuple[int, int] | None,
    options: list[tuple[int | str, str]] | None = None,
    sub_columns: list[tuple[str, str]] | None = None,
) -> dict:
    return {
        "canonical_id": qid,
        "raw_id": qid,
        "question_text": text,
        "type_hint": "values_range",
        "value_range": value_range,
        "options": options or [],
        "sub_columns": sub_columns or [],
        "parent_canonical_id": None,
        "source_row": 1,
        "warnings": [],
    }


class TestNPSClassifier(unittest.TestCase):
    def test_recommend_zero_to_ten_question_is_nps(self) -> None:
        data_map = {
            "questions": [
                _question(
                    "Q19",
                    "On a scale of 0-10, how likely are you to recommend <brand>?",
                    value_range=(0, 10),
                    sub_columns=[("Q19 | Covetrus", "Covetrus")],
                )
            ],
            "source_path": "<test>",
            "sheet_name": "Datamap",
            "total_rows_in_sheet": 1,
            "parser_warnings": [],
        }

        schema = classify_questions(data_map, ["Q19 | Covetrus"])

        self.assertIs(schema.get_question("Q19").question_type, QuestionType.NPS)

    def test_one_to_five_performance_matrix_stays_grid_rated(self) -> None:
        data_map = {
            "questions": [
                _question(
                    "Q22",
                    "On a scale from 1-5, how well does <brand> perform?",
                    value_range=(1, 5),
                    options=[(1, "1"), (2, "2"), (3, "3"), (4, "4"), (5, "5")],
                    sub_columns=[("Q22: Quality | Covetrus", "Quality | Covetrus")],
                )
            ],
            "source_path": "<test>",
            "sheet_name": "Datamap",
            "total_rows_in_sheet": 1,
            "parser_warnings": [],
        }

        schema = classify_questions(data_map, ["Q22: Quality | Covetrus"])

        self.assertIs(schema.get_question("Q22").question_type, QuestionType.GRID_RATED)

    def test_zero_to_ten_without_recommend_is_not_nps(self) -> None:
        data_map = {
            "questions": [
                _question(
                    "Q10",
                    "On a scale of 0-10, how satisfied are you?",
                    value_range=(0, 10),
                    sub_columns=[("Q10 | Brand", "Brand")],
                )
            ],
            "source_path": "<test>",
            "sheet_name": "Datamap",
            "total_rows_in_sheet": 1,
            "parser_warnings": [],
        }

        schema = classify_questions(data_map, ["Q10 | Brand"])

        self.assertIsNot(schema.get_question("Q10").question_type, QuestionType.NPS)


if __name__ == "__main__":
    unittest.main()
