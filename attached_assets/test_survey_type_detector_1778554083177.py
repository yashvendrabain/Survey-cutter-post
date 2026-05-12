"""Tests for deterministic survey type detection."""

from __future__ import annotations

import unittest

import pandas as pd

from src.survey_type_detector import detect_survey_type


def _question(
    qid: str,
    text: str,
    *,
    type_hint: str | None = "values_range",
    value_range: tuple[int, int] | None = (1, 5),
    options: list[tuple[int, str]] | None = None,
    sub_columns: list[tuple[str, str]] | None = None,
    question_type: str | None = None,
) -> dict:
    question = {
        "canonical_id": qid,
        "raw_id": qid,
        "question_text": text,
        "type_hint": type_hint,
        "value_range": value_range,
        "options": options if options is not None else [(1, "Low"), (2, "High")],
        "sub_columns": sub_columns if sub_columns is not None else [],
        "parent_canonical_id": None,
        "source_row": 1,
        "warnings": [],
    }
    if question_type is not None:
        question["question_type"] = question_type
    return question


def _schema(questions: list[dict]) -> dict:
    return {
        "questions": questions,
        "source_path": "synthetic.xlsx",
        "sheet_name": "Sheet1",
        "total_rows_in_sheet": len(questions),
        "parser_warnings": [],
    }


def _df(rows: int = 10) -> pd.DataFrame:
    return pd.DataFrame({"record": list(range(rows))})


class TestSurveyTypeDetector(unittest.TestCase):
    def test_ltb_detection(self) -> None:
        result = detect_survey_type(
            _schema(
                [
                    _question("Q1", "Which vendors were in your consideration set?"),
                    _question("Q2", "Which vendor made the shortlist?"),
                    _question("Q3", "Which selected vendor was the winner?"),
                ]
            ),
            _df(),
        )

        self.assertEqual(result.survey_type, "purchase_ltb")
        self.assertEqual(result.confidence, 0.95)
        self.assertEqual(result.outcome_question_id, "Q3")

    def test_growth_strategy_detection(self) -> None:
        result = detect_survey_type(
            _schema(
                [
                    _question("Q1", "What was your revenue growth last year?"),
                    _question("Q2", "Did you achieve your growth target?"),
                    _question("Q3", "How much market share gain did you see?"),
                ]
            ),
            _df(),
        )

        self.assertEqual(result.survey_type, "growth_strategy")
        self.assertEqual(result.outcome_question_id, "Q1")
        self.assertGreaterEqual(result.confidence, 0.9)

    def test_growth_agenda_short_circuit(self) -> None:
        result = detect_survey_type(
            _schema(
                [
                    _question(
                        "Q1",
                        "How much did your organization's revenue change in 2023?",
                    ),
                    _question(
                        "Q2",
                        "How likely are you to recommend us?",
                        value_range=(0, 10),
                    ),
                    _question("Q3", "Which vendor did you select as the winner?"),
                ]
            ),
            _df(),
        )

        self.assertEqual(result.survey_type, "growth_strategy")
        self.assertGreaterEqual(result.confidence, 0.95)
        self.assertEqual(result.outcome_question_id, "Q1")

    def test_nps_detection(self) -> None:
        result = detect_survey_type(
            _schema(
                [
                    _question(
                        "Q1",
                        "How likely are you to recommend us?",
                        value_range=(0, 10),
                    )
                ]
            ),
            _df(),
        )

        self.assertEqual(result.survey_type, "nps")
        self.assertEqual(result.outcome_question_id, "Q1")
        self.assertEqual(result.confidence, 0.85)

    def test_employee_engagement_detection(self) -> None:
        result = detect_survey_type(
            _schema(
                [
                    _question("Q1", "What is your employee satisfaction rating?"),
                    _question("Q2", "How motivated are you at work?"),
                ]
            ),
            _df(),
        )

        self.assertEqual(result.survey_type, "employee_engagement")
        self.assertEqual(result.confidence, 0.85)

    def test_digital_maturity_detection(self) -> None:
        result = detect_survey_type(
            _schema(
                [
                    _question("Q1", "How advanced is your digital transformation?"),
                    _question("Q2", "What is your AI adoption stage?"),
                    _question("Q3", "How mature is your cloud migration?"),
                ]
            ),
            _df(),
        )

        self.assertEqual(result.survey_type, "digital_maturity")
        self.assertEqual(result.confidence, 0.8)

    def test_opinion_fallback(self) -> None:
        result = detect_survey_type(
            _schema(
                [
                    _question("Q1", "Which color do you prefer?"),
                    _question("Q2", "Which format do you like most?"),
                ]
            ),
            _df(),
        )

        self.assertEqual(result.survey_type, "opinion")
        self.assertEqual(result.confidence, 0.3)
        self.assertIsNone(result.outcome_question_id)

    def test_empty_schema(self) -> None:
        result = detect_survey_type(_schema([]), _df())

        self.assertEqual(result.survey_type, "unknown")
        self.assertEqual(result.confidence, 0.0)

    def test_all_eligible_questions_coverage(self) -> None:
        questions = [
            _question(f"Q{i}", f"Generic measurable question {i}")
            for i in range(1, 19)
        ]
        questions.append(_question("Q19", "Open text comment", type_hint="open_text"))
        questions.append(_question("Q20", "Metadata record id", question_type="metadata_or_id"))

        result = detect_survey_type(_schema(questions), _df())

        self.assertEqual(len(result.all_eligible_questions), 18)
        self.assertTrue(
            all(option.question_id not in {"Q19", "Q20"} for option in result.all_eligible_questions)
        )

    def test_relevance_scoring_hierarchy(self) -> None:
        result = detect_survey_type(
            _schema(
                [
                    _question("Q1", "Which selected vendor was the winner?"),
                    _question("Q2", "How would you rate performance?"),
                    _question("Q3", "Which option do you prefer?"),
                    _question("Q4", "Which region are you in?"),
                    _question("Q5", "Which vendors were in your consideration set?"),
                ]
            ),
            _df(),
        )

        scores = {option.question_id: option.relevance_score for option in result.all_eligible_questions}
        self.assertGreater(scores["Q1"], scores["Q2"])
        self.assertGreater(scores["Q2"], scores["Q3"])
        self.assertGreater(scores["Q3"], scores["Q4"])

    def test_multi_signal_confidence_boost(self) -> None:
        two_signal = detect_survey_type(
            _schema(
                [
                    _question("Q1", "What was your sales growth?"),
                    _question("Q2", "Did you achieve your growth target?"),
                ]
            ),
            _df(),
        )
        three_signal = detect_survey_type(
            _schema(
                [
                    _question("Q1", "What was your sales growth?"),
                    _question("Q2", "Did you achieve your growth target?"),
                    _question("Q3", "How much market share gain did you see?"),
                ]
            ),
            _df(),
        )

        self.assertEqual(two_signal.survey_type, "growth_strategy")
        self.assertEqual(three_signal.survey_type, "growth_strategy")
        self.assertGreater(three_signal.confidence, two_signal.confidence)


if __name__ == "__main__":
    unittest.main()
