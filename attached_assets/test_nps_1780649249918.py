from __future__ import annotations

import unittest

import pandas as pd

from src.calculation_log import CalculationLog
from src.models import NPSResult, QuestionSpec, QuestionType
from src.single_cut.nps import compute_nps


def _nps_spec(columns: tuple[str, ...] = ("Q1 | Brand",)) -> QuestionSpec:
    labels = {column: column.split("|", 1)[-1].strip() for column in columns}
    return QuestionSpec(
        question_id="Q1",
        canonical_id="Q1",
        question_text="On a scale of 0-10, how likely are you to recommend Brand?",
        question_type=QuestionType.NPS,
        raw_columns=columns,
        option_map=labels,
        value_range=(0, 10),
        grid_column_labels=labels,
    )


class TestNPSCalculator(unittest.TestCase):
    def test_canonical_example_nps_is_30(self) -> None:
        values = [9] * 25 + [10] * 25 + [7] * 15 + [8] * 15 + [0] * 20
        df = pd.DataFrame({"Q1 | Brand": values})
        log = CalculationLog()

        result = compute_nps(_nps_spec(), df, log)

        self.assertIsInstance(result, NPSResult)
        entity = result.entities[0]
        self.assertEqual(entity.promoters, 50)
        self.assertEqual(entity.passives, 30)
        self.assertEqual(entity.detractors, 20)
        self.assertEqual(entity.valid_n, 100)
        self.assertAlmostEqual(entity.nps_score, 30.0)
        self.assertEqual(len(log), 8)

    def test_missing_and_out_of_range_are_excluded_from_valid_n(self) -> None:
        df = pd.DataFrame({"Q1 | Brand": [10, 9, 8, 7, 6, 11, -1, 5.5, "abc", None]})
        result = compute_nps(_nps_spec(), df, CalculationLog())

        entity = result.entities[0]
        self.assertEqual(entity.valid_n, 5)
        self.assertEqual(entity.missing_n, 5)
        self.assertEqual(entity.promoters, 2)
        self.assertEqual(entity.passives, 2)
        self.assertEqual(entity.detractors, 1)
        self.assertAlmostEqual(entity.nps_score, 20.0)

    def test_per_entity_columns_are_scored_independently(self) -> None:
        df = pd.DataFrame(
            {
                "Q1 | A": [9, 10, 8, 6],
                "Q1 | B": [0, 6, 7, 10],
            }
        )
        result = compute_nps(_nps_spec(("Q1 | A", "Q1 | B")), df, CalculationLog())

        self.assertEqual([entity.entity_label for entity in result.entities], ["A", "B"])
        self.assertAlmostEqual(result.entities[0].nps_score, 25.0)
        self.assertAlmostEqual(result.entities[1].nps_score, -25.0)


if __name__ == "__main__":
    unittest.main()
