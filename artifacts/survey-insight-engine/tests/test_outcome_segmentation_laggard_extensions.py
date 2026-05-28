from __future__ import annotations

from datetime import datetime, timezone
import unittest

import pandas as pd

from src.models import (
    DenominatorPolicy,
    QuestionSpec,
    QuestionType,
    SegmentDefinition,
    SurveySchema,
)
from src.outcome_segmentation import _build_segment_masks


def _schema(question: QuestionSpec) -> SurveySchema:
    return SurveySchema(
        questions=(question,),
        respondent_id_column="rid",
        total_respondents=4,
        source_datamap_path="map.xlsx",
        source_rawdata_path="raw.xlsx",
        parsed_at=datetime.now(timezone.utc),
    )


class TestOutcomeSegmentationLaggardExtensions(unittest.TestCase):
    def test_segment_definition_accepts_laggard_fields(self) -> None:
        definition = SegmentDefinition(
            outcome_question_id="QO",
            segment_mode="categorical",
            winner_values=(1,),
            laggard_values=(2,),
            laggard_threshold=10.0,
            laggard_threshold_direction="lte",
            laggard_label="Laggards",
        )
        self.assertEqual(definition.laggard_values, (2,))
        self.assertEqual(definition.laggard_threshold, 10.0)
        self.assertEqual(definition.laggard_threshold_direction, "lte")

    def test_categorical_without_laggard_values_keeps_inverse_behavior(self) -> None:
        question = QuestionSpec(
            question_id="QO",
            canonical_id="QO",
            question_text="Outcome",
            question_type=QuestionType.SINGLE_SELECT,
            raw_columns=("QO",),
            option_map={1: "Winner", 2: "Laggard", 3: "Middle"},
            denominator_policy=DenominatorPolicy.VALID_RESPONSES,
        )
        df = pd.DataFrame({"QO": [1, 2, 3, None]})
        winner, laggard, valid, _warnings = _build_segment_masks(
            df,
            question,
            SegmentDefinition("QO", "categorical", winner_values=(1,)),
        )
        self.assertEqual(winner.tolist(), [True, False, False, False])
        self.assertEqual(laggard.tolist(), [False, True, True, False])
        self.assertEqual(valid.tolist(), [True, True, True, False])

    def test_laggard_mask_honors_explicit_laggard_values(self) -> None:
        question = QuestionSpec(
            question_id="QO",
            canonical_id="QO",
            question_text="Outcome",
            question_type=QuestionType.SINGLE_SELECT,
            raw_columns=("QO",),
            option_map={1: "Winner", 2: "Laggard", 3: "Middle"},
            denominator_policy=DenominatorPolicy.VALID_RESPONSES,
        )
        df = pd.DataFrame({"QO": [1, 2, 3, None]})
        _winner, laggard, _valid, _warnings = _build_segment_masks(
            df,
            question,
            SegmentDefinition(
                "QO",
                "categorical",
                winner_values=(1,),
                laggard_values=(2,),
            ),
        )
        self.assertEqual(laggard.tolist(), [False, True, False, False])


if __name__ == "__main__":
    unittest.main()
