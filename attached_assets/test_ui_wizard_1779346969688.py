"""Tests for setup wizard state helpers."""

from __future__ import annotations

from datetime import datetime, timezone
import unittest

import pandas as pd

from src.models import QuestionSpec, QuestionType, SurveySchema
from src.ui.wizard import (
    apply_wizard_schema_overrides,
    category_assignments_from_themes,
    distinct_value_preview,
    eligible_filter_question_ids,
    normalise_custom_filter_count,
    normalise_per_question_filter_count,
    selected_demographics_from_schema,
    themes_from_category_assignments,
    themes_from_wizard_assignments,
)


def wizard_schema() -> SurveySchema:
    return SurveySchema(
        questions=(
            QuestionSpec(
                question_id="[Q1]",
                canonical_id="Q1",
                question_text="Country",
                question_type=QuestionType.SINGLE_SELECT,
                raw_columns=("Q1",),
                option_map={1: "USA", 2: "India"},
                is_demographic=True,
            ),
            QuestionSpec(
                question_id="[Q2]",
                canonical_id="Q2",
                question_text="Revenue growth",
                question_type=QuestionType.SINGLE_SELECT,
                raw_columns=("Q2",),
                option_map={1: "Up", 2: "Down"},
            ),
            QuestionSpec(
                question_id="[Q3]",
                canonical_id="Q3",
                question_text="Open text",
                question_type=QuestionType.OPEN_TEXT,
                raw_columns=("Q3",),
                option_map={},
            ),
        ),
        respondent_id_column="record",
        total_respondents=10,
        source_datamap_path="datamap.xlsx",
        source_rawdata_path="raw.xlsx",
        parsed_at=datetime(2026, 5, 19, tzinfo=timezone.utc),
    )


class TestWizardHelpers(unittest.TestCase):
    def test_category_assignments_from_ai_themes(self) -> None:
        schema = wizard_schema()
        assignments = category_assignments_from_themes(
            schema,
            {"themes": [{"name": "Growth", "question_ids": ["Q2"]}]},
        )

        self.assertEqual(assignments["Q2"], "Growth")
        self.assertEqual(assignments["Q1"], "Demographics")

    def test_themes_from_assignments_preserves_schema_order(self) -> None:
        schema = wizard_schema()
        themes = themes_from_category_assignments(
            schema,
            {"Q2": "Growth", "Q1": "Market"},
        )

        self.assertEqual(
            themes["themes"],
            [
                {"name": "Market", "question_ids": ["Q1"]},
                {"name": "Growth", "question_ids": ["Q2"]},
            ],
        )

    def test_themes_from_wizard_assignments_matches_exporter_payload(self) -> None:
        schema = wizard_schema()

        self.assertEqual(
            themes_from_wizard_assignments(schema, {"Q1": "Market"}),
            themes_from_category_assignments(schema, {"Q1": "Market"}),
        )

    def test_apply_overrides_removes_unassigned_and_sets_demographics(self) -> None:
        schema = wizard_schema()
        updated = apply_wizard_schema_overrides(
            schema,
            {"Q1": "Market"},
            ["Q2"],
        )

        self.assertTrue(updated.get_question("Q1").analysis_eligible)
        self.assertFalse(updated.get_question("Q1").is_demographic)
        self.assertFalse(updated.get_question("Q2").analysis_eligible)
        self.assertEqual(updated.get_question("Q2").exclusion_reason, "removed in setup wizard")

    def test_eligible_filter_questions_exclude_open_text(self) -> None:
        self.assertEqual(eligible_filter_question_ids(wizard_schema()), ["Q1", "Q2"])

    def test_demographic_defaults_and_distinct_value_preview(self) -> None:
        schema = wizard_schema()
        dataframe = pd.DataFrame({"Q1": ["USA", "Canada", "USA", None]})

        self.assertEqual(selected_demographics_from_schema(schema), ["Q1"])
        self.assertEqual(
            distinct_value_preview(dataframe, schema.get_question("Q1")),
            "2 distinct values: USA, Canada",
        )

    def test_filter_count_normalisation(self) -> None:
        self.assertEqual(normalise_custom_filter_count(9), 5)
        self.assertEqual(normalise_custom_filter_count(-1), 0)
        self.assertEqual(normalise_per_question_filter_count(7), 3)
        self.assertEqual(normalise_per_question_filter_count(-1), 0)


if __name__ == "__main__":
    unittest.main()
