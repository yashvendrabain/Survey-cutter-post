from __future__ import annotations

from datetime import datetime, timezone
import unittest

import pandas as pd

from src.filter_options import (
    build_filter_specs_from_selection,
    cross_cut_question_options,
    filter_mask_for_spec,
    filter_question_options,
)
from src.models import FilterSpec, QuestionSpec, QuestionType, SurveySchema


def make_schema() -> SurveySchema:
    return SurveySchema(
        questions=(
            QuestionSpec(
                question_id="[Q_STR]",
                canonical_id="Q_STR",
                question_text="String-coded single select",
                question_type=QuestionType.SINGLE_SELECT,
                raw_columns=("Q_STR",),
                option_map={"A": "Alpha", "B": "Beta"},
            ),
            QuestionSpec(
                question_id="[Q_MS]",
                canonical_id="Q_MS",
                question_text="Multi select",
                question_type=QuestionType.MULTI_SELECT_BINARY,
                raw_columns=("Q_MSr1", "Q_MSr2", "Q_MS_count"),
                option_map={
                    "Q_MSr1": "First",
                    "Q_MSr2": "Second",
                    "Q_MS_count": "Computed(Count choices)",
                },
            ),
            QuestionSpec(
                question_id="[Q_NPS]",
                canonical_id="Q_NPS",
                question_text="Recommend",
                question_type=QuestionType.NPS,
                raw_columns=("Q_NPS",),
                option_map={},
            ),
            QuestionSpec(
                question_id="[Q_GRID]",
                canonical_id="Q_GRID",
                question_text="Grid single",
                question_type=QuestionType.GRID_SINGLE_SELECT,
                raw_columns=("Q_GRIDr1", "Q_GRIDr2"),
                option_map={1: "Low", 2: "High"},
                grid_row_labels={"Q_GRIDr1": "Row one", "Q_GRIDr2": "Row two"},
            ),
            QuestionSpec(
                question_id="[Q_GB]",
                canonical_id="Q_GB",
                question_text="Grid binary",
                question_type=QuestionType.GRID_BINARY_SELECT,
                raw_columns=("Q_GBr1", "Q_GBr2"),
                option_map={"Q_GBr1": "Binary one", "Q_GBr2": "Binary two"},
                grid_row_labels={"Q_GBr1": "Binary one", "Q_GBr2": "Binary two"},
            ),
            QuestionSpec(
                question_id="[Q_RANK]",
                canonical_id="Q_RANK",
                question_text="Rank order",
                question_type=QuestionType.RANK_ORDER,
                raw_columns=("Q_RANKr1", "Q_RANKr2"),
                option_map={"Q_RANKr1": "Rank one", "Q_RANKr2": "Rank two"},
            ),
        ),
        respondent_id_column="respondent_id",
        total_respondents=4,
        source_datamap_path="datamap.xlsx",
        source_rawdata_path="raw.csv",
        parsed_at=datetime(2026, 6, 5, tzinfo=timezone.utc),
    )


class TestFilterOptions(unittest.TestCase):
    def test_eligible_filter_questions_cover_added_types(self) -> None:
        options = {option.question_id: option for option in filter_question_options(make_schema())}

        self.assertIn("Q_STR", options)
        self.assertIn("Q_MS", options)
        self.assertIn("Q_NPS", options)
        self.assertIn("Q_GRID", options)
        self.assertIn("Q_GB", options)
        self.assertIn("Promoter", [choice.value for choice in options["Q_NPS"].values])
        self.assertIn("Q_MSr1", [choice.filter_question_id for choice in options["Q_MS"].values])
        self.assertIn("Q_GRIDr1", [choice.filter_question_id for choice in options["Q_GRID"].values])

    def test_multi_select_filter_options_exclude_computed_columns(self) -> None:
        options = {option.question_id: option for option in filter_question_options(make_schema())}

        values = [choice.value for choice in options["Q_MS"].values]

        self.assertEqual(values, ["Q_MSr1", "Q_MSr2"])
        self.assertNotIn("Q_MS_count", values)

    def test_cross_cut_options_include_new_metric_types(self) -> None:
        options = {option.question_id: option for option in cross_cut_question_options(make_schema())}

        self.assertEqual(options["Q_MS"].question_type, QuestionType.MULTI_SELECT_BINARY)
        self.assertEqual(options["Q_GB"].question_type, QuestionType.GRID_BINARY_SELECT)
        self.assertEqual(options["Q_RANK"].question_type, QuestionType.RANK_ORDER)

    def test_multi_select_parent_selection_builds_subcolumn_filter(self) -> None:
        specs = build_filter_specs_from_selection(make_schema(), "Q_MS", ["Q_MSr1"])

        self.assertEqual(specs, [FilterSpec("Q_MSr1", 1)])

    def test_filter_mask_handles_selected_subcolumn_and_nps_bucket(self) -> None:
        schema = make_schema()
        dataframe = pd.DataFrame(
            {
                "Q_MSr1": [1, 0, "Selected", None],
                "Q_NPS": [10, 8, 4, None],
            }
        )

        selected_mask = filter_mask_for_spec(dataframe, schema, FilterSpec("Q_MSr1", 1))
        promoter_mask = filter_mask_for_spec(dataframe, schema, FilterSpec("Q_NPS", "Promoter"))

        self.assertEqual(selected_mask.tolist(), [True, False, True, False])
        self.assertEqual(promoter_mask.tolist(), [True, False, False, False])


if __name__ == "__main__":
    unittest.main()
