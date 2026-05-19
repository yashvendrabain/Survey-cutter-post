"""Tests for the public single-cut engine orchestrator."""

from __future__ import annotations

from datetime import datetime, timezone
import unittest

import pandas as pd

from src.calculation_log import CalculationLog
from src.models import (
    DenominatorPolicy,
    GridSingleSelectResult,
    MultiSelectResult,
    NumericResult,
    QuestionSpec,
    QuestionType,
    SingleSelectResult,
    SurveySchema,
)
from src.single_cut.engine import compute_single_cuts
from tests.conftest import GOLDEN_30_RESPONDENTS_PATH


GRID_OPTION_MAP = {
    1: "Strongly disagree",
    2: "Disagree",
    3: "Agree",
    4: "Strongly agree",
}
GRID_ROW_LABELS = {
    "Q_GRID_1r1": "First grid row",
    "Q_GRID_1r2": "Second grid row",
    "Q_GRID_1r3": "Third grid row",
}


def load_golden() -> pd.DataFrame:
    return pd.read_csv(GOLDEN_30_RESPONDENTS_PATH)


def make_schema() -> SurveySchema:
    questions = (
        QuestionSpec(
            question_id="[Q_SS_1]",
            canonical_id="Q_SS_1",
            question_text="Single select",
            question_type=QuestionType.SINGLE_SELECT,
            raw_columns=("Q_SS_1",),
            option_map={1: "Yes", 2: "No"},
        ),
        QuestionSpec(
            question_id="[Q_MS_1]",
            canonical_id="Q_MS_1",
            question_text="Multi select",
            question_type=QuestionType.MULTI_SELECT_BINARY,
            raw_columns=("Q_MS_1r1", "Q_MS_1r2", "Q_MS_1r3"),
            option_map={
                "Q_MS_1r1": "First",
                "Q_MS_1r2": "Second",
                "Q_MS_1r3": "Third",
            },
            value_range=(0, 1),
        ),
        QuestionSpec(
            question_id="[Q_NUM_1]",
            canonical_id="Q_NUM_1",
            question_text="Direct numeric",
            question_type=QuestionType.DIRECT_NUMERIC,
            raw_columns=("Q_NUM_1",),
            option_map={},
        ),
        QuestionSpec(
            question_id="[Q_ALLOC_2]",
            canonical_id="Q_ALLOC_2",
            question_text="Allocation",
            question_type=QuestionType.NUMERIC_ALLOCATION,
            raw_columns=("Q_ALLOC_2r1", "Q_ALLOC_2r2", "Q_ALLOC_2r3"),
            option_map={
                "Q_ALLOC_2r1": "First",
                "Q_ALLOC_2r2": "Second",
                "Q_ALLOC_2r3": "Third",
            },
            value_range=(0, 999),
        ),
        QuestionSpec(
            question_id="[Q_GRID_1]",
            canonical_id="Q_GRID_1",
            question_text="Grid",
            question_type=QuestionType.GRID_SINGLE_SELECT,
            raw_columns=("Q_GRID_1r1", "Q_GRID_1r2", "Q_GRID_1r3"),
            option_map=GRID_OPTION_MAP,
            value_range=(1, 4),
            grid_row_labels=GRID_ROW_LABELS,
        ),
        QuestionSpec(
            question_id="[Q_TEXT]",
            canonical_id="Q_TEXT",
            question_text="Open text",
            question_type=QuestionType.OPEN_TEXT,
            raw_columns=("Q_TEXT",),
            option_map={},
        ),
        QuestionSpec(
            question_id="record",
            canonical_id="record",
            question_text="Respondent ID",
            question_type=QuestionType.METADATA_OR_ID,
            raw_columns=("respondent_id",),
            option_map={},
        ),
        QuestionSpec(
            question_id="[Q_UNKNOWN]",
            canonical_id="Q_UNKNOWN",
            question_text="Unknown",
            question_type=QuestionType.UNKNOWN,
            raw_columns=(),
            option_map={},
        ),
        QuestionSpec(
            question_id="[Q_INELIGIBLE]",
            canonical_id="Q_INELIGIBLE",
            question_text="Ineligible",
            question_type=QuestionType.SINGLE_SELECT,
            raw_columns=("Q_SS_1",),
            option_map={1: "Yes", 2: "No"},
            analysis_eligible=False,
            exclusion_reason="manual exclusion",
        ),
        QuestionSpec(
            question_id="[Q_MISSING]",
            canonical_id="Q_MISSING",
            question_text="Missing column",
            question_type=QuestionType.SINGLE_SELECT,
            raw_columns=("Q_MISSING",),
            option_map={1: "Yes", 2: "No"},
        ),
    )
    return SurveySchema(
        questions=questions,
        respondent_id_column="respondent_id",
        total_respondents=30,
        source_datamap_path="datamap.xlsx",
        source_rawdata_path="raw.csv",
        parsed_at=datetime(2026, 5, 4, tzinfo=timezone.utc),
    )


class TestSingleCutEngine(unittest.TestCase):
    def test_engine_processes_all_eligible_questions(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()

        results, skips = compute_single_cuts(make_schema(), dataframe, log)

        self.assertEqual(
            [result.question_id for result in results],
            ["Q_SS_1", "Q_MS_1", "Q_NUM_1", "Q_ALLOC_2", "Q_GRID_1"],
        )
        self.assertEqual(len(skips), 5)

    def test_engine_skips_open_text(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()

        _, skips = compute_single_cuts(make_schema(), dataframe, log)

        skip = next(record for record in skips if record.canonical_id == "Q_TEXT")
        self.assertEqual(skip.skip_reason, "unsupported_type: OPEN_TEXT")

    def test_engine_skips_metadata(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()

        _, skips = compute_single_cuts(make_schema(), dataframe, log)

        skip = next(record for record in skips if record.canonical_id == "record")
        self.assertEqual(skip.skip_reason, "unsupported_type: METADATA_OR_ID")

    def test_engine_skips_unknown(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()

        _, skips = compute_single_cuts(make_schema(), dataframe, log)

        skip = next(record for record in skips if record.canonical_id == "Q_UNKNOWN")
        self.assertEqual(skip.skip_reason, "unsupported_type: UNKNOWN")

    def test_engine_skips_ineligible_with_reason(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()

        _, skips = compute_single_cuts(make_schema(), dataframe, log)

        skip = next(record for record in skips if record.canonical_id == "Q_INELIGIBLE")
        self.assertEqual(skip.skip_reason, "ineligible")
        self.assertEqual(skip.details, "manual exclusion")

    def test_engine_catches_exception_returns_skip_record(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()

        _, skips = compute_single_cuts(make_schema(), dataframe, log)

        skip = next(record for record in skips if record.canonical_id == "Q_MISSING")
        self.assertEqual(skip.skip_reason, "calculation_error")
        self.assertIn("ValueError: raw column not found in data", skip.details)

    def test_engine_skips_question_when_all_raw_columns_empty(self) -> None:
        schema = SurveySchema(
            questions=(
                QuestionSpec(
                    question_id="[Q_EMPTY]",
                    canonical_id="Q_EMPTY",
                    question_text="Empty grid",
                    question_type=QuestionType.GRID_SINGLE_SELECT,
                    raw_columns=("Q_EMPTYr1", "Q_EMPTYr2"),
                    option_map={1: "Decision maker", 2: "Influencer"},
                    value_range=(1, 2),
                    grid_row_labels={
                        "Q_EMPTYr1": "IT",
                        "Q_EMPTYr2": "Finance",
                    },
                ),
            ),
            respondent_id_column="respondent_id",
            total_respondents=3,
            source_datamap_path="datamap.xlsx",
            source_rawdata_path="raw.csv",
            parsed_at=datetime(2026, 5, 4, tzinfo=timezone.utc),
        )
        dataframe = pd.DataFrame(
            {
                "respondent_id": [1, 2, 3],
                "Q_EMPTYr1": [pd.NA, pd.NA, pd.NA],
                "Q_EMPTYr2": [pd.NA, pd.NA, pd.NA],
            }
        )
        log = CalculationLog()

        results, skips = compute_single_cuts(schema, dataframe, log)

        self.assertEqual(results, [])
        self.assertEqual(len(skips), 1)
        self.assertEqual(skips[0].canonical_id, "Q_EMPTY")
        self.assertEqual(skips[0].skip_reason, "all raw columns empty in dataset")

    def test_engine_returns_results_in_schema_order(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()

        results, _ = compute_single_cuts(make_schema(), dataframe, log)

        self.assertEqual(
            [result.question_id for result in results],
            ["Q_SS_1", "Q_MS_1", "Q_NUM_1", "Q_ALLOC_2", "Q_GRID_1"],
        )

    def test_engine_dispatches_correctly_per_type(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()

        results, _ = compute_single_cuts(make_schema(), dataframe, log)

        self.assertIsInstance(results[0], SingleSelectResult)
        self.assertIsInstance(results[1], MultiSelectResult)
        self.assertIsInstance(results[2], NumericResult)
        self.assertIsInstance(results[3], NumericResult)
        self.assertIsInstance(results[4], GridSingleSelectResult)

    def test_engine_filter_mask_propagated_to_calculators(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()
        filter_mask = dataframe.index < 10

        results, _ = compute_single_cuts(
            make_schema(),
            dataframe,
            log,
            filter_mask=filter_mask,
            filter_expr="first ten respondents",
        )

        self.assertEqual([result.valid_n for result in results], [10, 10, 10, 10, 10])
        for result in results:
            self.assertEqual(result.audit_records[0].filter_expr, "first ten respondents")

    def test_engine_audit_log_populated_for_all_results(self) -> None:
        dataframe = load_golden()
        log = CalculationLog()

        compute_single_cuts(make_schema(), dataframe, log)

        self.assertEqual(len(log), 18)
        metric_names = [record.metric_name for record in log.all_records()]
        self.assertEqual(metric_names.count("rate_per_value"), 4)
        self.assertEqual(metric_names.count("selection_rate"), 1)
        self.assertEqual(metric_names.count("numeric_summary"), 1)
        self.assertEqual(metric_names.count("numeric_std"), 1)
        self.assertEqual(metric_names.count("numeric_p25"), 1)
        self.assertEqual(metric_names.count("numeric_p50"), 1)
        self.assertEqual(metric_names.count("numeric_p75"), 1)
        self.assertEqual(metric_names.count("allocation_summary"), 1)
        self.assertEqual(metric_names.count("numeric_allocation_mean"), 3)
        self.assertEqual(metric_names.count("numeric_allocation_median"), 3)
        self.assertEqual(metric_names.count("grid_overall"), 1)


if __name__ == "__main__":
    unittest.main()
