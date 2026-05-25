import unittest
from datetime import datetime, timezone

from src.models import (
    AuditRecord,
    DenominatorPolicy,
    DifferentiatorResult,
    GridSingleSelectResult,
    MultiSelectResult,
    NumericResult,
    OutcomeSegmentationResult,
    QuestionSpec,
    QuestionType,
    SegmentDefinition,
    SingleCutResult,
    SingleSelectResult,
    SurveySchema,
    WinnerProfile,
)


UTC_NOW = datetime(2026, 5, 3, 12, 0, tzinfo=timezone.utc)


def make_question_spec(canonical_id: str = "Q1") -> QuestionSpec:
    return QuestionSpec(
        question_id=f"[{canonical_id}]",
        canonical_id=canonical_id,
        question_text="Would you recommend the product?",
        question_type=QuestionType.SINGLE_SELECT,
        raw_columns=(canonical_id,),
        option_map={1: "Yes", 2: "No"},
        value_range=(1, 2),
    )


def make_audit_record() -> AuditRecord:
    return AuditRecord(
        output_sheet="Single Cuts",
        metric_name="selection_rate",
        source_question_id="Q1",
        source_columns=("Q1",),
        filter_expr=None,
        numerator=4,
        denominator=10,
        formula="count(Q1 == 1) / count(Q1.notna())",
        value_raw=0.4,
        valid_n=10,
        missing_n=2,
        timestamp=UTC_NOW,
    )


def make_single_select_result(question_id: str = "Q1") -> SingleSelectResult:
    return SingleSelectResult(
        question_id=question_id,
        question_type=QuestionType.SINGLE_SELECT,
        valid_n=10,
        missing_n=0,
        denominator_policy=DenominatorPolicy.VALID_RESPONSES,
        distribution={1: {"label": "Yes", "count": 4, "rate": 0.4}},
        audit_records=(make_audit_record(),),
    )


def make_differentiator(index: int) -> DifferentiatorResult:
    return DifferentiatorResult(
        question_id=f"Q_DIFF_{index:02d}",
        question_text=f"Differentiator {index}",
        question_type=QuestionType.SINGLE_SELECT.value,
        cramers_v=1.0 - (index * 0.01),
        top_option_label="Option",
        top_option_winner_rate=0.6,
        top_option_loser_rate=0.3,
        top_option_lift=2.0,
        winner_n=50,
        loser_n=50,
        p_value=0.01,
    )


class TestModels(unittest.TestCase):
    def test_question_spec_valid_construction_supports_grid_questions(self) -> None:
        spec = QuestionSpec(
            question_id="[Q15]",
            canonical_id="Q15",
            question_text="Rate each area.",
            question_type=QuestionType.GRID_SINGLE_SELECT,
            raw_columns=("Q15r1", "Q15r2"),
            option_map={1: "Poor", 2: "Fair", 3: "Good", 4: "Excellent"},
            value_range=(1, 4),
            grid_row_labels={
                "Q15r1": "Overall company strategy",
                "Q15r2": "Leadership communication",
            },
        )

        self.assertIs(spec.question_type, QuestionType.GRID_SINGLE_SELECT)
        self.assertEqual(
            spec.grid_row_labels,
            {
                "Q15r1": "Overall company strategy",
                "Q15r2": "Leadership communication",
            },
        )

    def test_question_spec_invalid_construction_rejects_bad_value_range(self) -> None:
        with self.assertRaisesRegex(ValueError, "value_range"):
            QuestionSpec(
                question_id="[Q1]",
                canonical_id="Q1",
                question_text="Would you recommend the product?",
                question_type=QuestionType.SINGLE_SELECT,
                raw_columns=("Q1",),
                option_map={1: "Yes", 2: "No"},
                value_range=(5, 1),
            )

    def test_question_spec_demographic_flag_defaults_false(self) -> None:
        spec = make_question_spec("QDemo")

        self.assertFalse(spec.is_demographic)

    def test_question_spec_demographic_flag_can_be_true(self) -> None:
        spec = QuestionSpec(
            question_id="[QIndustry]",
            canonical_id="QIndustry",
            question_text="Which industry best describes your company?",
            question_type=QuestionType.SINGLE_SELECT,
            raw_columns=("QIndustry",),
            option_map={1: "Technology", 2: "Finance"},
            is_demographic=True,
        )

        self.assertTrue(spec.is_demographic)

    def test_survey_schema_valid_construction_and_lookup(self) -> None:
        eligible = make_question_spec("Q1")
        ineligible = QuestionSpec(
            question_id="[Q2]",
            canonical_id="Q2",
            question_text="Unclassified response.",
            question_type=QuestionType.UNKNOWN,
            raw_columns=(),
            option_map={},
            analysis_eligible=False,
            exclusion_reason="Question type could not be classified.",
        )

        schema = SurveySchema(
            questions=(eligible, ineligible),
            respondent_id_column="RespondentID",
            total_respondents=12,
            source_datamap_path="Datamap_sample.xlsx",
            source_rawdata_path="Rawdata_sample.xlsx",
            parsed_at=UTC_NOW,
        )

        self.assertEqual(schema.get_question("Q1"), eligible)
        self.assertIsNone(schema.get_question("missing"))
        self.assertEqual(schema.analysis_eligible_questions(), (eligible,))

    def test_survey_schema_invalid_construction_rejects_duplicate_ids(self) -> None:
        q1 = make_question_spec("Q1")
        q1_duplicate = make_question_spec("Q1")

        with self.assertRaisesRegex(ValueError, "unique"):
            SurveySchema(
                questions=(q1, q1_duplicate),
                respondent_id_column="RespondentID",
                total_respondents=12,
                source_datamap_path="Datamap_sample.xlsx",
                source_rawdata_path="Rawdata_sample.xlsx",
                parsed_at=UTC_NOW,
            )

    def test_audit_record_valid_construction(self) -> None:
        record = make_audit_record()

        self.assertEqual(record.value_raw, 0.4)
        self.assertIs(record.timestamp.tzinfo, timezone.utc)

    def test_audit_record_invalid_construction_rejects_empty_formula(self) -> None:
        with self.assertRaisesRegex(ValueError, "formula"):
            AuditRecord(
                output_sheet="Single Cuts",
                metric_name="selection_rate",
                source_question_id="Q1",
                source_columns=("Q1",),
                filter_expr=None,
                numerator=4,
                denominator=10,
                formula=" ",
                value_raw=0.4,
                valid_n=10,
                missing_n=2,
                timestamp=UTC_NOW,
            )

    def test_single_cut_result_valid_construction(self) -> None:
        result = SingleCutResult(
            question_id="Q1",
            question_type=QuestionType.SINGLE_SELECT,
            valid_n=10,
            missing_n=2,
            denominator_policy=DenominatorPolicy.VALID_RESPONSES,
            warnings=("Low base size",),
            audit_records=(make_audit_record(),),
        )

        self.assertEqual(result.valid_n, 10)
        self.assertEqual(result.audit_records, (make_audit_record(),))

    def test_single_cut_result_invalid_construction_rejects_negative_counts(self) -> None:
        with self.assertRaisesRegex(ValueError, "valid_n"):
            SingleCutResult(
                question_id="Q1",
                question_type=QuestionType.SINGLE_SELECT,
                valid_n=-1,
                missing_n=2,
                denominator_policy=DenominatorPolicy.VALID_RESPONSES,
            )

    def test_single_select_result_valid_construction(self) -> None:
        result = make_single_select_result()

        self.assertEqual(result.distribution[1]["rate"], 0.4)

    def test_single_select_result_invalid_construction_rejects_empty_distribution(
        self,
    ) -> None:
        with self.assertRaisesRegex(ValueError, "distribution"):
            SingleSelectResult(
                question_id="Q1",
                question_type=QuestionType.SINGLE_SELECT,
                valid_n=10,
                missing_n=0,
                denominator_policy=DenominatorPolicy.VALID_RESPONSES,
                distribution={},
            )

    def test_multi_select_result_valid_construction(self) -> None:
        result = MultiSelectResult(
            question_id="Q53",
            question_type=QuestionType.MULTI_SELECT_BINARY,
            valid_n=10,
            missing_n=1,
            denominator_policy=DenominatorPolicy.VALID_RESPONSES,
            selections={
                "Q53r1": {
                    "label": "Email",
                    "count": 6,
                    "selection_rate": 0.6,
                }
            },
            respondents_who_answered_any=10,
        )

        self.assertEqual(result.respondents_who_answered_any, 10)

    def test_multi_select_result_invalid_construction_rejects_negative_answered_any(
        self,
    ) -> None:
        with self.assertRaisesRegex(ValueError, "respondents_who_answered_any"):
            MultiSelectResult(
                question_id="Q53",
                question_type=QuestionType.MULTI_SELECT_BINARY,
                valid_n=10,
                missing_n=1,
                denominator_policy=DenominatorPolicy.VALID_RESPONSES,
                selections={
                    "Q53r1": {
                        "label": "Email",
                        "count": 6,
                        "selection_rate": 0.6,
                    }
                },
                respondents_who_answered_any=-1,
            )

    def test_numeric_result_valid_construction(self) -> None:
        result = NumericResult(
            question_id="Q10",
            question_type=QuestionType.DIRECT_NUMERIC,
            valid_n=10,
            missing_n=0,
            denominator_policy=DenominatorPolicy.VALID_RESPONSES,
            mean=7.2,
            median=7.0,
            std=1.1,
            min_val=5.0,
            max_val=10.0,
            percentiles={25: 6.0, 50: 7.0, 75: 8.0},
        )

        self.assertEqual(result.percentiles[50], 7.0)
        self.assertEqual(result.p25, 6.0)
        self.assertEqual(result.p50, 7.0)
        self.assertEqual(result.p75, 8.0)

    def test_numeric_result_invalid_construction_rejects_missing_percentiles(
        self,
    ) -> None:
        with self.assertRaisesRegex(ValueError, "percentiles"):
            NumericResult(
                question_id="Q10",
                question_type=QuestionType.DIRECT_NUMERIC,
                valid_n=10,
                missing_n=0,
                denominator_policy=DenominatorPolicy.VALID_RESPONSES,
                mean=7.2,
                median=7.0,
                std=1.1,
                min_val=5.0,
                max_val=10.0,
                percentiles={25: 6.0, 50: 7.0},
            )

    def test_grid_single_select_result_valid_construction(self) -> None:
        row_result = make_single_select_result("Q15r1")
        result = GridSingleSelectResult(
            question_id="Q15",
            question_type=QuestionType.GRID_SINGLE_SELECT,
            valid_n=10,
            missing_n=0,
            denominator_policy=DenominatorPolicy.VALID_RESPONSES,
            rows={"Q15r1": row_result},
        )

        self.assertEqual(result.rows["Q15r1"], row_result)

    def test_grid_single_select_result_invalid_construction_rejects_empty_rows(
        self,
    ) -> None:
        with self.assertRaisesRegex(ValueError, "rows"):
            GridSingleSelectResult(
                question_id="Q15",
                question_type=QuestionType.GRID_SINGLE_SELECT,
                valid_n=10,
                missing_n=0,
                denominator_policy=DenominatorPolicy.VALID_RESPONSES,
                rows={},
            )

    def test_top_differentiators_slicing(self) -> None:
        differentiators = tuple(make_differentiator(index) for index in range(25))
        result = OutcomeSegmentationResult(
            outcome_question_id="Q_OUT",
            segment_definition=SegmentDefinition(
                outcome_question_id="Q_OUT",
                segment_mode="categorical",
                winner_values=(1,),
            ),
            winner_n=50,
            loser_n=50,
            total_n=100,
            differentiators=differentiators,
            winner_profile=WinnerProfile(
                outcome_question_id="Q_OUT",
                winner_label="Winner",
                winner_n=50,
                loser_n=50,
                defining_traits=(),
            ),
            skipped_questions=(),
        )

        self.assertEqual(len(result.top_differentiators(10)), 10)
        self.assertEqual(len(result.top_differentiators(50)), 25)
        self.assertEqual(result.top_differentiators(0), ())
        self.assertEqual(result.max_available_differentiators, 25)


if __name__ == "__main__":
    unittest.main()
