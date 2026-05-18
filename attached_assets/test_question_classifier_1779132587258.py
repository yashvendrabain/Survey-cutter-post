"""Tests for question classification into SurveySchema."""

from __future__ import annotations

import unittest

from src.models import QuestionType
from src.question_classifier import classify_questions


CLASSIFIER_DATA_MAP = {
    "questions": [
        {
            "canonical_id": "record",
            "raw_id": "record",
            "question_text": "Respondent ID",
            "type_hint": None,
            "value_range": None,
            "options": [],
            "sub_columns": [],
            "parent_canonical_id": None,
            "source_row": 1,
            "warnings": [],
        },
        {
            "canonical_id": "Q3",
            "raw_id": "[Q3]",
            "question_text": "Are you currently in a full-time position",
            "type_hint": "values_range",
            "value_range": (1, 2),
            "options": [(1, "Yes"), (2, "No")],
            "sub_columns": [],
            "parent_canonical_id": None,
            "source_row": 2,
            "warnings": [],
        },
        {
            "canonical_id": "Q53",
            "raw_id": "Q53",
            "question_text": "Which challenges apply?",
            "type_hint": "values_range",
            "value_range": (0, 1),
            "options": [],
            "sub_columns": [("Q53r1", "Knowledge gap"), ("Q53r2", "Alignment gap")],
            "parent_canonical_id": None,
            "source_row": 10,
            "warnings": [],
        },
        {
            "canonical_id": "Q33",
            "raw_id": "Q33",
            "question_text": "Allocate 100 points",
            "type_hint": "values_range",
            "value_range": (0, 999),
            "options": [],
            "sub_columns": [("Q33r1", "Pricing"), ("Q33r2", "Customer shift")],
            "parent_canonical_id": None,
            "source_row": 15,
            "warnings": [],
        },
        {
            "canonical_id": "Q15",
            "raw_id": "Q15",
            "question_text": "Rate involvement",
            "type_hint": "values_range",
            "value_range": (1, 4),
            "options": [
                (1, "Directly involved in decision making AND budget"),
                (2, "Directly involved in decision making OR budget"),
                (3, "Indirectly involved"),
                (4, "Not involved"),
            ],
            "sub_columns": [
                ("Q15r1", "Overall company strategy"),
                ("Q15r2", "Go-to-market strategy"),
                ("Q15r3", "Partner strategy"),
            ],
            "parent_canonical_id": None,
            "source_row": 20,
            "warnings": [],
        },
        {
            "canonical_id": "Q4r98oe",
            "raw_id": "[Q4r98oe]",
            "question_text": "Other industry",
            "type_hint": "open_text",
            "value_range": None,
            "options": [],
            "sub_columns": [],
            "parent_canonical_id": "Q4",
            "source_row": 30,
            "warnings": [],
        },
        {
            "canonical_id": "vQTIME_MINUTES",
            "raw_id": "[vQTIME_MINUTES]",
            "question_text": "Survey length in minutes",
            "type_hint": "values_range",
            "value_range": (-99999999999999, 999999999999999),
            "options": [],
            "sub_columns": [],
            "parent_canonical_id": None,
            "source_row": 35,
            "warnings": [],
        },
        {
            "canonical_id": "QUnknown",
            "raw_id": "QUnknown",
            "question_text": "Missing type hint",
            "type_hint": None,
            "value_range": None,
            "options": [],
            "sub_columns": [],
            "parent_canonical_id": None,
            "source_row": 40,
            "warnings": ["header followed by blank row, no type hint"],
        },
        {
            "canonical_id": "QMissingRaw",
            "raw_id": "QMissingRaw",
            "question_text": "Single select missing from raw",
            "type_hint": "values_range",
            "value_range": (1, 2),
            "options": [(1, "Yes"), (2, "No")],
            "sub_columns": [],
            "parent_canonical_id": None,
            "source_row": 45,
            "warnings": [],
        },
        {
            "canonical_id": "QOther",
            "raw_id": "QOther",
            "question_text": "Pick an option",
            "type_hint": "values_range",
            "value_range": (1, 98),
            "options": [(1, "No"), (98, "Other (please specify)")],
            "sub_columns": [],
            "parent_canonical_id": None,
            "source_row": 50,
            "warnings": [],
        },
        {
            "canonical_id": "QIndustry",
            "raw_id": "QIndustry",
            "question_text": "Which industry best describes your company?",
            "type_hint": "values_range",
            "value_range": (1, 3),
            "options": [(1, "Technology"), (2, "Finance"), (3, "Healthcare")],
            "sub_columns": [],
            "parent_canonical_id": None,
            "source_row": 55,
            "warnings": [],
        },
        {
            "canonical_id": "QRegion",
            "raw_id": "QRegion",
            "question_text": "Which region are you based in?",
            "type_hint": "values_range",
            "value_range": (1, 3),
            "options": [(1, "APAC"), (2, "EMEA"), (3, "Americas")],
            "sub_columns": [],
            "parent_canonical_id": None,
            "source_row": 60,
            "warnings": [],
        },
        {
            "canonical_id": "QUnrelated",
            "raw_id": "QUnrelated",
            "question_text": "How satisfied are you?",
            "type_hint": "values_range",
            "value_range": (1, 2),
            "options": [(1, "Satisfied"), (2, "Not satisfied")],
            "sub_columns": [],
            "parent_canonical_id": None,
            "source_row": 65,
            "warnings": [],
        },
        {
            "canonical_id": "QTooManyIndustry",
            "raw_id": "QTooManyIndustry",
            "question_text": "Which industry sector is most relevant?",
            "type_hint": "values_range",
            "value_range": (1, 13),
            "options": [
                (1, "One"),
                (2, "Two"),
                (3, "Three"),
                (4, "Four"),
                (5, "Five"),
                (6, "Six"),
                (7, "Seven"),
                (8, "Eight"),
                (9, "Nine"),
                (10, "Ten"),
                (11, "Eleven"),
                (12, "Twelve"),
                (13, "Thirteen"),
            ],
            "sub_columns": [],
            "parent_canonical_id": None,
            "source_row": 70,
            "warnings": [],
        },
    ],
    "source_path": "classifier_datamap.xlsx",
    "sheet_name": "Sheet1",
    "total_rows_in_sheet": 50,
    "parser_warnings": [],
}

RAW_COLUMNS = [
    "record",
    "Q3",
    "Q53r1",
    "Q53r2",
    "Q33r1",
    "Q33r2",
    "Q15r1",
    "Q15r2",
    "Q4r98oe",
    "vQTIME_MINUTES",
    "QOther",
    "QIndustry",
    "QRegion",
    "QUnrelated",
    "QTooManyIndustry",
]


def classify_test_schema():
    return classify_questions(
        CLASSIFIER_DATA_MAP,
        RAW_COLUMNS,
        respondent_id_column="record",
        total_respondents=20,
        source_rawdata_path="raw.csv",
    )


def sibling_grid_question(
    canonical_id: str,
    criterion_label: str,
    root_text: str,
    options: list[tuple[int, str]] | None = None,
    value_range: tuple[int, int] = (1, 10),
) -> dict:
    return {
        "canonical_id": canonical_id,
        "raw_id": canonical_id,
        "question_text": f"{criterion_label} - {root_text}",
        "type_hint": "values_range",
        "value_range": value_range,
        "options": options
        or [(index, str(index)) for index in range(value_range[0], value_range[1] + 1)],
        "sub_columns": [],
        "parent_canonical_id": None,
        "source_row": 90,
        "warnings": [],
    }


def classify_with_extra_questions(extra_questions: list[dict], raw_columns: list[str]):
    data_map = {
        **CLASSIFIER_DATA_MAP,
        "questions": [*CLASSIFIER_DATA_MAP["questions"], *extra_questions],
    }
    return classify_questions(
        data_map,
        [*RAW_COLUMNS, *raw_columns],
        respondent_id_column="record",
        total_respondents=20,
        source_rawdata_path="raw.csv",
    )


class TestQuestionClassifier(unittest.TestCase):
    def test_single_select_classification(self) -> None:
        schema = classify_test_schema()
        question = schema.get_question("Q3")

        self.assertIsNotNone(question)
        self.assertIs(question.question_type, QuestionType.SINGLE_SELECT)
        self.assertEqual(question.raw_columns, ("Q3",))
        self.assertEqual(question.option_map, {1: "Yes", 2: "No"})

    def test_multi_select_binary_classification(self) -> None:
        schema = classify_test_schema()
        question = schema.get_question("Q53")

        self.assertIsNotNone(question)
        self.assertIs(question.question_type, QuestionType.MULTI_SELECT_BINARY)
        self.assertEqual(question.raw_columns, ("Q53r1", "Q53r2"))
        self.assertEqual(
            question.option_map,
            {"Q53r1": "Knowledge gap", "Q53r2": "Alignment gap"},
        )

    def test_numeric_allocation_classification(self) -> None:
        schema = classify_test_schema()
        question = schema.get_question("Q33")

        self.assertIsNotNone(question)
        self.assertIs(question.question_type, QuestionType.NUMERIC_ALLOCATION)
        self.assertEqual(question.raw_columns, ("Q33r1", "Q33r2"))

    def test_grid_single_select_classification(self) -> None:
        schema = classify_test_schema()
        question = schema.get_question("Q15")

        self.assertIsNotNone(question)
        self.assertIs(question.question_type, QuestionType.GRID_SINGLE_SELECT)
        self.assertEqual(question.possible_role, "GRID_CATEGORICAL")
        self.assertEqual(
            question.option_map,
            {
                1: "Directly involved in decision making AND budget",
                2: "Directly involved in decision making OR budget",
                3: "Indirectly involved",
                4: "Not involved",
            },
        )

    def test_open_text_classification(self) -> None:
        schema = classify_test_schema()
        question = schema.get_question("Q4r98oe")

        self.assertIsNotNone(question)
        self.assertIs(question.question_type, QuestionType.OPEN_TEXT)
        self.assertEqual(question.parent_question_id, "Q4")

    def test_metadata_columns_classified_correctly(self) -> None:
        schema = classify_test_schema()
        record = schema.get_question("record")
        timing = schema.get_question("vQTIME_MINUTES")

        self.assertIsNotNone(record)
        self.assertIsNotNone(timing)
        self.assertIs(record.question_type, QuestionType.METADATA_OR_ID)
        self.assertIs(timing.question_type, QuestionType.METADATA_OR_ID)

    def test_unknown_when_type_hint_is_none(self) -> None:
        schema = classify_test_schema()
        question = schema.get_question("QUnknown")

        self.assertIsNotNone(question)
        self.assertIs(question.question_type, QuestionType.UNKNOWN)

    def test_raw_column_not_found_sets_ineligible(self) -> None:
        schema = classify_test_schema()
        question = schema.get_question("QMissingRaw")

        self.assertIsNotNone(question)
        self.assertFalse(question.analysis_eligible)
        self.assertEqual(question.exclusion_reason, "raw column not found in data")

    def test_grid_row_labels_match_raw_columns(self) -> None:
        schema = classify_test_schema()
        question = schema.get_question("Q15")

        self.assertIsNotNone(question)
        self.assertEqual(question.raw_columns, ("Q15r1", "Q15r2"))
        self.assertEqual(
            question.grid_row_labels,
            {
                "Q15r1": "Overall company strategy",
                "Q15r2": "Go-to-market strategy",
            },
        )

    def test_grid_rated_subtype_detected_from_numeric_scale_labels(self) -> None:
        data_map = {
            **CLASSIFIER_DATA_MAP,
            "questions": [
                *CLASSIFIER_DATA_MAP["questions"],
                {
                    "canonical_id": "QRated",
                    "raw_id": "QRated",
                    "question_text": "Rate each attribute",
                    "type_hint": "values_range",
                    "value_range": (1, 10),
                    "options": [(index, str(index)) for index in range(1, 11)],
                    "sub_columns": [
                        ("QRatedr1", "Speed"),
                        ("QRatedr2", "Fit"),
                    ],
                    "parent_canonical_id": None,
                    "source_row": 80,
                    "warnings": [],
                },
            ],
        }
        schema = classify_questions(
            data_map,
            [*RAW_COLUMNS, "QRatedr1", "QRatedr2"],
            respondent_id_column="record",
            total_respondents=20,
            source_rawdata_path="raw.csv",
        )
        question = schema.get_question("QRated")

        self.assertIsNotNone(question)
        self.assertIs(question.question_type, QuestionType.GRID_SINGLE_SELECT)
        self.assertEqual(question.possible_role, "GRID_RATED")

    def test_grid_c_columns_expand_from_raw_data_headers(self) -> None:
        data_map = {
            **CLASSIFIER_DATA_MAP,
            "questions": [
                *CLASSIFIER_DATA_MAP["questions"],
                {
                    "canonical_id": "QVendorGrid",
                    "raw_id": "QVendorGrid",
                    "question_text": "Rate each vendor attribute",
                    "type_hint": "values_range",
                    "value_range": (1, 10),
                    "options": [(index, str(index)) for index in range(1, 11)],
                    "sub_columns": [
                        ("QVendorGridr1", "Implementation speed"),
                        ("QVendorGridr2", "Integration fit"),
                        ("QVendorGridr3", "Efficiency gains"),
                    ],
                    "parent_canonical_id": None,
                    "source_row": 85,
                    "warnings": [],
                },
            ],
        }
        raw_columns = [
            *RAW_COLUMNS,
            "QVendorGridr1c1",
            "QVendorGridr1c2",
            "QVendorGridr2c1",
            "QVendorGridr2c2",
            "QVendorGridr3c1",
            "QVendorGridr3c2",
        ]
        schema = classify_questions(
            data_map,
            raw_columns,
            respondent_id_column="record",
            total_respondents=20,
            source_rawdata_path="raw.csv",
        )
        question = schema.get_question("QVendorGrid")

        self.assertIsNotNone(question)
        self.assertEqual(
            question.raw_columns,
            (
                "QVendorGridr1c1",
                "QVendorGridr1c2",
                "QVendorGridr2c1",
                "QVendorGridr2c2",
                "QVendorGridr3c1",
                "QVendorGridr3c2",
            ),
        )
        self.assertEqual(question.grid_row_labels["QVendorGridr1c1"], "Implementation speed")
        self.assertEqual(question.grid_row_labels["QVendorGridr1c2"], "Implementation speed")

    def test_grid_sibling_questions_merge_into_parent_spec(self) -> None:
        root_text = "Rate each vendor attribute from 1 to 10"
        schema = classify_with_extra_questions(
            [
                sibling_grid_question("Q1r1", "Implementation speed", root_text),
                sibling_grid_question("Q1r2", "Integration fit", root_text),
                sibling_grid_question("Q1r3", "Efficiency gains", root_text),
            ],
            [
                "Q1r1c1",
                "Q1r1c2",
                "Q1r2c1",
                "Q1r2c2",
                "Q1r3c1",
                "Q1r3c2",
            ],
        )
        question = schema.get_question("Q1")

        self.assertIsNotNone(question)
        self.assertIsNone(schema.get_question("Q1r1"))
        self.assertIs(question.question_type, QuestionType.GRID_SINGLE_SELECT)
        self.assertEqual(question.possible_role, "GRID_RATED")
        self.assertEqual(
            question.raw_columns,
            ("Q1r1c1", "Q1r1c2", "Q1r2c1", "Q1r2c2", "Q1r3c1", "Q1r3c2"),
        )
        self.assertEqual(question.question_text, root_text)
        self.assertEqual(question.grid_row_labels["Q1r1c1"], "Implementation speed")
        self.assertEqual(question.grid_row_labels["Q1r1c2"], "Implementation speed")
        self.assertEqual(
            set(question.grid_row_labels.values()),
            {"Implementation speed", "Integration fit", "Efficiency gains"},
        )

    def test_grid_sibling_merge_keeps_families_separate(self) -> None:
        q1_root = "Rate each vendor attribute from 1 to 10"
        q2_root = "What role did each stakeholder play"
        role_options = [
            (1, "Decision maker"),
            (2, "Influencer"),
            (3, "Not involved"),
        ]
        schema = classify_with_extra_questions(
            [
                sibling_grid_question("Q1r1", "Implementation speed", q1_root),
                sibling_grid_question("Q1r2", "Integration fit", q1_root),
                sibling_grid_question("Q1r3", "Efficiency gains", q1_root),
                sibling_grid_question("Q2r1", "IT / Technical", q2_root, role_options, (1, 3)),
                sibling_grid_question("Q2r2", "Security", q2_root, role_options, (1, 3)),
                sibling_grid_question("Q2r3", "Finance", q2_root, role_options, (1, 3)),
                sibling_grid_question("Q2r4", "Operations", q2_root, role_options, (1, 3)),
            ],
            ["Q1r1", "Q1r2", "Q1r3", "Q2r1", "Q2r2", "Q2r3", "Q2r4"],
        )

        self.assertIsNotNone(schema.get_question("Q1"))
        self.assertIsNotNone(schema.get_question("Q2"))
        self.assertIsNone(schema.get_question("Q1r2"))
        self.assertIsNone(schema.get_question("Q2r4"))
        self.assertEqual(schema.get_question("Q1").possible_role, "GRID_RATED")
        self.assertEqual(schema.get_question("Q2").possible_role, "GRID_CATEGORICAL")

    def test_grid_sibling_questions_do_not_merge_when_root_or_options_differ(self) -> None:
        role_options = [(1, "A"), (2, "B")]
        schema = classify_with_extra_questions(
            [
                sibling_grid_question("Q1r1", "Speed", "Rate vendor attributes"),
                sibling_grid_question("Q1r2", "Fit", "Different root text"),
                sibling_grid_question("Q2r1", "Security", "Stakeholder role", role_options, (1, 2)),
                sibling_grid_question(
                    "Q2r2",
                    "Finance",
                    "Stakeholder role",
                    [(1, "A"), (2, "Different")],
                    (1, 2),
                ),
            ],
            ["Q1r1", "Q1r2", "Q2r1", "Q2r2"],
        )

        self.assertIsNone(schema.get_question("Q1"))
        self.assertIsNone(schema.get_question("Q2"))
        self.assertIsNotNone(schema.get_question("Q1r1"))
        self.assertIsNotNone(schema.get_question("Q2r2"))

    def test_grid_binary_select_siblings_are_not_merged(self) -> None:
        options = [(0, "NO TO: Selected"), (1, "Selected")]
        schema = classify_with_extra_questions(
            [
                sibling_grid_question(
                    "Q38r1",
                    "Discount offered",
                    "What would have changed your decision",
                    options,
                    (0, 1),
                ),
                sibling_grid_question(
                    "Q38r2",
                    "Better references",
                    "What would have changed your decision",
                    options,
                    (0, 1),
                ),
            ],
            ["Q38r1", "Q38r2"],
        )

        self.assertIsNone(schema.get_question("Q38"))
        self.assertIsNotNone(schema.get_question("Q38r1"))
        self.assertIsNotNone(schema.get_question("Q38r2"))

    def test_grid_binary_subtype_detected_from_selection_labels(self) -> None:
        data_map = {
            **CLASSIFIER_DATA_MAP,
            "questions": [
                *CLASSIFIER_DATA_MAP["questions"],
                {
                    "canonical_id": "QBinaryGrid",
                    "raw_id": "QBinaryGrid",
                    "question_text": "Select any changes",
                    "type_hint": "values_range",
                    "value_range": (0, 1),
                    "options": [(0, "Unchecked"), (1, "Selected")],
                    "sub_columns": [
                        ("QBinaryGridr1", "Discount"),
                        ("QBinaryGridr2", "Support"),
                    ],
                    "parent_canonical_id": None,
                    "source_row": 90,
                    "warnings": [],
                },
            ],
        }
        schema = classify_questions(
            data_map,
            [*RAW_COLUMNS, "QBinaryGridr1", "QBinaryGridr2"],
            respondent_id_column="record",
            total_respondents=20,
            source_rawdata_path="raw.csv",
        )
        question = schema.get_question("QBinaryGrid")

        self.assertIsNotNone(question)
        self.assertIs(question.question_type, QuestionType.GRID_SINGLE_SELECT)
        self.assertEqual(question.possible_role, "GRID_BINARY_SELECT")

    def test_option_other_code_detected(self) -> None:
        schema = classify_test_schema()
        question = schema.get_question("QOther")

        self.assertIsNotNone(question)
        self.assertEqual(question.option_other_code, 98)

    def test_survey_schema_canonical_ids_unique(self) -> None:
        schema = classify_test_schema()
        canonical_ids = [question.canonical_id for question in schema.questions]

        self.assertEqual(len(canonical_ids), len(set(canonical_ids)))
        self.assertEqual(schema.respondent_id_column, "record")
        self.assertEqual(schema.total_respondents, 20)

    def test_demographic_question_detected_industry(self) -> None:
        schema = classify_test_schema()
        question = schema.get_question("QIndustry")

        self.assertIsNotNone(question)
        self.assertTrue(question.is_demographic)

    def test_demographic_question_detected_region(self) -> None:
        schema = classify_test_schema()
        question = schema.get_question("QRegion")

        self.assertIsNotNone(question)
        self.assertTrue(question.is_demographic)

    def test_demographic_question_not_detected_for_unrelated_question(self) -> None:
        schema = classify_test_schema()
        question = schema.get_question("QUnrelated")

        self.assertIsNotNone(question)
        self.assertFalse(question.is_demographic)

    def test_demographic_question_not_detected_for_too_many_options(self) -> None:
        schema = classify_test_schema()
        question = schema.get_question("QTooManyIndustry")

        self.assertIsNotNone(question)
        self.assertFalse(question.is_demographic)

    def test_schema_demographic_questions_helper(self) -> None:
        schema = classify_test_schema()
        demographic_ids = tuple(
            question.canonical_id for question in schema.demographic_questions()
        )

        self.assertEqual(demographic_ids, ("QIndustry", "QRegion"))

    def test_pipe_reference_resolves_to_readable_prior_selection(self) -> None:
        data_map = {
            **CLASSIFIER_DATA_MAP,
            "questions": [
                *CLASSIFIER_DATA_MAP["questions"],
                {
                    "canonical_id": "H_Q30_Col2_Selection",
                    "raw_id": "H_Q30_Col2_Selection",
                    "question_text": "Vendor evaluated during procurement process",
                    "type_hint": "values_range",
                    "value_range": (1, 2),
                    "options": [(1, "A"), (2, "B")],
                    "sub_columns": [],
                    "parent_canonical_id": None,
                    "source_row": 75,
                    "warnings": [],
                },
                {
                    "canonical_id": "Q38",
                    "raw_id": "Q38",
                    "question_text": (
                        "Thinking about [pipe: H_Q30_Col2_Selection] "
                        "Changes to decision for vendor"
                    ),
                    "type_hint": "values_range",
                    "value_range": (1, 2),
                    "options": [(1, "Yes"), (2, "No")],
                    "sub_columns": [],
                    "parent_canonical_id": None,
                    "source_row": 80,
                    "warnings": [],
                },
            ],
        }
        schema = classify_questions(
            data_map,
            [*RAW_COLUMNS, "H_Q30_Col2_Selection", "Q38"],
            respondent_id_column="record",
            total_respondents=20,
            source_rawdata_path="raw.csv",
        )

        question = schema.get_question("Q38")
        self.assertIsNotNone(question)
        self.assertIn(
            "prior selection (vendor evaluated during procurement process)",
            question.question_text,
        )
        self.assertEqual(question.conditional_on, "H_Q30_Col2_Selection")

    def test_pn_reference_resolves_to_readable_prior_selection(self) -> None:
        data_map = {
            **CLASSIFIER_DATA_MAP,
            "questions": [
                *CLASSIFIER_DATA_MAP["questions"],
                {
                    "canonical_id": "Q9",
                    "raw_id": "Q9",
                    "question_text": "Software category selected for evaluation",
                    "type_hint": "values_range",
                    "value_range": (1, 2),
                    "options": [(1, "A"), (2, "B")],
                    "sub_columns": [],
                    "parent_canonical_id": None,
                    "source_row": 75,
                    "warnings": [],
                },
                {
                    "canonical_id": "Q10",
                    "raw_id": "Q10",
                    "question_text": "Follow-up for [Pn: Q9]",
                    "type_hint": "values_range",
                    "value_range": (1, 2),
                    "options": [(1, "Yes"), (2, "No")],
                    "sub_columns": [],
                    "parent_canonical_id": None,
                    "source_row": 80,
                    "warnings": [],
                },
            ],
        }
        schema = classify_questions(
            data_map,
            [*RAW_COLUMNS, "Q9", "Q10"],
            respondent_id_column="record",
            total_respondents=20,
            source_rawdata_path="raw.csv",
        )

        question = schema.get_question("Q10")
        self.assertIsNotNone(question)
        self.assertEqual(
            question.question_text,
            "Follow-up for prior selection (software category selected for evaluation)",
        )

    def test_unresolvable_pipe_reference_falls_back_to_prior_selection(self) -> None:
        data_map = {
            **CLASSIFIER_DATA_MAP,
            "questions": [
                *CLASSIFIER_DATA_MAP["questions"],
                {
                    "canonical_id": "Q404",
                    "raw_id": "Q404",
                    "question_text": "Follow-up for [pipe: MissingQuestion]",
                    "type_hint": "values_range",
                    "value_range": (1, 2),
                    "options": [(1, "Yes"), (2, "No")],
                    "sub_columns": [],
                    "parent_canonical_id": None,
                    "source_row": 80,
                    "warnings": [],
                },
            ],
        }
        schema = classify_questions(
            data_map,
            [*RAW_COLUMNS, "Q404"],
            respondent_id_column="record",
            total_respondents=20,
            source_rawdata_path="raw.csv",
        )

        question = schema.get_question("Q404")
        self.assertIsNotNone(question)
        self.assertEqual(question.question_text, "Follow-up for prior selection")


if __name__ == "__main__":
    unittest.main()
