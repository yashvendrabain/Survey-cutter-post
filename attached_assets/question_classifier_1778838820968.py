"""Question classifier for parsed data maps."""

from __future__ import annotations

from dataclasses import replace
from datetime import datetime, timezone
import re

from src.datamap_parser import DataMap, ParsedQuestion
from src.models import (
    DenominatorPolicy,
    QuestionSpec,
    QuestionType,
    SurveySchema,
)


RESPONDENT_ID_CANDIDATES = (
    "record",
    "uuid",
    "respondent_id",
    "id",
    "ID",
    "RespondentID",
)
METADATA_CANONICAL_IDS = {
    "record",
    "uuid",
    "date",
    "markers",
    "status",
    "vend",
    "hQMODE",
    "noanswer",
    "vmobileos",
}
VERY_WIDE_RANGE_THRESHOLD = 1_000_000
DEMOGRAPHIC_KEYWORDS = (
    "industry",
    "sector",
    "region",
    "country",
    "geography",
    "size",
    "employees",
    "headcount",
    "revenue range",
    "function",
    "department",
    "role",
    "seniority",
    "tier",
    "company",
    "organization",
    "organisation",
    "vertical",
    "segment",
    "market",
)
_PIPE_PATTERN = re.compile(r"\[(pipe|pn):\s*([^\]]+)\]", re.IGNORECASE)
_TRAILING_PUNCTUATION = ".,;:!?- "


def classify_questions(
    data_map: DataMap,
    raw_columns: list[str],
    respondent_id_column: str | None = None,
    total_respondents: int = 1,
    source_rawdata_path: str = "<unknown raw data path>",
) -> SurveySchema:
    """Classify parsed questions into the SurveySchema contract."""

    raw_column_set = set(raw_columns)
    all_sub_columns = {
        sub_column_id
        for question in data_map["questions"]
        for sub_column_id, _ in question["sub_columns"]
    }

    question_text_lookup = _question_text_lookup(data_map)
    resolved_texts = {
        question["canonical_id"]: _resolve_pipes(
            question["question_text"],
            question_text_lookup,
        )
        for question in data_map["questions"]
    }
    conditional_refs = {
        question["canonical_id"]: _first_pipe_reference(question["question_text"])
        for question in data_map["questions"]
    }

    question_specs = tuple(
        _build_question_spec(
            question,
            raw_column_set,
            all_sub_columns,
            resolved_texts.get(question["canonical_id"], question["question_text"]),
            conditional_refs.get(question["canonical_id"]),
        )
        for question in data_map["questions"]
    )

    return SurveySchema(
        questions=question_specs,
        respondent_id_column=respondent_id_column
        or _identify_respondent_id_column(raw_columns),
        total_respondents=total_respondents,
        source_datamap_path=data_map["source_path"],
        source_rawdata_path=source_rawdata_path,
        parsed_at=datetime.now(timezone.utc),
    )


def _build_question_spec(
    question: ParsedQuestion,
    raw_column_set: set[str],
    all_sub_columns: set[str],
    resolved_question_text: str,
    conditional_on: str | None,
) -> QuestionSpec:
    question_type = _classify_question(question, all_sub_columns)
    expected_columns = _expected_columns(question)
    present_columns = tuple(column for column in expected_columns if column in raw_column_set)
    analysis_eligible = True
    exclusion_reason: str | None = None

    if question_type in {
        QuestionType.SINGLE_SELECT,
        QuestionType.DIRECT_NUMERIC,
        QuestionType.OPEN_TEXT,
    }:
        if question["canonical_id"] not in raw_column_set:
            analysis_eligible = False
            exclusion_reason = "raw column not found in data"

    if question_type in {
        QuestionType.MULTI_SELECT_BINARY,
        QuestionType.NUMERIC_ALLOCATION,
        QuestionType.GRID_SINGLE_SELECT,
    }:
        if not present_columns:
            analysis_eligible = False
            exclusion_reason = "raw column not found in data"

    raw_columns = _raw_columns_for_spec(
        question_type, question["canonical_id"], expected_columns, present_columns
    )
    option_map = _option_map_for_spec(question_type, question)
    grid_row_labels = _grid_row_labels_for_spec(question_type, question, raw_columns)

    if (
        question_type is QuestionType.GRID_SINGLE_SELECT
        and 0 < len(present_columns) < len(expected_columns)
    ):
        missing = tuple(column for column in expected_columns if column not in raw_column_set)
        exclusion_reason = "missing grid raw columns: " + ", ".join(missing)

    spec = QuestionSpec(
        question_id=question["raw_id"],
        canonical_id=question["canonical_id"],
        question_text=resolved_question_text,
        question_type=question_type,
        raw_columns=raw_columns,
        option_map=option_map,
        value_range=question["value_range"],
        denominator_policy=DenominatorPolicy.VALID_RESPONSES,
        analysis_eligible=analysis_eligible,
        exclusion_reason=exclusion_reason,
        parent_question_id=question["parent_canonical_id"],
        grid_row_labels=grid_row_labels,
        option_other_code=_option_other_code(question),
        conditional_on=conditional_on,
    )
    return replace(
        spec,
        is_demographic=_is_demographic_question(spec, raw_column_set),
    )


def _classify_question(
    question: ParsedQuestion,
    all_sub_columns: set[str],
) -> QuestionType:
    canonical_id = question["canonical_id"]
    value_range = question["value_range"]

    if _is_metadata(canonical_id, value_range, all_sub_columns):
        return QuestionType.METADATA_OR_ID
    if _is_uncertain_v_metadata_candidate(canonical_id, all_sub_columns):
        return QuestionType.UNKNOWN

    type_hint = question["type_hint"]
    has_options = bool(question["options"])
    has_sub_columns = bool(question["sub_columns"])

    if type_hint == "open_text":
        return QuestionType.OPEN_TEXT
    if type_hint == "open_numeric":
        return QuestionType.DIRECT_NUMERIC
    if type_hint is None:
        return QuestionType.UNKNOWN

    if type_hint == "values_range":
        if has_sub_columns and has_options:
            return QuestionType.GRID_SINGLE_SELECT
        if has_sub_columns and not has_options:
            return _classify_sub_column_numeric_group(value_range)
        if has_options and not has_sub_columns:
            return QuestionType.SINGLE_SELECT
        return QuestionType.DIRECT_NUMERIC

    return QuestionType.UNKNOWN


def _question_text_lookup(data_map: DataMap) -> dict[str, str]:
    lookup: dict[str, str] = {}
    for question in data_map["questions"]:
        text = question["question_text"]
        lookup[question["canonical_id"]] = text
        lookup[question["raw_id"].strip("[]")] = text
    return lookup


def _resolve_pipes(question_text: str, schema_questions: dict[str, str]) -> str:
    def replacement(match: re.Match[str]) -> str:
        reference = match.group(2).strip()
        referenced_text = schema_questions.get(reference)
        if not referenced_text:
            return "prior selection"
        return f"prior selection ({_short_question_label(referenced_text)})"

    return _PIPE_PATTERN.sub(replacement, question_text)


def _first_pipe_reference(question_text: str) -> str | None:
    match = _PIPE_PATTERN.search(question_text)
    if match is None:
        return None
    return match.group(2).strip()


def _short_question_label(question_text: str, max_words: int = 5) -> str:
    text = re.sub(r"^\s*Q\d+[A-Za-z]*\s*[-:]?\s*", "", question_text).strip()
    words = re.findall(r"[A-Za-z0-9&%]+", text.lower())
    if not words:
        return "prior question"
    return " ".join(words[:max_words]).rstrip(_TRAILING_PUNCTUATION)


def _is_metadata(
    canonical_id: str,
    value_range: tuple[int, int] | None,
    all_sub_columns: set[str],
) -> bool:
    if canonical_id in METADATA_CANONICAL_IDS:
        return True
    if (
        canonical_id.startswith("v")
        and value_range is not None
        and abs(value_range[1] - value_range[0]) > VERY_WIDE_RANGE_THRESHOLD
    ):
        return True
    return False


def _is_uncertain_v_metadata_candidate(
    canonical_id: str,
    all_sub_columns: set[str],
) -> bool:
    return canonical_id.startswith("v") and canonical_id not in all_sub_columns


def _classify_sub_column_numeric_group(
    value_range: tuple[int, int] | None,
) -> QuestionType:
    if value_range is None:
        return QuestionType.UNKNOWN

    low, high = value_range
    if low == 0 and high == 1:
        return QuestionType.MULTI_SELECT_BINARY
    if low == 0 and 1 < high <= 10:
        return QuestionType.MULTI_SELECT_BINARY
    if low == 0 and high == 999:
        return QuestionType.NUMERIC_ALLOCATION
    if high > 10:
        return QuestionType.NUMERIC_ALLOCATION
    return QuestionType.UNKNOWN


def _expected_columns(question: ParsedQuestion) -> tuple[str, ...]:
    if question["sub_columns"]:
        return tuple(sub_column_id for sub_column_id, _ in question["sub_columns"])
    return (question["canonical_id"],)


def _raw_columns_for_spec(
    question_type: QuestionType,
    canonical_id: str,
    expected_columns: tuple[str, ...],
    present_columns: tuple[str, ...],
) -> tuple[str, ...]:
    if question_type in {
        QuestionType.MULTI_SELECT_BINARY,
        QuestionType.NUMERIC_ALLOCATION,
        QuestionType.GRID_SINGLE_SELECT,
    }:
        return present_columns or expected_columns

    if question_type in {
        QuestionType.SINGLE_SELECT,
        QuestionType.DIRECT_NUMERIC,
        QuestionType.OPEN_TEXT,
    }:
        return (canonical_id,)

    if canonical_id in present_columns:
        return (canonical_id,)
    return ()


def _option_map_for_spec(
    question_type: QuestionType,
    question: ParsedQuestion,
) -> dict[int | str, str]:
    if question_type is QuestionType.SINGLE_SELECT:
        return {code: label for code, label in question["options"]}
    if question_type is QuestionType.MULTI_SELECT_BINARY:
        return {sub_column_id: label for sub_column_id, label in question["sub_columns"]}
    if question_type is QuestionType.GRID_SINGLE_SELECT:
        return {code: label for code, label in question["options"]}
    return {}


def _grid_row_labels_for_spec(
    question_type: QuestionType,
    question: ParsedQuestion,
    raw_columns: tuple[str, ...],
) -> dict[str, str] | None:
    if question_type is not QuestionType.GRID_SINGLE_SELECT:
        return None

    row_label_lookup = {
        sub_column_id: label for sub_column_id, label in question["sub_columns"]
    }
    return {
        sub_column_id: row_label_lookup[sub_column_id]
        for sub_column_id in raw_columns
        if sub_column_id in row_label_lookup
    }


def _option_other_code(question: ParsedQuestion) -> int | str | None:
    for code, label in question["options"]:
        label_lower = label.lower()
        if "other" in label_lower or "specify" in label_lower:
            return code
    return None


def _is_demographic_question(
    question: QuestionSpec,
    raw_columns: set[str],
) -> bool:
    if question.question_type not in (
        QuestionType.SINGLE_SELECT,
        QuestionType.DEMOGRAPHIC_OR_SEGMENT,
    ):
        return False
    if not (2 <= len(question.option_map) <= 12):
        return False
    text_lower = question.question_text.lower()
    return any(keyword in text_lower for keyword in DEMOGRAPHIC_KEYWORDS)


def _identify_respondent_id_column(raw_columns: list[str]) -> str:
    for candidate in RESPONDENT_ID_CANDIDATES:
        if candidate in raw_columns:
            return candidate
    return "respondent_id"
