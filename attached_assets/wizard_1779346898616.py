"""Pure helpers for the Streamlit setup wizard."""

from __future__ import annotations

from dataclasses import replace
from typing import Any

from src.models import QuestionSpec, QuestionType, SurveySchema


WIZARD_STEPS = (
    (1, "Categories"),
    (2, "Demos"),
    (3, "Custom"),
    (4, "PerQ"),
    (5, "CrossCut"),
)


CUSTOM_FILTER_DEFAULT = 2
PER_QUESTION_FILTER_DEFAULT = 1


def question_type_label(question: QuestionSpec) -> str:
    """Return a compact user-facing question type label."""

    labels = {
        QuestionType.SINGLE_SELECT: "Single-select",
        QuestionType.MULTI_SELECT_BINARY: "Multi-select",
        QuestionType.GRID_SINGLE_SELECT: "Grid",
        QuestionType.NUMERIC_ALLOCATION: "Numeric allocation",
        QuestionType.DIRECT_NUMERIC: "Numeric",
        QuestionType.OPEN_TEXT: "Open text",
    }
    return labels.get(question.question_type, question.question_type.value.title())


def question_display_text(question: QuestionSpec, max_chars: int = 60) -> str:
    """Return truncated question text suitable for a compact card."""

    text = str(question.question_text or question.canonical_id).strip()
    if len(text) <= max_chars:
        return text
    return text[: max_chars - 3].rstrip() + "..."


def question_lookup(schema: SurveySchema) -> dict[str, QuestionSpec]:
    return {question.canonical_id: question for question in schema.questions}


def category_assignments_from_themes(
    schema: SurveySchema,
    themes: dict | None,
) -> dict[str, str]:
    """Build question_id -> category from AI themes with a deterministic fallback."""

    eligible_ids = {
        question.canonical_id
        for question in schema.questions
        if question.analysis_eligible
    }
    assignments: dict[str, str] = {}
    for theme in (themes or {}).get("themes", []):
        theme_name = str(theme.get("name") or "Theme").strip() or "Theme"
        for question_id in theme.get("question_ids", []):
            if question_id in eligible_ids and question_id not in assignments:
                assignments[question_id] = theme_name

    for question in schema.questions:
        if not question.analysis_eligible:
            continue
        assignments.setdefault(
            question.canonical_id,
            "Demographics" if question.is_demographic else "All Questions",
        )
    return assignments


def themes_from_category_assignments(
    schema: SurveySchema,
    assignments: dict[str, str],
) -> dict:
    """Convert wizard assignments to the exporter theme payload."""

    ordered: dict[str, list[str]] = {}
    valid_ids = {question.canonical_id for question in schema.questions}
    for question in schema.questions:
        category = assignments.get(question.canonical_id)
        if not category or question.canonical_id not in valid_ids:
            continue
        ordered.setdefault(category, []).append(question.canonical_id)
    return {
        "themes": [
            {"name": category, "question_ids": question_ids}
            for category, question_ids in ordered.items()
            if question_ids
        ],
        "was_template": False,
        "error_message": "",
    }


def themes_from_wizard_assignments(
    schema: SurveySchema,
    assignments: dict[str, str],
) -> dict:
    """Exporter-compatible themes from wizard category assignments."""

    return themes_from_category_assignments(schema, assignments)


def selected_demographics_from_schema(schema: SurveySchema) -> list[str]:
    return [
        question.canonical_id
        for question in schema.questions
        if question.is_demographic and question.analysis_eligible
    ]


def eligible_filter_question_ids(schema: SurveySchema) -> list[str]:
    """Questions that can reasonably act as workbook filters."""

    allowed_types = {
        QuestionType.SINGLE_SELECT,
        QuestionType.GRID_SINGLE_SELECT,
        QuestionType.MULTI_SELECT_BINARY,
    }
    return [
        question.canonical_id
        for question in schema.questions
        if question.analysis_eligible and question.question_type in allowed_types
    ]


def apply_wizard_schema_overrides(
    schema: SurveySchema,
    assignments: dict[str, str],
    selected_demographics: list[str],
) -> SurveySchema:
    """Apply wizard category removal and demographic filter selections."""

    selected_demo_set = set(selected_demographics)
    assignment_ids = set(assignments)
    questions = []
    for question in schema.questions:
        if question.analysis_eligible and question.canonical_id not in assignment_ids:
            questions.append(
                replace(
                    question,
                    analysis_eligible=False,
                    exclusion_reason="removed in setup wizard",
                    is_demographic=False,
                )
            )
            continue
        questions.append(
            replace(
                question,
                is_demographic=question.canonical_id in selected_demo_set,
            )
        )
    return replace(schema, questions=tuple(questions))


def distinct_value_preview(
    dataframe: Any,
    question: QuestionSpec,
    max_values: int = 5,
) -> str:
    """Return a short distinct-value preview for a filter candidate."""

    column = question.canonical_id
    if dataframe is None or not hasattr(dataframe, "columns") or column not in dataframe.columns:
        return "No preview available"
    values = [
        str(value)
        for value in dataframe[column].dropna().unique().tolist()
        if str(value).strip()
    ]
    if not values:
        return "0 distinct values"
    sample = ", ".join(values[:max_values])
    suffix = "" if len(values) <= max_values else ", ..."
    return f"{len(values)} distinct values: {sample}{suffix}"


def normalise_custom_filter_count(value: int | None) -> int:
    return max(0, min(5, int(value if value is not None else CUSTOM_FILTER_DEFAULT)))


def normalise_per_question_filter_count(value: int | None) -> int:
    return max(0, min(3, int(value if value is not None else PER_QUESTION_FILTER_DEFAULT)))
