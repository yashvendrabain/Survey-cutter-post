"""Shared filter and cross-cut option helpers."""

from __future__ import annotations

from dataclasses import dataclass, replace
from typing import Any

import pandas as pd

from src.models import FilterSpec, QuestionSpec, QuestionType, SurveySchema


NPS_BUCKETS: tuple[tuple[str, int, int], ...] = (
    ("Promoter", 9, 10),
    ("Passive", 7, 8),
    ("Detractor", 0, 6),
)


@dataclass(frozen=True, slots=True)
class FilterValueOption:
    value: int | str
    label: str
    filter_question_id: str
    filter_value: int | str


@dataclass(frozen=True, slots=True)
class FilterQuestionOption:
    question_id: str
    label: str
    values: tuple[FilterValueOption, ...]
    question_type: QuestionType
    source_question_id: str


@dataclass(frozen=True, slots=True)
class CrossCutQuestionOption:
    question_id: str
    label: str
    question_type: QuestionType
    source_question_id: str


def filter_question_options(schema: SurveySchema) -> list[FilterQuestionOption]:
    """Return question/value options supported by UI and workbook filters."""

    options: list[FilterQuestionOption] = []
    for question in schema.questions:
        if not getattr(question, "analysis_eligible", False):
            continue
        values = _filter_values_for_question(question)
        if not values:
            continue
        options.append(
            FilterQuestionOption(
                question_id=question.canonical_id,
                label=_question_label(question),
                values=tuple(values),
                question_type=question.question_type,
                source_question_id=question.canonical_id,
            )
        )
    return options


def filter_question_option_map(
    schema: SurveySchema,
) -> dict[str, FilterQuestionOption]:
    return {option.question_id: option for option in filter_question_options(schema)}


def build_filter_specs_from_selection(
    schema: SurveySchema,
    question_id: str,
    selected_values: list[Any],
) -> list[FilterSpec]:
    """Map a UI parent question selection to addressable FilterSpec objects."""

    if not selected_values:
        return [FilterSpec(filter_question_id=question_id)]

    option = filter_question_option_map(schema).get(question_id)
    if option is None:
        if len(selected_values) == 1:
            return [FilterSpec(filter_question_id=question_id, filter_value=selected_values[0])]
        return [
            FilterSpec(
                filter_question_id=question_id,
                filter_values=list(selected_values),
            )
        ]

    by_value = {choice.value: choice for choice in option.values}
    grouped: dict[str, list[int | str]] = {}
    for raw_value in selected_values:
        choice = by_value.get(raw_value)
        if choice is None:
            continue
        grouped.setdefault(choice.filter_question_id, []).append(choice.filter_value)

    specs: list[FilterSpec] = []
    for filter_question_id, values in grouped.items():
        deduped = _dedupe(values)
        if len(deduped) == 1:
            specs.append(
                FilterSpec(
                    filter_question_id=filter_question_id,
                    filter_value=deduped[0],
                )
            )
        else:
            specs.append(
                FilterSpec(
                    filter_question_id=filter_question_id,
                    filter_values=list(deduped),
                )
            )
    return specs


def filter_mask_for_spec(
    df: pd.DataFrame,
    schema: SurveySchema | None,
    filter_spec: FilterSpec,
) -> pd.Series:
    """Evaluate a FilterSpec against decoded data."""

    values = _effective_values(filter_spec)
    if not values:
        return pd.Series(True, index=df.index)

    filter_question_id = filter_spec.filter_question_id
    question = resolve_filter_question(schema, filter_question_id) if schema is not None else None
    if (
        question is not None
        and question.question_type is QuestionType.NPS
        and filter_question_id == question.canonical_id
    ):
        return _nps_question_mask(df, question, values)

    if filter_question_id not in df.columns:
        raise ValueError(f"filter column {filter_question_id!r} not in raw data")

    series = df[filter_question_id]
    if question is not None and question.question_type in {
        QuestionType.MULTI_SELECT_BINARY,
        QuestionType.GRID_BINARY_SELECT,
    }:
        return _selected_series_mask(series, values)

    if question is not None and question.question_type is QuestionType.NPS:
        return _nps_series_mask(series, values)

    option_map = question.option_map if question is not None else {}
    return _filter_values_mask(series, values, option_map)


def resolve_filter_question(
    schema: SurveySchema,
    question_id: str,
) -> QuestionSpec | None:
    direct = schema.get_question(question_id)
    if direct is not None:
        return direct
    for question in schema.questions:
        if question_id in question.raw_columns:
            return question
    return None


def cross_cut_question_options(schema: SurveySchema) -> list[CrossCutQuestionOption]:
    """Return manual cross-cut source options."""

    options: list[CrossCutQuestionOption] = []
    for question in schema.questions:
        if not getattr(question, "analysis_eligible", False):
            continue
        if question.question_type in {
            QuestionType.SINGLE_SELECT,
            QuestionType.DEMOGRAPHIC_OR_SEGMENT,
            QuestionType.NPS,
            QuestionType.DIRECT_NUMERIC,
            QuestionType.NUMERIC_ALLOCATION,
            QuestionType.GRID_RATED,
        }:
            options.append(
                CrossCutQuestionOption(
                    question_id=question.canonical_id,
                    label=_question_label(question),
                    question_type=question.question_type,
                    source_question_id=question.canonical_id,
                )
            )
        elif question.question_type is QuestionType.GRID_SINGLE_SELECT:
            for source_column in question.raw_columns:
                row_label = _row_label(question, source_column)
                options.append(
                    CrossCutQuestionOption(
                        question_id=source_column,
                        label=f"{question.canonical_id}: {row_label}",
                        question_type=QuestionType.SINGLE_SELECT,
                        source_question_id=question.canonical_id,
                    )
                )
    return options


def cross_cut_question_option_map(
    schema: SurveySchema,
) -> dict[str, CrossCutQuestionOption]:
    return {option.question_id: option for option in cross_cut_question_options(schema)}


def resolve_cross_cut_question(
    schema: SurveySchema,
    question_id: str,
) -> QuestionSpec | None:
    direct = schema.get_question(question_id)
    if direct is not None:
        return direct
    parent = resolve_filter_question(schema, question_id)
    if parent is None or question_id not in parent.raw_columns:
        return None
    if parent.question_type is QuestionType.GRID_SINGLE_SELECT:
        row_label = _row_label(parent, question_id)
        return replace(
            parent,
            question_id=f"{parent.question_id}:{question_id}",
            canonical_id=question_id,
            question_text=f"{parent.question_text} - {row_label}",
            question_type=QuestionType.SINGLE_SELECT,
            raw_columns=(question_id,),
            parent_question_id=parent.canonical_id,
            grid_row_labels=None,
        )
    return None


def _filter_values_for_question(question: QuestionSpec) -> list[FilterValueOption]:
    qtype = question.question_type
    if qtype in {QuestionType.SINGLE_SELECT, QuestionType.DEMOGRAPHIC_OR_SEGMENT}:
        return [
            FilterValueOption(
                value=key,
                label=str(label),
                filter_question_id=question.canonical_id,
                filter_value=key,
            )
            for key, label in question.option_map.items()
        ]

    if qtype is QuestionType.MULTI_SELECT_BINARY:
        return [
            FilterValueOption(
                value=source_column,
                label=_row_label(question, source_column),
                filter_question_id=source_column,
                filter_value=1,
            )
            for source_column in question.raw_columns
            if not _is_computed_multi_select_column(question, source_column)
        ]

    if qtype is QuestionType.NPS:
        bucket_values = [
            FilterValueOption(
                value=label,
                label=label,
                filter_question_id=question.canonical_id,
                filter_value=label,
            )
            for label, _low, _high in NPS_BUCKETS
        ]
        exact_values = [
            FilterValueOption(
                value=str(score),
                label=f"Score {score}",
                filter_question_id=question.canonical_id,
                filter_value=str(score),
            )
            for score in range(11)
        ]
        return bucket_values + exact_values

    if qtype is QuestionType.GRID_SINGLE_SELECT:
        values: list[FilterValueOption] = []
        for source_column in question.raw_columns:
            row_label = _row_label(question, source_column)
            for key, label in question.option_map.items():
                token = _composite_value(source_column, key)
                values.append(
                    FilterValueOption(
                        value=token,
                        label=f"{row_label}: {label}",
                        filter_question_id=source_column,
                        filter_value=key,
                    )
                )
        return values

    if qtype is QuestionType.GRID_BINARY_SELECT:
        return [
            FilterValueOption(
                value=source_column,
                label=_row_label(question, source_column),
                filter_question_id=source_column,
                filter_value=1,
            )
            for source_column in question.raw_columns
        ]

    return []


def _question_label(question: QuestionSpec) -> str:
    text = (question.question_text or "").strip()
    if len(text) > 80:
        text = text[:77].rstrip() + "..."
    return f"{question.canonical_id}: {text}" if text else question.canonical_id


def _row_label(question: QuestionSpec, source_column: str) -> str:
    if question.grid_row_labels and source_column in question.grid_row_labels:
        return str(question.grid_row_labels[source_column])
    if source_column in question.option_map:
        return str(question.option_map[source_column])
    if source_column in question.grid_column_labels:
        return str(question.grid_column_labels[source_column])
    return str(source_column)


def _is_computed_multi_select_column(question: QuestionSpec, source_column: str) -> bool:
    label = _row_label(question, source_column)
    haystack = f"{source_column} {label}".lower()
    computed_markers = (
        "computed(",
        "computed_",
        "is_computed",
        "count choices",
        "count_choices",
    )
    return any(marker in haystack for marker in computed_markers)


def _composite_value(source_column: str, value: Any) -> str:
    return f"{source_column}||{value}"


def _effective_values(filter_spec: FilterSpec) -> list[int | str]:
    values = getattr(filter_spec, "filter_values", None)
    if values:
        return list(values)
    if filter_spec.filter_value is None:
        return []
    return [filter_spec.filter_value]


def _filter_values_mask(
    series: pd.Series,
    values: list[int | str],
    option_map: dict[int | str, str],
) -> pd.Series:
    candidates: list[int | str] = []
    for value in values:
        candidates.extend(_expanded_filter_values(value, option_map))
    return series.isin(tuple(_dedupe(candidates)))


def _expanded_filter_values(
    value: int | str,
    option_map: dict[int | str, str],
) -> tuple[int | str, ...]:
    candidates: list[int | str] = [value]
    if value in option_map:
        candidates.append(option_map[value])
    if isinstance(value, str):
        stripped = value.strip()
        candidates.append(stripped)
        try:
            numeric = float(stripped)
        except ValueError:
            numeric = None
        if numeric is not None and numeric.is_integer():
            numeric_int = int(numeric)
            candidates.append(numeric_int)
            if numeric_int in option_map:
                candidates.append(option_map[numeric_int])
    return tuple(_dedupe(candidates))


def _selected_series_mask(
    series: pd.Series,
    values: list[int | str],
) -> pd.Series:
    wants_selected = any(_value_means_selected(value) for value in values)
    if not wants_selected:
        return pd.Series(False, index=series.index)
    return series.map(_is_selected_value).fillna(False)


def _value_means_selected(value: int | str) -> bool:
    if isinstance(value, str):
        return value.strip().lower() in {"1", "selected", "true", "yes", "y"}
    try:
        return float(value) != 0.0
    except (TypeError, ValueError):
        return False


def _is_selected_value(value: Any) -> bool:
    if value is None:
        return False
    try:
        if pd.isna(value):
            return False
    except Exception:
        pass
    if isinstance(value, str):
        return value.strip().lower() not in {"", "0", "false", "no", "unchecked"}
    try:
        return float(value) != 0.0
    except (TypeError, ValueError):
        return True


def _nps_question_mask(
    df: pd.DataFrame,
    question: QuestionSpec,
    values: list[int | str],
) -> pd.Series:
    mask = pd.Series(False, index=df.index)
    for source_column in question.raw_columns:
        if source_column in df.columns:
            mask = mask | _nps_series_mask(df[source_column], values)
    return mask


def _nps_series_mask(
    series: pd.Series,
    values: list[int | str],
) -> pd.Series:
    scores = series.map(_coerce_nps_score)
    mask = pd.Series(False, index=series.index)
    for value in values:
        bucket = _nps_bucket_range(value)
        if bucket is not None:
            low, high = bucket
            mask = mask | ((scores >= low) & (scores <= high))
            continue
        score = _coerce_nps_score(value)
        if score is not None:
            mask = mask | (scores == score)
    return mask.fillna(False)


def _nps_bucket_range(value: Any) -> tuple[int, int] | None:
    label = str(value).strip().lower()
    for bucket_label, low, high in NPS_BUCKETS:
        if label == bucket_label.lower():
            return low, high
    return None


def _coerce_nps_score(value: Any) -> int | None:
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    if isinstance(value, bool):
        return None
    try:
        numeric = float(str(value).strip())
    except (TypeError, ValueError):
        return None
    if not numeric.is_integer():
        return None
    score = int(numeric)
    if 0 <= score <= 10:
        return score
    return None


def _dedupe(values: list[Any]) -> list[Any]:
    deduped: list[Any] = []
    for value in values:
        if value not in deduped:
            deduped.append(value)
    return deduped
