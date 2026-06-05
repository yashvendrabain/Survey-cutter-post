"""Adapter for six-column combined xlsx survey workbooks."""

from __future__ import annotations

from dataclasses import dataclass
import re
from typing import Any

import pandas as pd

from src.adapters.base import AdapterDetectionResult, DataMapAdapter
from src.datamap_parser import ParsedQuestion, _attach_numeric_label_metadata


KNOWN_TYPE_LABELS = {
    "multiple choice",
    "matrix",
    "columns",
    "text input",
    "rank",
    "allocation",
}


class SixColumnCombinedAdapter(DataMapAdapter):
    """Parser for Datamap/Raw Data workbooks with six datamap columns."""

    name = "six_column_combined"

    def detect(self, workbook: Any, raw_df: Any | None = None) -> AdapterDetectionResult:
        del raw_df
        raw_sheet = _find_raw_sheet(workbook)
        datamap_sheet = _find_datamap_sheet(workbook)
        if raw_sheet is None or datamap_sheet is None:
            return AdapterDetectionResult(0.0, "Raw Data + Datamap sheets not found")

        worksheet = workbook[datamap_sheet]
        first_row = next(worksheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
        if len(first_row) < 6:
            return AdapterDetectionResult(0.0, "Datamap sheet is not six columns")
        if _coerce_question_number(first_row[1]) is None:
            return AdapterDetectionResult(0.0, "Datamap row 1 col2 is not a question number")
        if _normalise_type_label(first_row[3]) not in KNOWN_TYPE_LABELS:
            return AdapterDetectionResult(0.0, "Datamap row 1 col4 is not a known type label")

        return AdapterDetectionResult(
            0.95,
            f"six-column Datamap + Raw Data workbook detected ({datamap_sheet!r}, {raw_sheet!r})",
        )

    def parse(self, workbook: Any, raw_df: Any | None = None) -> dict:
        datamap_sheet = _find_datamap_sheet(workbook)
        raw_sheet = _find_raw_sheet(workbook)
        if datamap_sheet is None or raw_sheet is None:
            raise ValueError("Raw Data + Datamap sheets not found")

        raw_columns = _raw_columns_from_workbook(workbook, raw_sheet, raw_df)
        raw_values = _raw_value_samples(workbook, raw_sheet, raw_columns, raw_df)
        questions: list[ParsedQuestion] = []
        parser_warnings: list[str] = []
        worksheet = workbook[datamap_sheet]
        total_rows = 0

        for row_number, row in enumerate(worksheet.iter_rows(values_only=True), start=1):
            total_rows = row_number
            if _row_is_blank(row):
                continue
            question_number = _coerce_question_number(row[1] if len(row) > 1 else None)
            type_label = _normalise_type_label(row[3] if len(row) > 3 else None)
            if question_number is None or type_label not in KNOWN_TYPE_LABELS:
                parser_warnings.append(f"ignored non-question row {row_number}")
                continue
            subtype = _normalise_subtype(row[4] if len(row) > 4 else None)
            question_text = str(row[5]).strip() if len(row) > 5 and row[5] is not None else f"Q{question_number}"
            qid = f"Q{question_number}"
            questions.append(
                _build_question(
                    qid=qid,
                    question_text=question_text,
                    type_label=type_label,
                    subtype=subtype,
                    raw_columns=raw_columns,
                    raw_values=raw_values,
                    source_row=row_number,
                )
            )

        questions = _attach_numeric_label_metadata(questions)
        return {
            "questions": questions,
            "source_path": str(getattr(workbook, "_survey_source_path", "<workbook>")),
            "sheet_name": datamap_sheet,
            "total_rows_in_sheet": total_rows,
            "parser_warnings": parser_warnings,
        }


@dataclass(frozen=True, slots=True)
class _RawColumnInfo:
    qid: str
    column: str
    detail: str
    entity: str
    is_computed: bool
    is_user_input: bool


def _build_question(
    *,
    qid: str,
    question_text: str,
    type_label: str,
    subtype: str,
    raw_columns: tuple[str, ...],
    raw_values: dict[str, list[Any]],
    source_row: int,
) -> ParsedQuestion:
    matching = [_parse_raw_column(column) for column in raw_columns]
    matching = [info for info in matching if info is not None and info.qid == qid]
    is_nps = _looks_like_nps_text(question_text)
    include_user_input = type_label in {"text input", "columns"}
    answer_columns = [
        info
        for info in matching
        if not info.is_computed and (include_user_input or not info.is_user_input)
    ]

    type_hint = "values_range"
    value_range: tuple[int, int] | None = None
    options: list[tuple[int | str, str]] = []
    sub_columns: list[tuple[str, str]] = []

    if is_nps:
        value_range = (0, 10)
        sub_columns = [
            (info.column, _entity_label(info, question_text))
            for info in answer_columns
        ]
    elif type_label == "text input" or type_label == "columns":
        type_hint = "open_text"
        sub_columns = [
            (info.column, _entity_label(info, question_text))
            for info in answer_columns
            if info.column != qid
        ]
    elif type_label == "multiple choice" and subtype == "single-select":
        columns = [info.column for info in answer_columns] or [qid]
        options = _infer_options(columns, raw_values)
        value_range = _infer_value_range(columns, raw_values)
    elif type_label == "multiple choice" and subtype == "multi-select":
        value_range = (0, 1)
        sub_columns = [(info.column, _option_label(info)) for info in answer_columns]
    elif type_label == "matrix":
        columns = [info.column for info in answer_columns]
        sub_columns = [(info.column, _matrix_label(info)) for info in answer_columns]
        if subtype == "multi-select":
            value_range = (0, 1)
            options = [(0, "Not selected"), (1, "Selected")]
        else:
            value_range = _infer_value_range(columns, raw_values)
            options = _infer_options(columns, raw_values)
    elif type_label == "rank":
        sub_columns = [(info.column, _option_label(info)) for info in answer_columns]
        value_range = (1, max(1, len(sub_columns)))
    elif type_label == "allocation":
        sub_columns = [(info.column, _option_label(info)) for info in answer_columns]
        value_range = (0, 999)

    return {
        "canonical_id": qid,
        "raw_id": qid,
        "question_text": question_text,
        "type_hint": type_hint,
        "value_range": value_range,
        "options": options,
        "sub_columns": sub_columns,
        "parent_canonical_id": None,
        "source_row": source_row,
        "warnings": [],
    }


def _parse_raw_column(column: Any) -> _RawColumnInfo | None:
    text = str(column).strip()
    match = re.match(r"^(Q\d+)(?P<rest>(?:\s|:|\|).*)?$", text, re.IGNORECASE)
    if match is None:
        return None
    qid = match.group(1).upper()
    rest = (match.group("rest") or "").strip()
    is_computed = "computed(answered)" in rest.casefold()
    is_user_input = "::" in rest and "user input" in rest.casefold()
    detail = ""
    entity = ""
    if "::" in rest:
        detail = rest.split("::", 1)[1].strip()
    elif "|" in rest:
        left, right = rest.split("|", 1)
        detail = left.lstrip(":").strip()
        entity = right.strip()
    elif rest.startswith(":"):
        detail = rest[1:].strip()
    return _RawColumnInfo(
        qid=qid,
        column=text,
        detail=detail,
        entity=entity,
        is_computed=is_computed,
        is_user_input=is_user_input,
    )


def _option_label(info: _RawColumnInfo) -> str:
    return info.detail or info.entity or info.column


def _matrix_label(info: _RawColumnInfo) -> str:
    if info.entity and info.detail:
        return f"{info.detail} | {info.entity}"
    return info.detail or info.entity or info.column


def _entity_label(info: _RawColumnInfo, question_text: str) -> str:
    if info.entity:
        return info.entity
    if info.detail:
        return info.detail
    extracted = _single_entity_from_question_text(question_text)
    return extracted or info.column


def _single_entity_from_question_text(question_text: str) -> str | None:
    match = re.search(
        r"\brecommend\s+(?P<entity>.+?)\s+to\s+an\s+interested\b",
        question_text,
        flags=re.IGNORECASE,
    )
    if match is None:
        return None
    entity = re.sub(r"\s+", " ", match.group("entity")).strip()
    if not entity or "<" in entity or "following" in entity.casefold():
        return None
    return entity


def _infer_options(columns: list[str], raw_values: dict[str, list[Any]]) -> list[tuple[int | str, str]]:
    values: list[Any] = []
    for column in columns:
        values.extend(raw_values.get(column, []))
    deduped: list[Any] = []
    for value in values:
        if _is_missing(value):
            continue
        scalar = _normalise_option_value(value)
        if scalar not in deduped:
            deduped.append(scalar)
        if len(deduped) >= 50:
            break
    deduped.sort(key=lambda value: (0, float(value)) if _is_number(value) else (1, str(value)))
    return [(value, str(value)) for value in deduped]


def _infer_value_range(columns: list[str], raw_values: dict[str, list[Any]]) -> tuple[int, int] | None:
    ints: list[int] = []
    for column in columns:
        for value in raw_values.get(column, []):
            numeric = _coerce_int(value)
            if numeric is not None:
                ints.append(numeric)
    if not ints:
        return None
    return (min(ints), max(ints))


def _normalise_option_value(value: Any) -> int | str:
    numeric = _coerce_int(value)
    if numeric is not None:
        return numeric
    return str(value).strip()


def _coerce_int(value: Any) -> int | None:
    if _is_missing(value) or isinstance(value, bool):
        return None
    try:
        numeric = float(str(value).strip())
    except (TypeError, ValueError):
        return None
    if numeric.is_integer():
        return int(numeric)
    return None


def _is_number(value: Any) -> bool:
    try:
        float(value)
    except (TypeError, ValueError):
        return False
    return True


def _is_missing(value: Any) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return isinstance(value, str) and not value.strip()


def _raw_columns_from_workbook(
    workbook: Any,
    raw_sheet: str,
    raw_df: Any | None,
) -> tuple[str, ...]:
    if raw_df is not None:
        return tuple(str(column) for column in getattr(raw_df, "columns", raw_df))
    worksheet = workbook[raw_sheet]
    first_row = next(worksheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
    return tuple(str(cell).strip() for cell in first_row if cell is not None and str(cell).strip())


def _raw_value_samples(
    workbook: Any,
    raw_sheet: str,
    raw_columns: tuple[str, ...],
    raw_df: Any | None,
    *,
    max_rows: int = 500,
) -> dict[str, list[Any]]:
    if raw_df is not None and hasattr(raw_df, "columns"):
        return {
            str(column): list(raw_df[column].dropna().head(max_rows))
            for column in raw_df.columns
        }

    wanted = set(raw_columns)
    values = {column: [] for column in raw_columns}
    worksheet = workbook[raw_sheet]
    header = next(worksheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
    column_indexes = {
        index: str(cell).strip()
        for index, cell in enumerate(header)
        if cell is not None and str(cell).strip() in wanted
    }
    for row_offset, row in enumerate(worksheet.iter_rows(min_row=2, values_only=True), start=1):
        if row_offset > max_rows:
            break
        for index, column in column_indexes.items():
            if index < len(row):
                values[column].append(row[index])
    return values


def _find_raw_sheet(workbook: Any) -> str | None:
    for sheet_name in getattr(workbook, "sheetnames", []):
        if _normalise_sheet_name(sheet_name) == "rawdata":
            return str(sheet_name)
    return None


def _find_datamap_sheet(workbook: Any) -> str | None:
    for sheet_name in getattr(workbook, "sheetnames", []):
        if _normalise_sheet_name(sheet_name) == "datamap":
            return str(sheet_name)
    return None


def _normalise_sheet_name(value: Any) -> str:
    return re.sub(r"[\s_]+", "", str(value).strip().casefold())


def _coerce_question_number(value: Any) -> int | None:
    if value is None:
        return None
    try:
        numeric = float(str(value).strip())
    except (TypeError, ValueError):
        return None
    if not numeric.is_integer():
        return None
    return int(numeric)


def _normalise_type_label(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip().casefold())


def _normalise_subtype(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip().casefold())


def _row_is_blank(row: tuple[Any, ...]) -> bool:
    return all(_is_missing(value) for value in row)


def _looks_like_nps_text(question_text: str) -> bool:
    text = str(question_text)
    if re.search(r"\brecommend\b", text, re.IGNORECASE) is None:
        return False
    return re.search(
        r"(?:\b0\s*(?:-|–|—|to)\s*10\b|\bscale\s+of\s+0\b)",
        text,
        re.IGNORECASE,
    ) is not None
