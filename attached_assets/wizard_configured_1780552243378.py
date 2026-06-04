"""Wizard-configured data-map adapter for unknown structured survey formats."""

from __future__ import annotations

from dataclasses import dataclass, field
import re
from typing import Any, Iterable

import pandas as pd

from src.adapters.base import AdapterDetectionResult, DataMapAdapter
from src.datamap_parser import DataMap, ParsedQuestion, _attach_numeric_label_metadata


_RESPONDENT_ID_CANDIDATES = (
    "Respondent",
    "respondent",
    "RespondentID",
    "respondent_id",
    "Respondent ID",
    "Response ID",
    "response_id",
    "record",
    "uuid",
    "id",
    "ID",
)
_HELPER_COLUMN_PATTERNS = (
    r"^qc_",
    r"^helper_",
    r"_helper$",
    r"^straightline_count$",
    r"^status$",
    r"^term",
    r"^vgeoip",
    r"^region$",
    r"^language$",
    r"^panel$",
    r"^browser$",
    r"^ip_address$",
)


@dataclass(frozen=True, slots=True)
class WizardConfig:
    """Parsing config supplied by the survey format wizard."""

    raw_data_sheet_name: str
    data_map_sheet_name: str
    respondent_id_column: str
    question_id_pattern: str
    sub_column_separator: str
    option_code_position: str
    section_prefixes: tuple[str, ...]
    config_name: str | None = None
    helper_columns: tuple[str, ...] = field(default_factory=tuple)
    question_id_example: str = ""
    multi_select_example: str = ""

    def __post_init__(self) -> None:
        object.__setattr__(self, "section_prefixes", tuple(self.section_prefixes))
        object.__setattr__(self, "helper_columns", tuple(self.helper_columns))
        if self.question_id_example and not str(self.question_id_pattern).strip():
            object.__setattr__(
                self,
                "question_id_pattern",
                infer_question_id_regex(self.question_id_example).pattern,
            )
        if self.multi_select_example and not str(self.sub_column_separator).strip():
            object.__setattr__(self, "sub_column_separator", "none")
        required = {
            "raw_data_sheet_name": self.raw_data_sheet_name,
            "data_map_sheet_name": self.data_map_sheet_name,
            "respondent_id_column": self.respondent_id_column,
            "question_id_pattern": self.question_id_pattern,
            "sub_column_separator": self.sub_column_separator,
            "option_code_position": self.option_code_position,
        }
        missing = [name for name, value in required.items() if not str(value).strip()]
        if missing:
            raise ValueError(f"WizardConfig missing required fields: {', '.join(missing)}")
        try:
            re.compile(self.question_id_pattern)
        except re.error as exc:
            raise ValueError(f"invalid question_id_pattern: {exc}") from exc
        if self.option_code_position not in {
            "column_b",
            "column_after_qid",
            "indented_below",
            "same_row",
            "custom",
        }:
            raise ValueError(f"unsupported option_code_position: {self.option_code_position}")


class WizardConfiguredAdapter(DataMapAdapter):
    """Adapter that uses an analyst-supplied wizard config."""

    name = "wizard_configured"

    def __init__(self, config: WizardConfig):
        self._config = config

    def detect(self, workbook: Any, raw_df: Any | None = None) -> AdapterDetectionResult:
        del workbook, raw_df
        return AdapterDetectionResult(confidence=1.0, reason="wizard configured")

    def parse(self, workbook: Any, raw_df: Any | None = None) -> DataMap:
        if self._config.data_map_sheet_name not in getattr(workbook, "sheetnames", []):
            raise ValueError(
                f"data map sheet {self._config.data_map_sheet_name!r} not found"
            )

        worksheet = workbook[self._config.data_map_sheet_name]
        raw_columns = _raw_columns_from_sheet_or_df(workbook, raw_df, self._config)
        if self._config.question_id_example and self._config.multi_select_example:
            questions = _parse_questions_by_example(worksheet, raw_columns, self._config)
        else:
            helper_columns = _configured_helper_columns(raw_columns, self._config)
            questions = _parse_questions_from_sheet(
                worksheet,
                raw_columns=raw_columns,
                helper_columns=helper_columns,
                config=self._config,
            )
        questions = _attach_numeric_label_metadata(questions)
        return {
            "questions": questions,
            "source_path": str(getattr(workbook, "_survey_source_path", "<workbook>")),
            "sheet_name": self._config.data_map_sheet_name,
            "total_rows_in_sheet": int(getattr(worksheet, "max_row", 0) or 0),
            "parser_warnings": [],
        }


def detect_question_id_pattern(sample_values: Iterable[Any]) -> str:
    """Infer a useful question ID regex from data-map samples."""

    samples = [str(value).strip() for value in sample_values if str(value).strip()]
    candidates = (
        (r"^Q\d+", re.compile(r"^Q\d+", re.IGNORECASE)),
        (r"^q\d+", re.compile(r"^q\d+")),
        (r"^Q_\d+", re.compile(r"^Q_\d+", re.IGNORECASE)),
        (r"^Question_\d+", re.compile(r"^Question_\d+", re.IGNORECASE)),
        (r"^Item_\d+", re.compile(r"^Item_\d+", re.IGNORECASE)),
    )
    scored: list[tuple[int, str]] = []
    for pattern, regex in candidates:
        scored.append((sum(1 for value in samples if regex.match(value)), pattern))
    scored.sort(reverse=True)
    return scored[0][1] if scored and scored[0][0] > 0 else r"^Q\d+"


def detect_sub_column_separator(raw_columns: Iterable[Any]) -> str:
    """Infer a sub-column separator from raw data headers."""

    columns = [str(column) for column in raw_columns]
    scores = {
        "r": sum(1 for column in columns if re.match(r"^[A-Za-z]+\d+r\d+", column)),
        "s": sum(1 for column in columns if re.match(r"^[A-Za-z]+\d+s\d+", column)),
        "_": sum(1 for column in columns if re.match(r"^[A-Za-z]+\d+_\d+", column)),
        r"\.": sum(1 for column in columns if re.match(r"^[A-Za-z]+\d+\.\d+", column)),
        ":": sum(1 for column in columns if re.match(r"^[A-Za-z]+\d+\s*:", column)),
    }
    best_separator, best_score = max(scores.items(), key=lambda item: item[1])
    return best_separator if best_score > 0 else "none"


def default_respondent_id_column(raw_columns: Iterable[Any]) -> str:
    columns = [str(column) for column in raw_columns]
    for candidate in _RESPONDENT_ID_CANDIDATES:
        if candidate in columns:
            return candidate
    return columns[0] if columns else "respondent_id"


def default_helper_columns(raw_columns: Iterable[Any]) -> tuple[str, ...]:
    helpers: list[str] = []
    for column in raw_columns:
        text = str(column)
        lowered = text.strip().casefold()
        if any(re.search(pattern, lowered) for pattern in _HELPER_COLUMN_PATTERNS):
            helpers.append(text)
    return tuple(helpers)


def _parse_questions_from_sheet(
    worksheet: Any,
    *,
    raw_columns: tuple[str, ...],
    helper_columns: set[str],
    config: WizardConfig,
) -> list[ParsedQuestion]:
    rows = list(worksheet.iter_rows(values_only=True))
    questions: list[ParsedQuestion] = []
    current: ParsedQuestion | None = None

    for row_number, row in enumerate(rows, start=1):
        values = _normalise_row(row)
        first_cell = values[0] if values else ""
        question_id = _extract_question_id(first_cell, config)
        if question_id is not None:
            if current is not None:
                _finalise_wizard_question(current, raw_columns, helper_columns, config)
                questions.append(current)
            current = _start_wizard_question(question_id, values, row_number)
            _capture_same_row_options(current, values, config)
            continue

        if current is not None:
            _capture_option_row(current, values, config)

    if current is not None:
        _finalise_wizard_question(current, raw_columns, helper_columns, config)
        questions.append(current)
    return questions


def _start_wizard_question(
    question_id: str,
    values: list[str],
    row_number: int,
) -> ParsedQuestion:
    question_text = values[1] if len(values) > 1 and values[1] else question_id
    if question_text == question_id:
        suffix = values[0][len(question_id):].strip(" :-")
        if suffix:
            question_text = suffix
    return {
        "canonical_id": question_id,
        "raw_id": question_id,
        "question_text": question_text,
        "type_hint": "values_range",
        "value_range": None,
        "options": [],
        "sub_columns": [],
        "parent_canonical_id": None,
        "source_row": row_number,
        "warnings": [],
    }


def _capture_same_row_options(
    question: ParsedQuestion,
    values: list[str],
    config: WizardConfig,
) -> None:
    if config.option_code_position not in {"column_after_qid", "same_row"}:
        return
    tail = values[2:]
    for index in range(0, len(tail) - 1, 2):
        _append_option(question, tail[index], tail[index + 1])


def _capture_option_row(
    question: ParsedQuestion,
    values: list[str],
    config: WizardConfig,
) -> None:
    if not values:
        return
    if config.option_code_position in {"column_b", "indented_below", "custom"}:
        if len(values) >= 2:
            _append_option(question, values[0], values[1])
    elif config.option_code_position in {"column_after_qid", "same_row"} and len(values) >= 2:
        _append_option(question, values[0], values[1])


def _append_option(question: ParsedQuestion, code_value: Any, label_value: Any) -> None:
    code_text = str(code_value).strip()
    label_text = str(label_value).strip()
    if not code_text or not label_text:
        return
    if re.match(r"^[A-Za-z_]+\d+", code_text):
        return
    code: int | str
    if re.match(r"^-?\d+$", code_text):
        code = int(code_text)
    else:
        code = code_text
    question["options"].append((code, label_text))


def _finalise_wizard_question(
    question: ParsedQuestion,
    raw_columns: tuple[str, ...],
    helper_columns: set[str],
    config: WizardConfig,
) -> None:
    question["sub_columns"] = _infer_sub_columns(
        question["canonical_id"],
        raw_columns,
        helper_columns,
        config.sub_column_separator,
    )
    question["value_range"] = _value_range_from_options(question["options"])
    if not question["options"]:
        question["type_hint"] = "open_numeric"


def _infer_sub_columns(
    question_id: str,
    raw_columns: tuple[str, ...],
    helper_columns: set[str],
    separator: str,
) -> list[tuple[str, str]]:
    if separator in {"none", "No multi-part questions in this survey"}:
        return []

    inferred: list[tuple[str, str]] = []
    for column in raw_columns:
        if column in helper_columns or column == question_id:
            continue
        label = _sub_column_label(question_id, column, separator)
        if label is None:
            continue
        inferred.append((column, label))
    return inferred


def _sub_column_label(question_id: str, column: str, separator: str) -> str | None:
    text = str(column)
    if separator == ":":
        match = re.match(rf"^{re.escape(question_id)}\s*:\s*(.+)$", text, re.IGNORECASE)
    else:
        match = re.match(
            rf"^{re.escape(question_id)}(?:{separator})(.+)$",
            text,
            re.IGNORECASE,
        )
    if match is None:
        return None
    label = re.sub(r"\s+", " ", match.group(1)).strip()
    return label or text


def _value_range_from_options(
    options: list[tuple[int | str, str]],
) -> tuple[int, int] | None:
    numeric_codes = [code for code, _label in options if isinstance(code, int)]
    if not numeric_codes:
        return None
    return (int(min(numeric_codes)), int(max(numeric_codes)))


def _extract_question_id(cell_value: Any, config: WizardConfig) -> str | None:
    text = str(cell_value or "").strip()
    if not text:
        return None
    match = re.compile(config.question_id_pattern).match(text)
    if match is not None:
        return match.group(0).strip()
    prefix_match = re.match(r"^([A-Za-z]+)(\d+[A-Za-z0-9_]*)", text)
    if prefix_match and prefix_match.group(1) in set(config.section_prefixes):
        return prefix_match.group(0)
    return None


def _normalise_row(row: Iterable[Any]) -> list[str]:
    return ["" if value is None else str(value).strip() for value in row]


def _raw_columns_from_sheet_or_df(
    workbook: Any,
    raw_df: Any | None,
    config: WizardConfig,
) -> tuple[str, ...]:
    if raw_df is not None:
        return tuple(str(column) for column in getattr(raw_df, "columns", []))
    if config.raw_data_sheet_name not in getattr(workbook, "sheetnames", []):
        return tuple()
    values = workbook[config.raw_data_sheet_name].iter_rows(
        min_row=1,
        max_row=1,
        values_only=True,
    )
    first_row = next(values, tuple())
    return tuple(str(value).strip() for value in first_row if value is not None)


def _configured_helper_columns(
    raw_columns: tuple[str, ...],
    config: WizardConfig,
) -> set[str]:
    helpers = set(str(column) for column in config.helper_columns)
    helpers.update(default_helper_columns(raw_columns))
    return helpers


def dataframe_from_workbook_sheet(workbook: Any, sheet_name: str) -> pd.DataFrame:
    """Build a DataFrame from a worksheet using the first row as headers."""

    worksheet = workbook[sheet_name]
    rows = list(worksheet.iter_rows(values_only=True))
    if not rows:
        return pd.DataFrame()
    headers = [str(value).strip() if value is not None else "" for value in rows[0]]
    data = [list(row) for row in rows[1:]]
    return pd.DataFrame(data, columns=headers)


_EXAMPLE_SEP_CLASS = r":\-_/|.\u2013\u2014\s"


def _tokenize_example(example: str) -> list[tuple[str, str]]:
    """Split into alternating ('int', digits) / ('lit', non-digits) runs."""
    tokens: list[tuple[str, str]] = []
    for match in re.finditer(r"\d+|\D+", str(example).strip()):
        piece = match.group(0)
        tokens.append(("int", piece) if piece.isdigit() else ("lit", piece))
    return tokens


def infer_question_id_regex(example: str) -> "re.Pattern[str]":
    """Infer a question-ID regex from ONE example (e.g. 'q1','1q','q-1','Q15').

    The first run of digits is the captured question key <key>; everything else
    is literal. Anchored both ends so a single-answer form can't also match a
    longer multi-answer column.
    """

    segments = _tokenize_example(example)
    int_positions = [i for i, (kind, _) in enumerate(segments) if kind == "int"]
    if not int_positions:
        raise ValueError(f"no integer found in question-id example {example!r}")
    key_index = int_positions[0]
    parts: list[str] = []
    for index, (kind, value) in enumerate(segments):
        if index == key_index:
            parts.append(r"(?P<key>\d+)")
        elif kind == "lit":
            parts.append(re.escape(value))
        else:
            parts.append(r"\d+")
    return re.compile("^" + "".join(parts) + "$", re.IGNORECASE)


def infer_multi_select_spec(example: str) -> "tuple[str, re.Pattern[str], re.Pattern[str]]":
    """Infer multi-select column matching from ONE example.

    Returns (mode, multi_regex, single_regex):
      mode 'index' -> two integer runs: groups <key>,<opt>  ('6q1a','Q6r1','q6_1')
      mode 'label' -> one integer run + separator + label    ('Q6: Field sales')
    single_regex matches a single-answer column for the same key (the multi
    pattern with the option part removed); used to link single-selects by key.
    """

    segments = _tokenize_example(example)
    int_positions = [i for i, (kind, _) in enumerate(segments) if kind == "int"]
    if not int_positions:
        raise ValueError(f"no integer found in multi-select example {example!r}")

    if len(int_positions) >= 2:
        key_index, opt_index = int_positions[0], int_positions[1]
        multi_parts: list[str] = []
        single_parts: list[str] = []
        for index, (kind, value) in enumerate(segments):
            if index == key_index:
                multi_parts.append(r"(?P<key>\d+)")
                single_parts.append(r"(?P<key>\d+)")
            elif index == opt_index:
                multi_parts.append(r"(?P<opt>\d+)")
            elif kind == "lit":
                multi_parts.append(re.escape(value))
                if index < opt_index:
                    single_parts.append(re.escape(value))
            else:
                multi_parts.append(r"\d+")
        multi_regex = re.compile("^" + "".join(multi_parts) + "$", re.IGNORECASE)
        single_regex = re.compile("^" + "".join(single_parts) + "$", re.IGNORECASE)
        return "index", multi_regex, single_regex

    key_index = int_positions[0]
    prefix = "".join(
        re.escape(value) if kind == "lit" else r"\d+"
        for kind, value in segments[:key_index]
    )
    multi_regex = re.compile(
        "^" + prefix + r"(?P<key>\d+)\s*[" + _EXAMPLE_SEP_CLASS + r"]+\s*(?P<label>.+?)\s*$",
        re.IGNORECASE,
    )
    single_regex = re.compile("^" + prefix + r"(?P<key>\d+)$", re.IGNORECASE)
    return "label", multi_regex, single_regex


def _new_question_by_key(
    canonical_id: str,
    question_key: int,
    question_text: str,
    source_row: int,
) -> ParsedQuestion:
    question: ParsedQuestion = {
        "canonical_id": canonical_id,
        "raw_id": canonical_id,
        "question_text": question_text,
        "type_hint": "values_range",
        "value_range": None,
        "options": [],
        "sub_columns": [],
        "parent_canonical_id": None,
        "source_row": source_row,
        "warnings": [],
    }
    question["_key"] = question_key  # type: ignore[typeddict-unknown-key]
    return question


def _codebook_questions_by_key(worksheet: Any, qid_regex: "re.Pattern[str]") -> list[ParsedQuestion]:
    """Walk a compact codebook; start a question on a qid-matching col A with
    text in col B; collect numeric-coded option rows. A note row (blank col A,
    non-blank col B) closes the current question's option list."""
    questions: list[ParsedQuestion] = []
    current: ParsedQuestion | None = None
    options_closed = False
    for row_number, row in enumerate(
        worksheet.iter_rows(min_col=1, max_col=2, values_only=True), start=1
    ):
        col_a = "" if (len(row) < 1 or row[0] is None) else str(row[0]).strip()
        col_b = "" if (len(row) < 2 or row[1] is None) else str(row[1]).strip()
        id_match = qid_regex.match(col_a) if col_a else None
        if id_match is not None:
            current = _new_question_by_key(
                col_a, int(id_match.group("key")), col_b or col_a, row_number
            )
            questions.append(current)
            options_closed = False
            continue
        if current is None:
            continue
        if not col_a and col_b:
            options_closed = True
            continue
        if col_a and re.fullmatch(r"-?\d+", col_a) and col_b and not options_closed:
            current["options"].append((int(col_a), col_b))
    return questions


def _link_raw_columns_by_key(
    questions: list[ParsedQuestion],
    raw_columns: tuple[str, ...],
    mode: str,
    multi_regex: "re.Pattern[str]",
    single_regex: "re.Pattern[str]",
) -> None:
    by_key = {question["_key"]: question for question in questions}  # type: ignore[typeddict-item]
    buckets: dict[int, list[tuple[Any, Any, str]]] = {}
    for column in raw_columns:
        text = str(column)
        match = multi_regex.match(text)
        if match is None:
            continue
        key = int(match.group("key"))
        if mode == "index":
            buckets.setdefault(key, []).append((int(match.group("opt")), None, text))
        else:
            buckets.setdefault(key, []).append((None, match.group("label").strip(), text))
    for key, items in buckets.items():
        question = by_key.get(key)
        if question is None:
            continue
        if mode == "index":
            items.sort(key=lambda item: item[0])
            option_label = {code: label for code, label in question["options"]}
            question["sub_columns"] = [
                (column, option_label.get(opt, f"option {opt}"))
                for opt, _none, column in items
            ]
        else:
            question["sub_columns"] = [(column, label) for _none, label, column in items]
    for column in raw_columns:
        text = str(column)
        if multi_regex.match(text) is not None:
            continue
        match = single_regex.match(text)
        if match is None:
            continue
        question = by_key.get(int(match.group("key")))
        if (
            question is not None
            and not question["sub_columns"]
            and question["raw_id"] == question["canonical_id"]
        ):
            question["raw_id"] = text
    for question in questions:
        numeric_codes = [
            code for code, _label in question["options"] if isinstance(code, int)
        ]
        question["value_range"] = (
            (min(numeric_codes), max(numeric_codes)) if numeric_codes else None
        )
        if not question["options"] and not question["sub_columns"]:
            question["type_hint"] = "open_numeric"


def _parse_questions_by_example(
    worksheet: Any,
    raw_columns: tuple[str, ...],
    config: WizardConfig,
) -> list[ParsedQuestion]:
    qid_regex = infer_question_id_regex(config.question_id_example)
    mode, multi_regex, single_regex = infer_multi_select_spec(config.multi_select_example)
    questions = _codebook_questions_by_key(worksheet, qid_regex)
    _link_raw_columns_by_key(questions, raw_columns, mode, multi_regex, single_regex)
    for question in questions:
        question.pop("_key", None)  # type: ignore[misc]
    return questions
