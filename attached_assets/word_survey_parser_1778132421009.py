"""Word survey parser for the Survey Insight Engine.

This module converts supported .docx survey scripts into the same DataMap
shape emitted by src.datamap_parser.parse_datamap.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
import re
from typing import Literal

from docx import Document
from docx.document import Document as DocxDocument
from docx.text.paragraph import Paragraph

from src.datamap_parser import DataMap, ParsedQuestion, TypeHint


FORMAT_A_SIGNAL_RE = re.compile(
    r"^(Q\d+[A-Za-z0-9_]*|H_Q\d+[A-Za-z0-9_]*|S_\w+)[\.\s:]"
)
FORMAT_B_SIGNAL_RE = re.compile(
    r"^(Multiple choice|Single-select|Multi-select|Matrix|Type-in)",
    re.IGNORECASE,
)
QID_PATTERN = re.compile(
    r"^("
    r"Q\d+[A-Za-z0-9_]*"
    r"|H_Q\d+[A-Za-z0-9_]*"
    r"|Hidden_Q\d+[A-Za-z0-9_]*"
    r"|S_[A-Za-z0-9_]+"
    r")[\.\:\s]+(.*)$",
    re.IGNORECASE,
)
TYPE_ANNOTATION_PATTERN = re.compile(
    r"\[(SINGLE[- ]SELECT|MULTI[- ]SELECT|ESSAY|OPEN[- ]END|OPEN END|"
    r"NUMERIC|DROP\s?DOWN|MATRIX|GRID|RANK|SCALE)[^\]]*\]",
    re.IGNORECASE,
)
PROGRAMMER_INSTRUCTION_PATTERN = re.compile(
    r"\[(TERMINATE|ANCHOR|EXCLUSIVE|ALPHABETIZE|RANDOMIZE|RANDOM ORDER|"
    r"ORDERED|DISPLAY|HIDE|BASE|PN:|RED HERRING|TAG:|QUALIFY)[^\]]*\]",
    re.IGNORECASE,
)
BRACKET_ID_PATTERN = re.compile(r"\[([A-Za-z][A-Za-z0-9_]+)\]")
BRACKET_ANNOTATION_PATTERN = re.compile(r"\[[^\]]+\]")
OPTION_CODE_PATTERN = re.compile(r"^(\d+)[\.\:]\s*(.*)$")
TYPE_B_TYPE_PATTERN = re.compile(
    r"(Multiple choice|Single-select|Multi-select|Matrix|Type-in|Scale|Rank|"
    r"Essay|Open-end)",
    re.IGNORECASE,
)
PARENT_ROW_OE_RE = re.compile(r"^([A-Za-z][A-Za-z0-9_]*?)r\d+oe$", re.IGNORECASE)
PARENT_OE_RE = re.compile(r"^([A-Za-z][A-Za-z0-9_]*?)oe$", re.IGNORECASE)


FormatName = Literal["FORMAT_A", "FORMAT_B"]


class _FormatBState(Enum):
    BETWEEN = "BETWEEN"
    IN_QUESTION_TEXT = "IN_QUESTION_TEXT"
    IN_TYPE_HINT = "IN_TYPE_HINT"
    IN_OPTIONS = "IN_OPTIONS"


@dataclass
class _ParagraphView:
    text: str
    is_bullet: bool
    is_heading: bool
    source_row: int


@dataclass
class _QuestionBlock:
    canonical_id: str
    raw_id: str
    question_text_parts: list[str]
    source_row: int
    options: list[tuple[int, str]] = field(default_factory=list)
    sub_columns: list[tuple[str, str]] = field(default_factory=list)
    annotations: list[str] = field(default_factory=list)
    type_hint: TypeHint | None = None
    value_range: tuple[int, int] | None = None
    next_option_code: int = 1


def parse_word_survey(path: str) -> DataMap:
    """Parse a supported Word survey document into a DataMap."""

    document = Document(path)
    paragraphs = _paragraph_views(document)
    detected_format = _detect_format(paragraphs)
    if detected_format == "FORMAT_A":
        data_map = _parse_format_a(path, paragraphs)
        if data_map["questions"]:
            return data_map
    return _parse_format_b(path, paragraphs)


def _detect_format(paragraphs: list[_ParagraphView]) -> FormatName:
    non_empty = [paragraph.text for paragraph in paragraphs if paragraph.text][:30]
    format_a_score = sum(1 for text in non_empty if FORMAT_A_SIGNAL_RE.match(text))
    format_b_score = sum(1 for text in non_empty if FORMAT_B_SIGNAL_RE.match(text))
    if format_a_score > format_b_score:
        return "FORMAT_A"
    if format_b_score > format_a_score:
        return "FORMAT_B"
    return "FORMAT_A"


def _parse_format_a(path: str, paragraphs: list[_ParagraphView]) -> DataMap:
    questions: list[ParsedQuestion] = []
    parser_warnings: list[str] = []
    current: _QuestionBlock | None = None

    for paragraph in paragraphs:
        text = paragraph.text
        if not text:
            continue

        qid_match = None if paragraph.is_bullet else QID_PATTERN.match(text)
        if qid_match:
            if current is not None:
                questions.append(_finalise_block(current))
            current = _start_format_a_block(qid_match, text, paragraph.source_row)
            continue

        if current is None:
            continue

        if paragraph.is_bullet:
            _add_option(current, text)
            continue

        annotations = _annotations(text)
        if annotations:
            current.annotations.extend(annotations)
            if TYPE_ANNOTATION_PATTERN.search(text):
                continue

        if current.options:
            continue

        cleaned = _clean_text(text)
        if cleaned:
            current.question_text_parts.append(cleaned)

    if current is not None:
        questions.append(_finalise_block(current))

    return _data_map(path, paragraphs, questions, parser_warnings)


def _parse_format_b(path: str, paragraphs: list[_ParagraphView]) -> DataMap:
    questions: list[ParsedQuestion] = []
    parser_warnings: list[str] = []
    current: _QuestionBlock | None = None
    state = _FormatBState.BETWEEN
    question_counter = 0

    for paragraph in paragraphs:
        text = paragraph.text
        if not text:
            continue

        type_line = _type_b_line(text)
        if current is not None and type_line is not None:
            _apply_type_b_hint(current, type_line)
            state = _FormatBState.IN_TYPE_HINT
            continue

        if paragraph.is_bullet:
            if current is not None:
                _add_option(current, text)
                state = _FormatBState.IN_OPTIONS
            continue

        if state in {_FormatBState.IN_TYPE_HINT, _FormatBState.IN_OPTIONS}:
            questions.append(_finalise_block(current))
            current = None
            state = _FormatBState.BETWEEN

        if current is None:
            question_counter += 1
            canonical_id = f"Q{question_counter:02d}"
            current = _QuestionBlock(
                canonical_id=canonical_id,
                raw_id=canonical_id,
                question_text_parts=[_clean_text(text)],
                source_row=paragraph.source_row,
            )
            state = _FormatBState.IN_QUESTION_TEXT
            continue

        cleaned = _clean_text(text)
        if cleaned:
            current.question_text_parts.append(cleaned)
            state = _FormatBState.IN_QUESTION_TEXT

    if current is not None:
        questions.append(_finalise_block(current))

    return _data_map(path, paragraphs, questions, parser_warnings)


def _paragraph_views(document: DocxDocument) -> list[_ParagraphView]:
    return [
        _ParagraphView(
            text=_normalise_text(paragraph.text),
            is_bullet=_is_bullet(paragraph),
            is_heading=_is_heading(paragraph),
            source_row=index,
        )
        for index, paragraph in enumerate(document.paragraphs, start=1)
    ]


def _is_bullet(paragraph: Paragraph) -> bool:
    style_name = paragraph.style.name.lower() if paragraph.style is not None else ""
    if "list" in style_name:
        return True
    paragraph_properties = paragraph._p.pPr
    return bool(
        paragraph_properties is not None
        and paragraph_properties.numPr is not None
    )


def _is_heading(paragraph: Paragraph) -> bool:
    style_name = paragraph.style.name if paragraph.style is not None else ""
    if style_name.startswith("Heading"):
        return True
    return bool(paragraph.runs and paragraph.runs[0].bold)


def _start_format_a_block(
    qid_match: re.Match[str],
    raw_text: str,
    source_row: int,
) -> _QuestionBlock:
    raw_id = qid_match.group(1).strip().rstrip(".:")
    canonical_id = _canonical_id(raw_id)
    remainder = qid_match.group(2).strip()
    annotations = _annotations(raw_text)
    cleaned_question = _clean_text(remainder)
    return _QuestionBlock(
        canonical_id=canonical_id,
        raw_id=raw_id,
        question_text_parts=[cleaned_question] if cleaned_question else [],
        source_row=source_row,
        annotations=annotations,
    )


def _add_option(block: _QuestionBlock, raw_text: str) -> None:
    block.annotations.extend(_annotations(raw_text))
    cleaned = _clean_text(PROGRAMMER_INSTRUCTION_PATTERN.sub("", raw_text))
    cleaned = _clean_text(cleaned)
    if not cleaned:
        return

    code_match = OPTION_CODE_PATTERN.match(cleaned)
    if code_match:
        code = int(code_match.group(1))
        label = _clean_text(code_match.group(2))
    else:
        code = block.next_option_code
        label = cleaned
        block.next_option_code += 1

    if label:
        block.options.append((code, label))


def _finalise_block(block: _QuestionBlock) -> ParsedQuestion:
    question_text = _clean_text(" ".join(block.question_text_parts))
    type_hint, value_range = _derive_type_hint(block, question_text)
    parent_canonical_id = (
        _derive_parent_canonical_id(block.canonical_id)
        if type_hint == "open_text"
        else None
    )
    options = list(block.options)
    warnings = [f"ANNOTATION: {annotation}" for annotation in block.annotations]
    if not question_text:
        warnings.append(
            "WARNING: question text is empty; "
            "question may have been in a table or "
            "have non-standard formatting"
        )
    if len(options) > 50:
        warnings.append(
            f"WARNING: {len(options)} options detected — "
            f"this may indicate nested sub-questions were "
            f"absorbed into this question block. Review "
            f"manually."
        )
    return {
        "canonical_id": block.canonical_id,
        "raw_id": block.raw_id,
        "question_text": question_text,
        "type_hint": type_hint,
        "value_range": value_range,
        "options": options,
        "sub_columns": list(block.sub_columns),
        "parent_canonical_id": parent_canonical_id,
        "source_row": block.source_row,
        "warnings": warnings,
    }


def _derive_type_hint(
    block: _QuestionBlock,
    question_text: str,
) -> tuple[TypeHint | None, tuple[int, int] | None]:
    if block.type_hint is not None:
        if block.type_hint == "values_range" and block.value_range is None:
            return "values_range", _default_value_range(block)
        return block.type_hint, block.value_range

    annotation_text = " ".join(block.annotations).upper()
    if "MULTI" in annotation_text and "SELECT" in annotation_text:
        return "values_range", (0, 1)
    if (
        "SINGLE" in annotation_text
        or "DROP" in annotation_text
        or "MATRIX" in annotation_text
        or "GRID" in annotation_text
        or "SCALE" in annotation_text
        or "RANK" in annotation_text
    ):
        return "values_range", _default_value_range(block)
    if (
        "ESSAY" in annotation_text
        or "OPEN" in annotation_text
        or "TEXT BOX" in annotation_text
    ):
        return "open_text", None
    if "NUMERIC" in annotation_text:
        return "open_numeric", None
    if block.canonical_id.lower().endswith("oe") or "please specify" in question_text.lower():
        return "open_text", None
    if block.options:
        return "values_range", _default_value_range(block)
    return None, None


def _apply_type_b_hint(block: _QuestionBlock, type_line: str) -> None:
    type_text = type_line.split("|")[-1].strip().lower()
    if "multi-select" in type_text or "multi select" in type_text:
        block.type_hint = "values_range"
        block.value_range = (0, 1)
        return
    if "essay" in type_text or "open-end" in type_text or "type-in" in type_text:
        block.type_hint = "open_text"
        block.value_range = None
        return
    if (
        "single-select" in type_text
        or "single select" in type_text
        or "matrix" in type_text
        or "scale" in type_text
        or "rank" in type_text
    ):
        block.type_hint = "values_range"
        block.value_range = None


def _default_value_range(block: _QuestionBlock) -> tuple[int, int]:
    option_count = len(block.options)
    return (1, option_count) if option_count > 0 else (1, 5)


def _type_b_line(text: str) -> str | None:
    return text if TYPE_B_TYPE_PATTERN.search(text) else None


def _annotations(text: str) -> list[str]:
    return BRACKET_ANNOTATION_PATTERN.findall(text)


def _clean_text(text: str) -> str:
    without_brackets = BRACKET_ANNOTATION_PATTERN.sub("", text)
    return _normalise_text(without_brackets)


def _normalise_text(text: str) -> str:
    return " ".join(text.strip().split())


def _canonical_id(raw_id: str) -> str:
    cleaned = raw_id.strip().strip("[]").rstrip(".:")
    return re.sub(r"[\s\-]+", "_", cleaned)


def _derive_parent_canonical_id(canonical_id: str) -> str | None:
    row_oe_match = PARENT_ROW_OE_RE.match(canonical_id)
    if row_oe_match:
        return row_oe_match.group(1)

    oe_match = PARENT_OE_RE.match(canonical_id)
    if oe_match:
        return oe_match.group(1)

    return None


def _data_map(
    path: str,
    paragraphs: list[_ParagraphView],
    questions: list[ParsedQuestion],
    parser_warnings: list[str],
) -> DataMap:
    return {
        "questions": questions,
        "source_path": path,
        "sheet_name": "word_document",
        "total_rows_in_sheet": len(paragraphs),
        "parser_warnings": parser_warnings,
    }
