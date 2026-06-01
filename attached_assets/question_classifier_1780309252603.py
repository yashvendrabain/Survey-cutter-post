"""Question classifier for parsed data maps."""

from __future__ import annotations

from dataclasses import replace
from datetime import datetime, timezone
import re

from src.adapters.label_pattern_subcolumn import (
    apply_label_pattern_matching,
    label_pattern_warnings,
)
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
_PIPE_PATTERN = re.compile(r"\[(pipe|pn)(?::|\s)+([^\]]+)\]", re.IGNORECASE)
_SIBLING_GRID_ID_PATTERN = re.compile(
    r"^(?P<parent>[A-Za-z_]+\d+)r(?P<row>\d+)(?:[a-z]\d+)?$",
    re.IGNORECASE,
)
_QUESTION_TEXT_SEPARATOR_PATTERN = re.compile(r"\s+[-\u2013\u2014]\s+")
_LEADING_QUESTION_PREFIX_PATTERN = re.compile(
    r"^\s*\[?[A-Za-z_]+\d+r\d+(?:[a-z]\d+)?\]?\s*[-:]?\s*",
    re.IGNORECASE,
)
_TRAILING_PUNCTUATION = ".,;:!?- "
GRID_RATED = "GRID_RATED"
GRID_CATEGORICAL = "GRID_CATEGORICAL"
GRID_BINARY_SELECT = "GRID_BINARY_SELECT"
_GRID_SUBTYPES = (GRID_RATED, GRID_CATEGORICAL, GRID_BINARY_SELECT)
_GRID_CONFIDENCE_LOW_THRESHOLD = 0.4
_LABEL_SIGNAL_WEIGHT = 0.4
_VALUE_RANGE_SIGNAL_WEIGHT = 0.2
_SUB_COLUMN_SIGNAL_WEIGHT = 0.2
_REJECTION_SIGNAL_WEIGHT = 0.3
_TEXT_SIGNAL_WEIGHT = 0.1
_REJECTION_PREFIXES = (
    "NO TO: ",
    "NOT: ",
    "No to: ",
    "NOT SELECTED: ",
    "NOT SELECTED - ",
    "NO - ",
    "Not selected: ",
)
_MISSING_VALUE_TOKENS = {
    "i don't know",
    "i don\u2019t know",
    "this was not something i considered",
    "not applicable",
    "n/a",
}


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
        question["canonical_id"]: (
            question.get("conditional_on")
            if "conditional_on" in question
            else _first_pipe_reference(question["question_text"])
        )
        for question in data_map["questions"]
    }

    question_specs = tuple(
        _build_question_spec(
            question,
            tuple(raw_columns),
            raw_column_set,
            all_sub_columns,
            resolved_texts.get(question["canonical_id"], question["question_text"]),
            conditional_refs.get(question["canonical_id"]),
        )
        for question in data_map["questions"]
    )
    question_specs = _merge_grid_siblings(question_specs, raw_column_set)

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
    raw_column_order: tuple[str, ...],
    raw_column_set: set[str],
    all_sub_columns: set[str],
    resolved_question_text: str,
    conditional_on: str | None,
) -> QuestionSpec:
    question = apply_label_pattern_matching(question, raw_column_order)
    question_type = _classify_question(question, all_sub_columns)
    expected_columns = _expected_columns(question)
    if question_type in {
        QuestionType.GRID_SINGLE_SELECT,
        QuestionType.GRID_RATED,
        QuestionType.GRID_BINARY_SELECT,
    }:
        expected_columns = _expand_grid_c_columns(
            question,
            expected_columns,
            raw_column_set,
        )
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
        QuestionType.RANK_ORDER,
        QuestionType.NUMERIC_ALLOCATION,
        QuestionType.GRID_SINGLE_SELECT,
        QuestionType.GRID_RATED,
        QuestionType.GRID_BINARY_SELECT,
    }:
        if not present_columns:
            analysis_eligible = False
            exclusion_reason = "raw column not found in data"

    raw_columns = _raw_columns_for_spec(
        question_type, question["canonical_id"], expected_columns, present_columns
    )
    option_map = _option_map_for_spec(question_type, question)
    grid_row_labels = _grid_row_labels_for_spec(question_type, question, raw_columns)
    grid_column_labels = _grid_column_labels_for_spec(question_type, question)

    if (
        question_type
        in {
            QuestionType.GRID_SINGLE_SELECT,
            QuestionType.GRID_RATED,
            QuestionType.GRID_BINARY_SELECT,
        }
        and 0 < len(present_columns) < len(expected_columns)
    ):
        missing = tuple(column for column in expected_columns if column not in raw_column_set)
        exclusion_reason = "missing grid raw columns: " + ", ".join(missing)

    possible_role: str | None = None
    classification_confidence_low = False
    if question_type in {
        QuestionType.GRID_SINGLE_SELECT,
        QuestionType.GRID_RATED,
        QuestionType.GRID_BINARY_SELECT,
    }:
        possible_role, confidence = classify_grid_subtype(question, raw_column_set)
        classification_confidence_low = confidence < _GRID_CONFIDENCE_LOW_THRESHOLD
    schema_warnings = label_pattern_warnings(question)
    if schema_warnings:
        possible_role = " | ".join(schema_warnings)
        classification_confidence_low = True

    spec = QuestionSpec(
        question_id=question["canonical_id"],
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
        grid_column_labels=grid_column_labels,
        option_other_code=_option_other_code(question),
        conditional_on=conditional_on,
        possible_role=possible_role,
        classification_confidence_low=classification_confidence_low,
        label_to_numeric_value=dict(question.get("label_to_numeric_value", {})),
        na_label_set=frozenset(question.get("na_label_set", frozenset())),
        allowed_numeric_range=question.get("allowed_numeric_range"),
    )
    return replace(
        spec,
        is_demographic=_is_demographic_question(spec, raw_column_set),
    )


def _merge_grid_siblings(
    question_specs: tuple[QuestionSpec, ...],
    raw_column_set: set[str] | None = None,
) -> tuple[QuestionSpec, ...]:
    """Merge row-sibling specs such as Q26r1..Q26r21 into one grid parent."""

    raw_column_set = raw_column_set or set()
    parent_groups: dict[str, list[QuestionSpec]] = {}
    for spec in question_specs:
        parent_id = _sibling_parent_id(spec.canonical_id)
        if parent_id is None:
            continue
        if spec.question_type not in {
            QuestionType.SINGLE_SELECT,
            QuestionType.GRID_SINGLE_SELECT,
            QuestionType.MULTI_SELECT_BINARY,
            QuestionType.RANK_ORDER,
            QuestionType.GRID_RATED,
            QuestionType.GRID_BINARY_SELECT,
            QuestionType.DIRECT_NUMERIC,
        }:
            continue
        parent_groups.setdefault(parent_id, []).append(spec)

    existing_parent_ids = {
        spec.canonical_id
        for spec in question_specs
        if _sibling_parent_id(spec.canonical_id) is None
    }
    merged_by_first_child: dict[str, QuestionSpec] = {}
    merged_child_ids: set[str] = set()
    for parent_id, members in parent_groups.items():
        if parent_id in existing_parent_ids or len(members) < 2:
            continue
        merged = _merge_sibling_group(parent_id, members, raw_column_set)
        if merged is None:
            continue
        ordered_members = _sort_sibling_members(members)
        merged_by_first_child[ordered_members[0].canonical_id] = merged
        merged_child_ids.update(member.canonical_id for member in ordered_members)

    if not merged_child_ids:
        return question_specs

    output: list[QuestionSpec] = []
    for spec in question_specs:
        if spec.canonical_id in merged_by_first_child:
            output.append(merged_by_first_child[spec.canonical_id])
        elif spec.canonical_id in merged_child_ids:
            continue
        else:
            output.append(spec)
    return tuple(output)


def _merge_sibling_group(
    parent_id: str,
    members: list[QuestionSpec],
    raw_column_set: set[str],
) -> QuestionSpec | None:
    ordered_members = _sort_sibling_members(members)
    if not _sibling_group_is_mergeable(ordered_members):
        return None

    option_map = dict(ordered_members[0].option_map)
    value_range = ordered_members[0].value_range
    possible_role = _sibling_grid_role(ordered_members)
    question_type = _merged_sibling_question_type(ordered_members, possible_role)

    raw_columns: list[str] = []
    grid_row_labels: dict[str, str] = {}
    grid_column_labels: dict[str, str] = {}
    for member in ordered_members:
        criterion_label = _sibling_criterion_label(member)
        member_columns = _expanded_member_raw_columns(member, raw_column_set)
        for column in member_columns:
            raw_columns.append(column)
            grid_row_labels[column] = criterion_label
            _row_id, group_key = _grid_row_and_group_ids(column)
            if group_key is not None and group_key in member.grid_column_labels:
                grid_column_labels.setdefault(group_key, member.grid_column_labels[group_key])
    if question_type in {QuestionType.MULTI_SELECT_BINARY, QuestionType.RANK_ORDER}:
        option_map = {column: grid_row_labels[column] for column in raw_columns}

    label_to_numeric_value: dict[str, float] = {}
    na_label_set: set[str] = set()
    for member in ordered_members:
        label_to_numeric_value.update(member.label_to_numeric_value)
        na_label_set.update(member.na_label_set)
    allowed_numeric_range = None
    if label_to_numeric_value:
        numeric_values = tuple(label_to_numeric_value.values())
        allowed_numeric_range = (float(min(numeric_values)), float(max(numeric_values)))

    conditional_values = {member.conditional_on for member in ordered_members}
    conditional_on = conditional_values.pop() if len(conditional_values) == 1 else None
    analysis_eligible = any(member.analysis_eligible for member in ordered_members) or any(
        column in raw_column_set for column in raw_columns
    )
    root_text = _sibling_question_root_text(ordered_members)

    merged_spec = QuestionSpec(
        question_id=parent_id,
        canonical_id=parent_id,
        question_text=root_text,
        question_type=question_type,
        raw_columns=tuple(raw_columns),
        option_map=option_map,
        value_range=value_range,
        denominator_policy=ordered_members[0].denominator_policy,
        theme_tags=ordered_members[0].theme_tags,
        possible_role=possible_role,
        analysis_eligible=analysis_eligible,
        exclusion_reason=None if analysis_eligible else "raw column not found in data",
        parent_question_id=None,
        grid_row_labels=grid_row_labels,
        grid_column_labels=grid_column_labels,
        option_other_code=ordered_members[0].option_other_code,
        is_demographic=False,
        conditional_on=conditional_on,
        label_to_numeric_value=label_to_numeric_value,
        na_label_set=frozenset(na_label_set),
        allowed_numeric_range=allowed_numeric_range,
    )
    merged_role, confidence = classify_grid_subtype(merged_spec, raw_column_set)
    role_to_type = {
        GRID_RATED: QuestionType.GRID_RATED,
        GRID_BINARY_SELECT: QuestionType.GRID_BINARY_SELECT,
        GRID_CATEGORICAL: QuestionType.GRID_SINGLE_SELECT,
    }
    grid_family_types = {
        QuestionType.GRID_SINGLE_SELECT,
        QuestionType.GRID_RATED,
        QuestionType.GRID_BINARY_SELECT,
    }
    final_type = merged_spec.question_type
    if final_type in grid_family_types and any(
        member.question_type in grid_family_types for member in ordered_members
    ):
        final_type = role_to_type.get(merged_role, final_type)
    final_option_map = merged_spec.option_map
    if final_type is QuestionType.GRID_RATED and _option_map_looks_numeric_scale(
        final_option_map
    ):
        final_option_map = {}
    return replace(
        merged_spec,
        question_type=final_type,
        option_map=final_option_map,
        possible_role=merged_role,
        classification_confidence_low=confidence < _GRID_CONFIDENCE_LOW_THRESHOLD,
    )


def _sibling_group_is_mergeable(members: list[QuestionSpec]) -> bool:
    first = members[0]
    first_role = _sibling_grid_role([first])

    for member in members:
        if member.question_type is not first.question_type:
            return False
        if member.option_map != first.option_map:
            return False
        if member.value_range != first.value_range:
            return False
        if _sibling_grid_role([member]) != first_role:
            return False
    if first_role == GRID_BINARY_SELECT and not _members_have_c_column_groups(members):
        return False
    return True


def _merged_sibling_question_type(
    members: list[QuestionSpec],
    possible_role: str,
) -> QuestionType:
    if _members_have_c_column_groups(members):
        return QuestionType.GRID_SINGLE_SELECT
    value_range = members[0].value_range
    if (
        value_range is not None
        and value_range[0] == 1
        and 2 <= value_range[1] <= 10
        and not members[0].option_map
    ):
        return QuestionType.RANK_ORDER
    if value_range == (0, 1):
        return QuestionType.MULTI_SELECT_BINARY
    return QuestionType.GRID_SINGLE_SELECT


def _sibling_grid_role(members: list[QuestionSpec]) -> str:
    explicit_roles = {
        member.possible_role for member in members if member.possible_role is not None
    }
    if len(explicit_roles) == 1:
        role = explicit_roles.pop() or GRID_CATEGORICAL
        return role
    labels = [str(label) for label in members[0].option_map.values()]
    return _grid_subtype_from_parts(members[0].value_range, labels)


def _option_map_looks_numeric_scale(option_map: dict[int | str, str]) -> bool:
    labels = [str(label) for label in option_map.values()]
    return _grid_subtype_from_parts(None, labels) == GRID_RATED


def _members_have_c_column_groups(members: list[QuestionSpec]) -> bool:
    for member in members:
        for column in member.raw_columns:
            if re.match(rf"^{re.escape(member.canonical_id)}c\d+$", column):
                return True
            if re.match(r"^.+r\d+c\d+$", column):
                return True
    return False


def _expanded_member_raw_columns(
    member: QuestionSpec,
    raw_column_set: set[str],
) -> tuple[str, ...]:
    c_columns = _matching_grid_c_columns(member.canonical_id, raw_column_set)
    if len(c_columns) >= 2:
        return tuple(c_columns)
    present = tuple(column for column in member.raw_columns if column in raw_column_set)
    return present or member.raw_columns


def _sibling_parent_id(canonical_id: str) -> str | None:
    match = _SIBLING_GRID_ID_PATTERN.match(canonical_id)
    if match is None:
        return None
    return match.group("parent")


def _sort_sibling_members(members: list[QuestionSpec]) -> list[QuestionSpec]:
    return sorted(members, key=lambda member: _sibling_sort_key(member.canonical_id))


def _sibling_sort_key(canonical_id: str) -> tuple[int, str]:
    match = _SIBLING_GRID_ID_PATTERN.match(canonical_id)
    if match is None:
        return (10**9, canonical_id)
    return (int(match.group("row")), canonical_id)


def _split_sibling_question_text(question_text: str) -> tuple[str, str] | None:
    parts = _QUESTION_TEXT_SEPARATOR_PATTERN.split(question_text, maxsplit=1)
    if len(parts) != 2:
        return None
    criterion_label = _LEADING_QUESTION_PREFIX_PATTERN.sub("", parts[0]).strip()
    root_text = parts[1].strip()
    if not criterion_label or not root_text:
        return None
    return criterion_label, root_text


def _sibling_question_root_text(members: list[QuestionSpec]) -> str:
    roots: dict[str, str] = {}
    for member in members:
        split_text = _split_sibling_question_text(member.question_text)
        root_text = split_text[1] if split_text is not None else member.question_text
        roots.setdefault(_normalise_sibling_root_text(root_text), root_text.strip())
    if not roots:
        return members[0].question_text
    _root_key, root_text = max(
        roots.items(),
        key=lambda item: sum(
            1
            for member in members
            if _normalise_sibling_root_text(
                (
                    _split_sibling_question_text(member.question_text)[1]
                    if _split_sibling_question_text(member.question_text) is not None
                    else member.question_text
                )
            )
            == item[0]
        ),
    )
    return root_text


def _sibling_criterion_label(member: QuestionSpec) -> str:
    split_text = _split_sibling_question_text(member.question_text)
    if split_text is not None:
        return split_text[0]
    return member.canonical_id


def _normalise_sibling_root_text(question_text: str) -> str:
    return re.sub(r"\s+", " ", question_text).strip().casefold()


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
        if has_sub_columns and _has_grid_c_columns(question):
            subtype, _confidence = classify_grid_subtype(question, set())
            if subtype == GRID_RATED:
                return QuestionType.GRID_RATED
            if subtype == GRID_BINARY_SELECT or value_range == (0, 1):
                return QuestionType.GRID_BINARY_SELECT
            return QuestionType.GRID_SINGLE_SELECT
        if has_sub_columns and has_options:
            return QuestionType.GRID_SINGLE_SELECT
        if has_sub_columns and not has_options:
            if value_range is not None and value_range[0] == 1 and 2 <= value_range[1] <= 10:
                return QuestionType.RANK_ORDER
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
    child_columns = _child_columns(question)
    if child_columns:
        return child_columns
    if question["sub_columns"]:
        return tuple(sub_column_id for sub_column_id, _ in question["sub_columns"])
    return (question["canonical_id"],)


def _child_columns(question: ParsedQuestion) -> tuple[str, ...]:
    children = question.get("children", [])
    columns: list[str] = []
    for child in children:
        if child["sub_columns"]:
            columns.extend(sub_column_id for sub_column_id, _label in child["sub_columns"])
        else:
            columns.append(child["canonical_id"])
    return tuple(columns)


def _has_grid_c_columns(question: ParsedQuestion) -> bool:
    return any(re.match(r"^.+r\d+c\d+$", column) for column in _expected_columns(question))


def _expand_grid_c_columns(
    question: ParsedQuestion,
    expected_columns: tuple[str, ...],
    raw_column_set: set[str],
) -> tuple[str, ...]:
    if not question["sub_columns"]:
        return expected_columns

    expanded: list[str] = []
    for base_column in expected_columns:
        c_columns = _matching_grid_c_columns(base_column, raw_column_set)
        if len(c_columns) >= 2:
            expanded.extend(c_columns)
        else:
            expanded.append(base_column)
    return tuple(expanded)


def _matching_grid_c_columns(
    base_column: str,
    raw_column_set: set[str],
) -> list[str]:
    pattern = re.compile(rf"^{re.escape(base_column)}c(?P<group>\d+)$")
    matches = []
    for raw_column in raw_column_set:
        match = pattern.match(raw_column)
        if match is not None:
            matches.append((int(match.group("group")), raw_column))
    matches.sort(key=lambda item: (item[0], item[1]))
    return [raw_column for _group, raw_column in matches]


def _raw_columns_for_spec(
    question_type: QuestionType,
    canonical_id: str,
    expected_columns: tuple[str, ...],
    present_columns: tuple[str, ...],
) -> tuple[str, ...]:
    if question_type in {
        QuestionType.MULTI_SELECT_BINARY,
        QuestionType.RANK_ORDER,
        QuestionType.GRID_RATED,
        QuestionType.GRID_BINARY_SELECT,
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
    if question_type in {
        QuestionType.GRID_RATED,
        QuestionType.GRID_BINARY_SELECT,
    }:
        if question["options"]:
            return {code: label for code, label in question["options"]}
        return _grid_group_labels_for_spec(question)
    if question_type is QuestionType.RANK_ORDER:
        return {sub_column_id: label for sub_column_id, label in question["sub_columns"]}
    return {}


def _grid_row_labels_for_spec(
    question_type: QuestionType,
    question: ParsedQuestion,
    raw_columns: tuple[str, ...],
) -> dict[str, str] | None:
    if question_type not in {
        QuestionType.GRID_SINGLE_SELECT,
        QuestionType.GRID_RATED,
        QuestionType.GRID_BINARY_SELECT,
    }:
        return None

    child_row_labels = _child_grid_row_labels(question)
    if child_row_labels:
        labels = {}
        for sub_column_id in raw_columns:
            if sub_column_id in child_row_labels:
                labels[sub_column_id] = child_row_labels[sub_column_id]
                continue
            base_column_id = _grid_base_sub_column_id(sub_column_id)
            if base_column_id in child_row_labels:
                labels[sub_column_id] = child_row_labels[base_column_id]
        if labels:
            return labels

    row_label_lookup = {
        sub_column_id: label for sub_column_id, label in question["sub_columns"]
    }
    labels: dict[str, str] = {}
    for sub_column_id in raw_columns:
        if sub_column_id in row_label_lookup:
            labels[sub_column_id] = row_label_lookup[sub_column_id]
            continue
        base_column_id = _grid_base_sub_column_id(sub_column_id)
        if base_column_id in row_label_lookup:
            labels[sub_column_id] = row_label_lookup[base_column_id]
    return labels


def _grid_column_labels_for_spec(
    question_type: QuestionType,
    question: ParsedQuestion,
) -> dict[str, str]:
    if question_type not in {
        QuestionType.GRID_SINGLE_SELECT,
        QuestionType.GRID_RATED,
        QuestionType.GRID_BINARY_SELECT,
    }:
        return {}
    return {
        str(group_key): str(label)
        for group_key, label in _grid_group_labels_for_spec(question).items()
    }


def _child_grid_row_labels(question: ParsedQuestion) -> dict[str, str]:
    labels: dict[str, str] = {}
    for child in question.get("children", []):
        criterion_label = _child_criterion_label(child)
        if child["sub_columns"]:
            for sub_column_id, _sub_label in child["sub_columns"]:
                labels[sub_column_id] = criterion_label
        else:
            labels[child["canonical_id"]] = criterion_label
    return labels


def _child_criterion_label(question: ParsedQuestion) -> str:
    parts = _QUESTION_TEXT_SEPARATOR_PATTERN.split(question["question_text"], maxsplit=1)
    if len(parts) == 2 and parts[0].strip():
        return _LEADING_QUESTION_PREFIX_PATTERN.sub("", parts[0]).strip()
    return question["canonical_id"]


def _grid_group_labels_for_spec(question: ParsedQuestion) -> dict[int | str, str]:
    labels: dict[int | str, str] = {}
    for sub_column_id, label in question["sub_columns"]:
        match = re.match(r"^.+r\d+c(?P<group>\d+)$", sub_column_id)
        if match is None:
            continue
        group_key = match.group("group")
        labels.setdefault(group_key, label)
    return labels


def _grid_row_and_group_ids(column: str) -> tuple[str, str | None]:
    match = re.match(r"^(?P<row>.+?r\d+)c(?P<group>\d+)$", str(column))
    if match is None:
        return str(column), None
    return match.group("row"), match.group("group")


def _grid_base_sub_column_id(sub_column_id: str) -> str:
    match = re.match(r"^(.+r\d+)c\d+$", sub_column_id)
    if match is None:
        return sub_column_id
    return match.group(1)


def _grid_subtype_for_question(question: ParsedQuestion) -> str:
    """Best-effort rendering subtype for grid-style questions.

    The classifier only sees the data map, not the raw dataframe. It therefore
    uses the strongest stable signals available here: bounded numeric scale
    labels imply rated grids; explicit selected/unselected language implies
    binary grids; everything else is a categorical grid.
    """

    subtype, _confidence = classify_grid_subtype(question, set())
    return subtype


def classify_grid_subtype(
    question: ParsedQuestion | QuestionSpec,
    raw_column_set: set[str] | None = None,
) -> tuple[str, float]:
    """Classify a grid subtype using several weak and strong signals."""

    labels = _grid_option_labels(question)
    value_range = _grid_value_range(question)
    raw_columns = _grid_signal_columns(question, raw_column_set or set())
    question_text = _grid_question_text(question)
    scores = {subtype: 0.0 for subtype in _GRID_SUBTYPES}
    max_possible = 0.0

    if labels:
        max_possible += _LABEL_SIGNAL_WEIGHT
        numeric_values = [
            value for label in labels if (value := _numeric_label_value(label)) is not None
        ]
        numeric_ratio = len(numeric_values) / len(labels)
        if _labels_have_rejection_prefix(labels):
            scores[GRID_BINARY_SELECT] += _LABEL_SIGNAL_WEIGHT
        elif _labels_are_binary_select(labels):
            scores[GRID_BINARY_SELECT] += _LABEL_SIGNAL_WEIGHT
        elif numeric_ratio >= 0.8 and _numeric_values_form_rating_scale(numeric_values):
            scores[GRID_RATED] += _LABEL_SIGNAL_WEIGHT
        elif numeric_ratio >= 0.8 and _numeric_values_are_binary(numeric_values):
            scores[GRID_BINARY_SELECT] += _LABEL_SIGNAL_WEIGHT
        elif numeric_ratio < 0.5:
            scores[GRID_CATEGORICAL] += _LABEL_SIGNAL_WEIGHT

    if value_range is not None:
        max_possible += _VALUE_RANGE_SIGNAL_WEIGHT
        low, high = value_range
        if (low, high) == (0, 1) and len(labels) <= 2:
            scores[GRID_BINARY_SELECT] += _VALUE_RANGE_SIGNAL_WEIGHT
        elif (low, high) == (0, 1) and len(labels) > 2:
            scores[GRID_CATEGORICAL] += _VALUE_RANGE_SIGNAL_WEIGHT
        elif labels:
            numeric_values = [
                value for label in labels if (value := _numeric_label_value(label)) is not None
            ]
            if _numeric_values_form_rating_scale(numeric_values):
                scores[GRID_RATED] += _VALUE_RANGE_SIGNAL_WEIGHT

    if raw_columns:
        max_possible += _SUB_COLUMN_SIGNAL_WEIGHT
        has_rc = any(re.match(r"^.+r\d+c\d+$", column) for column in raw_columns)
        has_r = any(re.match(r"^.+r\d+$", column) for column in raw_columns)
        has_c = any(re.match(r"^.+c\d+$", column) for column in raw_columns)
        if has_rc:
            scores[GRID_RATED] += _SUB_COLUMN_SIGNAL_WEIGHT * 0.35
            scores[GRID_CATEGORICAL] += _SUB_COLUMN_SIGNAL_WEIGHT * 0.65
        elif has_r:
            scores[GRID_CATEGORICAL] += _SUB_COLUMN_SIGNAL_WEIGHT * 0.5
        elif has_c:
            scores[GRID_CATEGORICAL] += _SUB_COLUMN_SIGNAL_WEIGHT * 0.5

    if _labels_have_rejection_prefix(labels):
        max_possible += _REJECTION_SIGNAL_WEIGHT
        scores[GRID_BINARY_SELECT] += _REJECTION_SIGNAL_WEIGHT

    text_lower = question_text.lower()
    if any(token in text_lower for token in ("rate", "score", "rating", "1-10", "scale of")):
        max_possible += _TEXT_SIGNAL_WEIGHT
        scores[GRID_RATED] += _TEXT_SIGNAL_WEIGHT
    if any(token in text_lower for token in ("select all that apply", "which of the following")):
        max_possible += _TEXT_SIGNAL_WEIGHT
        scores[GRID_BINARY_SELECT] += _TEXT_SIGNAL_WEIGHT
    if any(token in text_lower for token in ("role", "category", "type", "stakeholder")):
        max_possible += _TEXT_SIGNAL_WEIGHT
        scores[GRID_CATEGORICAL] += _TEXT_SIGNAL_WEIGHT

    ordered_scores = sorted(scores.items(), key=lambda item: item[1], reverse=True)
    winning_subtype, winning_score = ordered_scores[0]
    second_score = ordered_scores[1][1]
    if max_possible <= 0:
        return GRID_CATEGORICAL, 0.0
    confidence = max(0.0, min(1.0, (winning_score - second_score) / max_possible))
    return winning_subtype, confidence


class _GridSubtypeParts:
    def __init__(self, value_range: tuple[int, int] | None, labels: list[str]) -> None:
        self.value_range = value_range
        self.labels = labels
        self.question_text = ""
        self.raw_columns: tuple[str, ...] = tuple()


def _grid_option_labels(question: ParsedQuestion | QuestionSpec | _GridSubtypeParts) -> list[str]:
    if isinstance(question, dict):
        labels = [str(label) for _code, label in question.get("options", [])]
        if labels:
            return labels
        for child in question.get("children", []) or []:
            for _code, label in child.get("options", []) or []:
                labels.append(str(label))
        return labels
    if isinstance(question, _GridSubtypeParts):
        return [str(label) for label in question.labels]
    return [str(label) for label in question.option_map.values()]


def _grid_value_range(question: ParsedQuestion | QuestionSpec | _GridSubtypeParts) -> tuple[int, int] | None:
    if isinstance(question, dict):
        return question["value_range"]
    return question.value_range


def _grid_question_text(question: ParsedQuestion | QuestionSpec | _GridSubtypeParts) -> str:
    if isinstance(question, dict):
        return str(question.get("question_text", ""))
    return str(question.question_text)


def _grid_signal_columns(
    question: ParsedQuestion | QuestionSpec | _GridSubtypeParts,
    raw_column_set: set[str],
) -> tuple[str, ...]:
    if isinstance(question, dict):
        expected_columns = _expected_columns(question)
        if raw_column_set:
            expected_columns = _expand_grid_c_columns(
                question,
                expected_columns,
                raw_column_set,
            )
        return expected_columns
    return tuple(getattr(question, "raw_columns", tuple()))


def _labels_have_rejection_prefix(labels: list[str]) -> bool:
    return any(
        str(label).strip().lower().startswith(prefix.strip().lower())
        for label in labels
        for prefix in _REJECTION_PREFIXES
    )


def _labels_are_binary_select(labels: list[str]) -> bool:
    label_set = {str(label).strip().lower() for label in labels if str(label).strip()}
    return bool(
        label_set
        and label_set <= {
            "selected",
            "not selected",
            "unselected",
            "checked",
            "unchecked",
            "yes",
            "no",
            "true",
            "false",
            "1",
            "0",
        }
        and label_set
        & {"selected", "checked", "yes", "true", "1"}
    )


def _numeric_values_are_binary(values: list[float]) -> bool:
    return bool(values) and set(values) <= {0.0, 1.0}


def _numeric_values_form_rating_scale(values: list[float]) -> bool:
    if not values:
        return False
    scale_min = min(values)
    scale_max = max(values)
    return scale_min >= 0 and scale_max <= 10 and scale_max - scale_min >= 2


def _grid_subtype_from_parts(
    value_range: tuple[int, int] | None,
    labels: list[str],
) -> str:
    subtype, _confidence = classify_grid_subtype(
        _GridSubtypeParts(value_range=value_range, labels=labels),
        set(),
    )
    return subtype


def _grid_options_look_binary_select(
    value_range: tuple[int, int] | None,
    labels: list[str],
) -> bool:
    lowered = " ".join(label.strip().lower() for label in labels)
    if any(prefix.strip().lower() in lowered for prefix in _REJECTION_PREFIXES):
        return True
    if any(token in lowered for token in ("checked", "unchecked")):
        return True
    if value_range == (0, 1):
        label_set = {label.strip().lower() for label in labels}
        return bool(
            label_set
            and label_set <= {
                "selected",
                "not selected",
                "unselected",
                "yes",
                "no",
                "true",
                "false",
                "1",
                "0",
            }
            and any(label in label_set for label in {"selected", "yes", "true", "1"})
        )
    return False


def _grid_options_look_rated(
    value_range: tuple[int, int] | None,
    labels: list[str],
) -> bool:
    if not labels:
        return False
    numeric_values = [
        value
        for label in labels
        if (value := _numeric_label_value(label)) is not None
    ]
    if len(numeric_values) / len(labels) < 0.8:
        return False
    scale_min = min(numeric_values)
    scale_max = max(numeric_values)
    return scale_min >= 0 and scale_max <= 10 and scale_max - scale_min >= 2


def _is_numeric_label(label: str) -> bool:
    return _numeric_label_value(label) is not None


def _extract_leading_int(value: str | int | float | None) -> int | None:
    """Extract a leading integer from labels such as '10 (extremely high)'."""

    numeric = _coerce_to_numeric(value)
    if numeric is None:
        return None
    if not float(numeric).is_integer():
        return None
    return int(numeric)


def _coerce_to_numeric(value: object) -> float | None:
    """Coerce numeric-coded survey values while skipping missing-value tokens."""

    if value is None:
        return None
    text = str(value).strip()
    if not text or text.casefold() in _MISSING_VALUE_TOKENS:
        return None
    match = re.match(r"^\s*(-?\d+(?:\.\d+)?)", text)
    if match is None:
        return None
    try:
        return float(match.group(1))
    except (TypeError, ValueError):
        return None


def _has_no_rank_repeats(
    df: object,
    sibling_cols: list[str],
    K: int,
    sample_rows: int = 500,
) -> bool:
    """Return True when no respondent repeats a rank across sibling columns."""

    if not sibling_cols or not hasattr(df, "loc"):
        return False
    sample = df.loc[:, [col for col in sibling_cols if col in df.columns]].head(sample_rows)
    for _idx, row in sample.iterrows():
        ranks = [
            rank
            for value in row
            if (rank := _extract_leading_int(value)) is not None and 1 <= rank <= K
        ]
        if len(ranks) != len(set(ranks)):
            return False
    return True


def _labels_look_numeric(labels: list[str]) -> tuple[bool, int]:
    """Return whether labels mostly form a numeric rating scale and its max."""

    values = [
        value for label in labels if (value := _extract_leading_int(label)) is not None
    ]
    if not labels or len(values) / len(labels) < 0.6:
        return False, 0
    max_value = max(values)
    return max_value <= 10 and max(values) - min(values) >= 2, max_value


def _numeric_label_value(label: str) -> float | None:
    return _coerce_to_numeric(label)


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
