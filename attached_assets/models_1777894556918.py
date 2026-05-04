"""Contract layer for the Survey Insight Engine.

All inter-module data crosses through these enums and dataclasses. This module
intentionally contains no file parsing, I/O, or numerical computation.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
import math
from numbers import Integral, Real
from typing import Any


class QuestionType(str, Enum):
    """Survey question classifications used by downstream analysis modules."""

    SINGLE_SELECT = "SINGLE_SELECT"
    MULTI_SELECT_BINARY = "MULTI_SELECT_BINARY"
    GRID_SINGLE_SELECT = "GRID_SINGLE_SELECT"
    NUMERIC_ALLOCATION = "NUMERIC_ALLOCATION"
    DIRECT_NUMERIC = "DIRECT_NUMERIC"
    OPEN_TEXT = "OPEN_TEXT"
    DEMOGRAPHIC_OR_SEGMENT = "DEMOGRAPHIC_OR_SEGMENT"
    METADATA_OR_ID = "METADATA_OR_ID"
    UNKNOWN = "UNKNOWN"


class DenominatorPolicy(str, Enum):
    """Rules for choosing the denominator used in metric calculations."""

    VALID_RESPONSES = "VALID_RESPONSES"
    ALL_RESPONDENTS = "ALL_RESPONDENTS"
    EXPOSED_TO_QUESTION = "EXPOSED_TO_QUESTION"


class AnalysisType(str, Enum):
    """Cross-cut analysis categories populated in later stages."""

    CROSS_TAB = "CROSS_TAB"
    SEGMENT_PROFILE = "SEGMENT_PROFILE"
    GROUP_COMPARISON = "GROUP_COMPARISON"
    EXPECTED_VS_REALIZED = "EXPECTED_VS_REALIZED"
    MULTI_QUESTION_METRIC = "MULTI_QUESTION_METRIC"


def _is_blank(value: str | None) -> bool:
    return value is None or not value.strip()


def _require_non_empty_string(value: str, field_name: str) -> None:
    if not isinstance(value, str) or _is_blank(value):
        raise ValueError(f"{field_name} must be a non-empty string")


def _require_timezone_aware(value: datetime, field_name: str) -> None:
    if value.tzinfo is None or value.utcoffset() is None:
        raise ValueError(f"{field_name} must be timezone-aware")


def _require_non_negative_int(value: Any, field_name: str) -> None:
    if isinstance(value, bool) or not isinstance(value, Integral):
        raise ValueError(f"{field_name} must be a non-negative integer")
    if int(value) < 0:
        raise ValueError(f"{field_name} must be a non-negative integer")


def _require_rate(value: Any, field_name: str) -> None:
    if not isinstance(value, Real) or isinstance(value, bool):
        raise ValueError(f"{field_name} must be numeric")
    if not 0.0 <= float(value) <= 1.0:
        raise ValueError(f"{field_name} must be between 0 and 1")


def _require_numeric(value: Any, field_name: str) -> None:
    if not isinstance(value, Real) or isinstance(value, bool):
        raise ValueError(f"{field_name} must be numeric")


@dataclass(frozen=True, slots=True)
class QuestionSpec:
    """Contract for one logical survey question from the data map."""

    question_id: str
    canonical_id: str
    question_text: str
    question_type: QuestionType
    raw_columns: tuple[str, ...]
    option_map: dict[int | str, str]
    value_range: tuple[int, int] | None = None
    denominator_policy: DenominatorPolicy = field(
        default=DenominatorPolicy.VALID_RESPONSES
    )
    theme_tags: tuple[str, ...] = field(default_factory=tuple)
    possible_role: str | None = None
    analysis_eligible: bool = True
    exclusion_reason: str | None = None
    parent_question_id: str | None = None
    grid_row_labels: dict[str, str] | None = None
    option_other_code: int | str | None = None

    def __post_init__(self) -> None:
        _require_non_empty_string(self.question_id, "question_id")
        _require_non_empty_string(self.canonical_id, "canonical_id")

        if (
            self.question_type not in {QuestionType.METADATA_OR_ID, QuestionType.UNKNOWN}
            and not self.raw_columns
        ):
            raise ValueError(
                "raw_columns must be non-empty unless question_type is "
                "METADATA_OR_ID or UNKNOWN"
            )

        if not self.analysis_eligible and _is_blank(self.exclusion_reason):
            raise ValueError(
                "exclusion_reason must be non-empty when analysis_eligible is False"
            )

        if self.question_type is QuestionType.GRID_SINGLE_SELECT:
            if not self.grid_row_labels:
                raise ValueError(
                    "grid_row_labels must be non-empty for GRID_SINGLE_SELECT"
                )
            if len(self.grid_row_labels) != len(self.raw_columns):
                raise ValueError(
                    "grid_row_labels length must match raw_columns length for "
                    "GRID_SINGLE_SELECT"
                )

        if self.question_type is QuestionType.MULTI_SELECT_BINARY:
            if any(not isinstance(key, str) or _is_blank(key) for key in self.option_map):
                raise ValueError(
                    "option_map keys must be non-empty sub-column ids for "
                    "MULTI_SELECT_BINARY"
                )

        if self.value_range is not None:
            if len(self.value_range) != 2:
                raise ValueError("value_range must contain exactly two values")
            if self.value_range[0] > self.value_range[1]:
                raise ValueError("value_range lower bound must be <= upper bound")


@dataclass(frozen=True, slots=True)
class SurveySchema:
    """Full parsed survey schema shared by parser and analysis modules."""

    questions: tuple[QuestionSpec, ...]
    respondent_id_column: str
    total_respondents: int
    source_datamap_path: str
    source_rawdata_path: str
    parsed_at: datetime

    def __post_init__(self) -> None:
        _require_non_empty_string(self.respondent_id_column, "respondent_id_column")
        _require_non_empty_string(self.source_datamap_path, "source_datamap_path")
        _require_non_empty_string(self.source_rawdata_path, "source_rawdata_path")

        if self.total_respondents <= 0:
            raise ValueError("total_respondents must be greater than 0")

        canonical_ids = [question.canonical_id for question in self.questions]
        if len(canonical_ids) != len(set(canonical_ids)):
            raise ValueError("question canonical_ids must be unique")

        _require_timezone_aware(self.parsed_at, "parsed_at")

    def get_question(self, canonical_id: str) -> QuestionSpec | None:
        for question in self.questions:
            if question.canonical_id == canonical_id:
                return question
        return None

    def analysis_eligible_questions(self) -> tuple[QuestionSpec, ...]:
        return tuple(question for question in self.questions if question.analysis_eligible)


@dataclass(frozen=True, slots=True)
class AuditRecord:
    """Audit trail for one deterministic computed metric."""

    output_sheet: str
    metric_name: str
    source_question_id: str
    source_columns: tuple[str, ...]
    filter_expr: str | None
    numerator: float | int | None
    denominator: float | int | None
    formula: str
    value_raw: float
    valid_n: int
    missing_n: int
    timestamp: datetime

    def __post_init__(self) -> None:
        _require_non_empty_string(self.output_sheet, "output_sheet")
        _require_non_empty_string(self.metric_name, "metric_name")
        _require_non_empty_string(self.source_question_id, "source_question_id")

        if not self.source_columns:
            raise ValueError("source_columns must be non-empty")
        if _is_blank(self.formula):
            raise ValueError("formula must be non-empty")

        _require_numeric(self.value_raw, "value_raw")
        _require_non_negative_int(self.valid_n, "valid_n")
        _require_non_negative_int(self.missing_n, "missing_n")
        _require_timezone_aware(self.timestamp, "timestamp")


@dataclass(frozen=True, slots=True)
class SingleCutResult:
    """Base contract for deterministic single-question analysis output."""

    question_id: str
    question_type: QuestionType
    valid_n: int
    missing_n: int
    denominator_policy: DenominatorPolicy
    warnings: tuple[str, ...] = field(default_factory=tuple, kw_only=True)
    audit_records: tuple[AuditRecord, ...] = field(default_factory=tuple, kw_only=True)

    def __post_init__(self) -> None:
        _require_non_empty_string(self.question_id, "question_id")
        _require_non_negative_int(self.valid_n, "valid_n")
        _require_non_negative_int(self.missing_n, "missing_n")

        if not isinstance(self.question_type, QuestionType):
            raise ValueError("question_type must be a QuestionType")
        if not isinstance(self.denominator_policy, DenominatorPolicy):
            raise ValueError("denominator_policy must be a DenominatorPolicy")


@dataclass(frozen=True, slots=True)
class SingleSelectResult(SingleCutResult):
    """Distribution result for single-response categorical questions."""

    distribution: dict[int | str, dict]

    def __post_init__(self) -> None:
        super(SingleSelectResult, self).__post_init__()
        if not self.distribution and self.valid_n > 0:
            raise ValueError("distribution must be non-empty when valid_n > 0")

        for option_code, payload in self.distribution.items():
            if not isinstance(option_code, (int, str)) or isinstance(option_code, bool):
                raise ValueError("distribution option codes must be int or str")
            _require_distribution_payload(payload, "rate", "distribution")


@dataclass(frozen=True, slots=True)
class MultiSelectResult(SingleCutResult):
    """Selection result for binary multi-select sub-columns.

    Under DenominatorPolicy.VALID_RESPONSES, the denominator for each
    selection_rate is respondents_who_answered_any — i.e. respondents
    who answered at least one sub-column of the multi-select group.
    """

    selections: dict[str, dict]
    respondents_who_answered_any: int

    def __post_init__(self) -> None:
        super(MultiSelectResult, self).__post_init__()
        if self.question_type is not QuestionType.MULTI_SELECT_BINARY:
            raise ValueError("MultiSelectResult requires question_type MULTI_SELECT_BINARY")
        if not self.selections:
            raise ValueError("selections must be non-empty")

        _require_non_negative_int(
            self.respondents_who_answered_any, "respondents_who_answered_any"
        )
        for column_id, payload in self.selections.items():
            _require_non_empty_string(column_id, "selections key")
            if (
                self.valid_n == 0
                and isinstance(payload, dict)
                and isinstance(payload.get("selection_rate"), Real)
                and math.isnan(float(payload["selection_rate"]))
            ):
                _require_distribution_payload_without_rate(
                    payload, "selection_rate", "selections"
                )
            else:
                _require_distribution_payload(payload, "selection_rate", "selections")


@dataclass(frozen=True, slots=True)
class NumericResult(SingleCutResult):
    """Descriptive statistics result for numeric survey responses."""

    mean: float
    median: float
    std: float
    min_val: float
    max_val: float
    percentiles: dict[int, float]
    allocation_target: float | None = None
    allocation_tolerance: float | None = None
    allocation_excluded_n: int | None = None
    per_option_stats: dict[str, dict[str, float]] | None = None

    def __post_init__(self) -> None:
        super(NumericResult, self).__post_init__()
        if self.question_type not in {
            QuestionType.DIRECT_NUMERIC,
            QuestionType.NUMERIC_ALLOCATION,
        }:
            raise ValueError(
                "NumericResult requires question_type DIRECT_NUMERIC or "
                "NUMERIC_ALLOCATION"
            )

        for field_name in ("mean", "median", "std", "min_val", "max_val"):
            _require_numeric(getattr(self, field_name), field_name)
        if self.min_val > self.max_val:
            raise ValueError("min_val must be <= max_val")

        required_percentiles = {25, 50, 75}
        if not required_percentiles.issubset(self.percentiles):
            raise ValueError("percentiles must include keys 25, 50, and 75")
        for percentile, value in self.percentiles.items():
            if not isinstance(percentile, int) or isinstance(percentile, bool):
                raise ValueError("percentile keys must be integers")
            _require_numeric(value, "percentile value")

        if self.question_type is QuestionType.NUMERIC_ALLOCATION:
            if self.allocation_target is None:
                raise ValueError(
                    "allocation_target must be set for NUMERIC_ALLOCATION"
                )
            if self.allocation_tolerance is None:
                raise ValueError(
                    "allocation_tolerance must be set for NUMERIC_ALLOCATION"
                )
            if self.allocation_excluded_n is None:
                raise ValueError(
                    "allocation_excluded_n must be set for NUMERIC_ALLOCATION"
                )
            if self.per_option_stats is None:
                raise ValueError(
                    "per_option_stats must be set for NUMERIC_ALLOCATION"
                )

        if self.allocation_target is not None:
            _require_numeric(self.allocation_target, "allocation_target")
        if self.allocation_tolerance is not None:
            _require_numeric(self.allocation_tolerance, "allocation_tolerance")
            if self.allocation_tolerance < 0:
                raise ValueError("allocation_tolerance must be non-negative")
        if self.allocation_excluded_n is not None:
            _require_non_negative_int(
                self.allocation_excluded_n, "allocation_excluded_n"
            )
        if self.per_option_stats is not None:
            for option_id, payload in self.per_option_stats.items():
                _require_non_empty_string(option_id, "per_option_stats key")
                if not isinstance(payload, dict):
                    raise ValueError("per_option_stats payloads must be dictionaries")
                for metric_name, metric_value in payload.items():
                    _require_non_empty_string(
                        metric_name, "per_option_stats metric name"
                    )
                    _require_numeric(metric_value, "per_option_stats metric value")


@dataclass(frozen=True, slots=True)
class GridSingleSelectResult(SingleCutResult):
    """Nested single-select results for each row of a grid question."""

    rows: dict[str, SingleSelectResult]
    overall_valid_n: int | None = None

    def __post_init__(self) -> None:
        super(GridSingleSelectResult, self).__post_init__()
        if self.question_type is not QuestionType.GRID_SINGLE_SELECT:
            raise ValueError(
                "GridSingleSelectResult requires question_type GRID_SINGLE_SELECT"
            )
        if not self.rows and self.valid_n > 0:
            raise ValueError("rows must be non-empty when valid_n > 0")
        if self.overall_valid_n is not None:
            _require_non_negative_int(self.overall_valid_n, "overall_valid_n")

        for row_column_id, result in self.rows.items():
            _require_non_empty_string(row_column_id, "rows key")
            if not isinstance(result, SingleSelectResult):
                raise ValueError("rows values must be SingleSelectResult instances")


@dataclass(frozen=True, slots=True)
class SkipRecord:
    """Records a question skipped or failed during analysis."""

    question_id: str
    canonical_id: str
    question_type: QuestionType
    skip_reason: str
    details: str | None = None

    def __post_init__(self) -> None:
        _require_non_empty_string(self.question_id, "question_id")
        _require_non_empty_string(self.canonical_id, "canonical_id")
        _require_non_empty_string(self.skip_reason, "skip_reason")


def _require_distribution_payload(
    payload: Any, rate_key: str, container_name: str
) -> None:
    if not isinstance(payload, dict):
        raise ValueError(f"{container_name} payloads must be dictionaries")

    for required_key in ("label", "count", rate_key):
        if required_key not in payload:
            raise ValueError(
                f"{container_name} payloads must include {required_key!r}"
            )

    _require_non_empty_string(payload["label"], f"{container_name} label")
    _require_non_negative_int(payload["count"], f"{container_name} count")
    _require_rate(payload[rate_key], f"{container_name} {rate_key}")


def _require_distribution_payload_without_rate(
    payload: Any, rate_key: str, container_name: str
) -> None:
    if not isinstance(payload, dict):
        raise ValueError(f"{container_name} payloads must be dictionaries")

    for required_key in ("label", "count", rate_key):
        if required_key not in payload:
            raise ValueError(
                f"{container_name} payloads must include {required_key!r}"
            )

    _require_non_empty_string(payload["label"], f"{container_name} label")
    _require_non_negative_int(payload["count"], f"{container_name} count")
    _require_numeric(payload[rate_key], f"{container_name} {rate_key}")


@dataclass(frozen=True, slots=True)
class DataQualityReport:
    """Data quality summary for parsed survey inputs."""

    total_rows: int
    total_columns: int
    columns_in_datamap: int
    columns_not_in_datamap: tuple[str, ...]
    per_column_missing_pct: dict[str, float]
    per_column_out_of_range_pct: dict[str, float]
    coercion_log: tuple[dict, ...]
    warnings: tuple[str, ...]

    def __post_init__(self) -> None:
        _require_non_negative_int(self.total_rows, "total_rows")
        _require_non_negative_int(self.total_columns, "total_columns")

        for column, value in self.per_column_missing_pct.items():
            _require_rate(value, f"per_column_missing_pct[{column!r}]")
        for column, value in self.per_column_out_of_range_pct.items():
            _require_rate(value, f"per_column_out_of_range_pct[{column!r}]")


@dataclass(frozen=True, slots=True)
class CrossCutResult:
    """Result contract for cross-question analyses."""

    cross_cut_id: str
    synthetic_question_title: str
    business_question: str
    source_question_ids: tuple[str, ...]
    analysis_type: AnalysisType
    result_table: dict
    ai_insight: str | None
    ai_insight_was_template: bool
    audit_records: tuple[AuditRecord, ...]
    warnings: tuple[str, ...]

    def __post_init__(self) -> None:
        _require_non_empty_string(self.cross_cut_id, "cross_cut_id")
        _require_non_empty_string(
            self.synthetic_question_title, "synthetic_question_title"
        )
        if not self.source_question_ids:
            raise ValueError("source_question_ids must be non-empty")
