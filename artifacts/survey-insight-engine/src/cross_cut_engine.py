"""Cross-question analysis engine for deterministic survey cuts."""

from __future__ import annotations

from dataclasses import asdict
from datetime import datetime, timezone
import re
from typing import Any

import pandas as pd

from src.calc_primitives import numeric_summary
from src.calculation_log import CalculationLog
from src.models import (
    AnalysisType,
    AuditRecord,
    CrossCutResult,
    CrossCutSpec,
    QuestionSpec,
    QuestionType,
    SkipRecord,
    SurveySchema,
)
from src.single_cut._grid import compute_grid
from src.single_cut._multi_select import compute_multi_select
from src.single_cut._numeric import compute_numeric
from src.single_cut._single_select import compute_single_select


FILTER_PATTERN = re.compile(r"^(\w+)\s*==\s*(\S+)$")
CATEGORICAL_TYPES = {
    QuestionType.SINGLE_SELECT,
    QuestionType.MULTI_SELECT_BINARY,
    QuestionType.DEMOGRAPHIC_OR_SEGMENT,
}
SIMPLE_CATEGORICAL_TYPES = {
    QuestionType.SINGLE_SELECT,
    QuestionType.DEMOGRAPHIC_OR_SEGMENT,
}
NUMERIC_TYPES = {
    QuestionType.DIRECT_NUMERIC,
    QuestionType.NUMERIC_ALLOCATION,
}


def compute_cross_cuts(
    specs: list[CrossCutSpec],
    schema: SurveySchema,
    df: pd.DataFrame,
    log: CalculationLog,
) -> tuple[list[CrossCutResult], list[SkipRecord]]:
    """Run cross-cut specs against decoded data."""

    results: list[CrossCutResult] = []
    skips: list[SkipRecord] = []

    for spec in specs:
        question_specs = [
            schema.get_question(source_id)
            for source_id in spec.source_question_ids
        ]
        if any(question_spec is None for question_spec in question_specs):
            skips.append(
                SkipRecord(
                    question_id=spec.cross_cut_id,
                    canonical_id=spec.cross_cut_id,
                    question_type=QuestionType.UNKNOWN,
                    skip_reason="source_question_not_found",
                    details=", ".join(
                        source_id
                        for source_id, question_spec in zip(
                            spec.source_question_ids, question_specs
                        )
                        if question_spec is None
                    ),
                )
            )
            continue

        try:
            result = _dispatch(spec, schema, df, log)
        except Exception as exc:
            skips.append(
                SkipRecord(
                    question_id=spec.cross_cut_id,
                    canonical_id=spec.cross_cut_id,
                    question_type=QuestionType.UNKNOWN,
                    skip_reason="cross_cut_error",
                    details=f"{type(exc).__name__}: {exc}",
                )
            )
            continue

        results.append(result)

    return results, skips


def _dispatch(
    spec: CrossCutSpec,
    schema: SurveySchema,
    df: pd.DataFrame,
    log: CalculationLog,
) -> CrossCutResult:
    if spec.analysis_type is AnalysisType.CROSS_TAB:
        return _compute_cross_tab(spec, schema, df, log)
    if spec.analysis_type is AnalysisType.SEGMENT_PROFILE:
        return _compute_segment_profile(spec, schema, df, log)
    if spec.analysis_type is AnalysisType.GROUP_COMPARISON:
        return _compute_group_comparison(spec, schema, df, log)
    if spec.analysis_type is AnalysisType.EXPECTED_VS_REALIZED:
        return _compute_expected_vs_realized(spec, schema, df, log)
    if spec.analysis_type is AnalysisType.MULTI_QUESTION_METRIC:
        raise NotImplementedError("MULTI_QUESTION_METRIC is not implemented")
    raise ValueError(f"unsupported analysis_type: {spec.analysis_type}")


def _compute_cross_tab(
    spec: CrossCutSpec,
    schema: SurveySchema,
    df: pd.DataFrame,
    log: CalculationLog,
) -> CrossCutResult:
    question_a, question_b = _source_specs(spec, schema, expected_count=2)
    if question_a.question_type not in CATEGORICAL_TYPES:
        raise ValueError(
            f"{question_a.canonical_id} is not categorical: {question_a.question_type}"
        )
    if question_b.question_type not in CATEGORICAL_TYPES:
        raise ValueError(
            f"{question_b.canonical_id} is not categorical: {question_b.question_type}"
        )
    column_a = _simple_categorical_column(question_a)
    column_b = _simple_categorical_column(question_b)
    _require_columns(df, (column_a, column_b))

    series_a = df[column_a]
    series_b = df[column_b]
    counts = pd.crosstab(series_a, series_b, dropna=True)
    row_pct = counts.div(counts.sum(axis=1), axis=0)
    column_pct = counts.div(counts.sum(axis=0), axis=1)
    row_totals = counts.sum(axis=1)
    column_totals = counts.sum(axis=0)
    grand_total = int(counts.values.sum())
    missing_n = int((series_a.isna() | series_b.isna()).sum())

    result_table = {
        "row_question_id": question_a.canonical_id,
        "column_question_id": question_b.canonical_id,
        "row_label_map": dict(question_a.option_map),
        "column_label_map": dict(question_b.option_map),
        "counts": _frame_to_nested_dict(counts, value_type=int),
        "row_pct": _frame_to_nested_dict(row_pct, value_type=float),
        "column_pct": _frame_to_nested_dict(column_pct, value_type=float),
        "row_totals": _series_to_dict(row_totals, value_type=int),
        "column_totals": _series_to_dict(column_totals, value_type=int),
        "grand_total": grand_total,
    }
    audit = _audit(
        output_sheet=f"CC_{spec.cross_cut_id}",
        metric_name="cross_tab",
        source_question_id=f"{question_a.canonical_id} x {question_b.canonical_id}",
        source_columns=(column_a, column_b),
        filter_expr=None,
        formula=(
            "pd.crosstab(col_a, col_b, dropna=True); "
            "row% = count/row_total; col% = count/col_total"
        ),
        value_raw=float(grand_total),
        valid_n=grand_total,
        missing_n=missing_n,
    )
    log.record(audit)
    return _cross_cut_result(spec, result_table, (audit,))


def _compute_segment_profile(
    spec: CrossCutSpec,
    schema: SurveySchema,
    df: pd.DataFrame,
    log: CalculationLog,
) -> CrossCutResult:
    if len(spec.source_question_ids) != 2:
        raise ValueError("SEGMENT_PROFILE requires exactly 2 source questions")
    if spec.filter_expr is None:
        raise ValueError("SEGMENT_PROFILE requires filter_expr")

    _, target_spec = _source_specs(spec, schema, expected_count=2)
    filter_mask = _parse_filter_expr(spec.filter_expr, df)
    single_cut_result = _compute_filtered_single_cut(
        target_spec,
        df,
        log,
        filter_mask,
        spec.filter_mask_description or spec.filter_expr,
    )
    filter_n = int(filter_mask.sum())
    audit = _audit(
        output_sheet=f"CC_{spec.cross_cut_id}",
        metric_name="segment_profile",
        source_question_id=target_spec.canonical_id,
        source_columns=tuple(target_spec.raw_columns),
        filter_expr=spec.filter_expr,
        formula="single_cut(target) restricted to filter_expr",
        value_raw=float(filter_n),
        valid_n=filter_n,
        missing_n=int(len(df) - filter_n),
    )
    log.record(audit)
    result_table = {
        "filter_expr": spec.filter_expr,
        "filter_n": filter_n,
        "target_question_id": target_spec.canonical_id,
        "target_result": asdict(single_cut_result),
    }
    return _cross_cut_result(spec, result_table, (audit,))


def _compute_group_comparison(
    spec: CrossCutSpec,
    schema: SurveySchema,
    df: pd.DataFrame,
    log: CalculationLog,
) -> CrossCutResult:
    segment_spec, metric_spec = _source_specs(spec, schema, expected_count=2)
    if segment_spec.question_type not in {
        QuestionType.SINGLE_SELECT,
        QuestionType.DEMOGRAPHIC_OR_SEGMENT,
    }:
        raise ValueError(
            f"{segment_spec.canonical_id} is not a supported segment question"
        )
    if metric_spec.question_type not in NUMERIC_TYPES:
        raise ValueError(
            f"{metric_spec.canonical_id} is not a supported numeric question"
        )

    segment_column = _simple_categorical_column(segment_spec)
    metric_column = _numeric_column(metric_spec)
    _require_columns(df, (segment_column, metric_column))

    per_segment: dict[Any, dict[str, Any]] = {}
    for segment_value in sorted(
        df[segment_column].dropna().unique(),
        key=lambda value: str(value),
    ):
        segment_mask = df[segment_column] == segment_value
        summary, audit = numeric_summary(
            series=df.loc[segment_mask, metric_column],
            question_id=metric_spec.canonical_id,
            source_columns=(metric_column,),
            output_sheet=f"CC_{spec.cross_cut_id}",
            filter_expr=f"{segment_column} == {segment_value}",
        )
        log.record(audit)
        segment_key = _python_scalar(segment_value)
        per_segment[segment_key] = {
            "label": segment_spec.option_map.get(segment_key, str(segment_key)),
            "n": int(summary["valid_n"]),
            "missing_n": int(summary["missing_n"]),
            "mean": summary["mean"],
            "median": summary["median"],
            "std": summary["std"],
        }

    overall_summary, overall_audit = numeric_summary(
        series=df[metric_column],
        question_id=metric_spec.canonical_id,
        source_columns=(metric_column,),
        output_sheet=f"CC_{spec.cross_cut_id}",
        filter_expr=None,
    )
    log.record(overall_audit)
    result_table = {
        "segment_question_id": segment_spec.canonical_id,
        "metric_question_id": metric_spec.canonical_id,
        "per_segment": per_segment,
        "overall": {
            "n": int(overall_summary["valid_n"]),
            "missing_n": int(overall_summary["missing_n"]),
            "mean": overall_summary["mean"],
            "median": overall_summary["median"],
            "std": overall_summary["std"],
            "min": overall_summary["min"],
            "max": overall_summary["max"],
            "p25": overall_summary["p25"],
            "p50": overall_summary["p50"],
            "p75": overall_summary["p75"],
        },
    }
    cross_audit = _audit(
        output_sheet=f"CC_{spec.cross_cut_id}",
        metric_name="group_comparison",
        source_question_id=f"{segment_spec.canonical_id} x {metric_spec.canonical_id}",
        source_columns=(segment_column, metric_column),
        filter_expr=None,
        formula="numeric_summary(metric) per unique segment value + overall",
        value_raw=float(overall_summary["valid_n"]),
        valid_n=int(overall_summary["valid_n"]),
        missing_n=int(overall_summary["missing_n"]),
    )
    log.record(cross_audit)
    return _cross_cut_result(spec, result_table, (cross_audit,))


def _compute_expected_vs_realized(
    spec: CrossCutSpec,
    schema: SurveySchema,
    df: pd.DataFrame,
    log: CalculationLog,
) -> CrossCutResult:
    expected_spec, realized_spec = _source_specs(spec, schema, expected_count=2)
    if expected_spec.question_type is not QuestionType.DIRECT_NUMERIC:
        raise ValueError(
            f"{expected_spec.canonical_id} must be DIRECT_NUMERIC"
        )
    if realized_spec.question_type is not QuestionType.DIRECT_NUMERIC:
        raise ValueError(
            f"{realized_spec.canonical_id} must be DIRECT_NUMERIC"
        )
    expected_column = expected_spec.canonical_id
    realized_column = realized_spec.canonical_id
    _require_columns(df, (expected_column, realized_column))

    expected = pd.to_numeric(df[expected_column], errors="coerce")
    realized = pd.to_numeric(df[realized_column], errors="coerce")
    gap = realized - expected

    expected_summary, expected_audit = numeric_summary(
        series=expected,
        question_id=expected_spec.canonical_id,
        source_columns=(expected_column,),
        output_sheet=f"CC_{spec.cross_cut_id}",
    )
    realized_summary, realized_audit = numeric_summary(
        series=realized,
        question_id=realized_spec.canonical_id,
        source_columns=(realized_column,),
        output_sheet=f"CC_{spec.cross_cut_id}",
    )
    gap_summary, gap_audit = numeric_summary(
        series=gap,
        question_id=f"{realized_spec.canonical_id}_minus_{expected_spec.canonical_id}",
        source_columns=(expected_column, realized_column),
        output_sheet=f"CC_{spec.cross_cut_id}",
    )
    for audit in (expected_audit, realized_audit, gap_audit):
        log.record(audit)

    paired_n = int(gap.notna().sum())
    result_table = {
        "expected_question_id": expected_spec.canonical_id,
        "realized_question_id": realized_spec.canonical_id,
        "paired_n": paired_n,
        "expected": expected_summary,
        "realized": realized_summary,
        "gap": gap_summary,
    }
    cross_audit = _audit(
        output_sheet=f"CC_{spec.cross_cut_id}",
        metric_name="expected_vs_realized",
        source_question_id=f"{expected_spec.canonical_id} x {realized_spec.canonical_id}",
        source_columns=(expected_column, realized_column),
        filter_expr=None,
        formula=(
            "expected, realized, and gap = realized - expected; "
            "numeric_summary on each"
        ),
        value_raw=float(paired_n),
        valid_n=paired_n,
        missing_n=int(len(df) - paired_n),
    )
    log.record(cross_audit)
    return _cross_cut_result(spec, result_table, (cross_audit,))


def _source_specs(
    spec: CrossCutSpec,
    schema: SurveySchema,
    expected_count: int,
) -> tuple[QuestionSpec, ...]:
    if len(spec.source_question_ids) != expected_count:
        raise ValueError(
            f"{spec.analysis_type.value} requires exactly {expected_count} "
            "source questions"
        )
    question_specs = tuple(
        schema.get_question(source_id) for source_id in spec.source_question_ids
    )
    if any(question_spec is None for question_spec in question_specs):
        missing = [
            source_id
            for source_id, question_spec in zip(spec.source_question_ids, question_specs)
            if question_spec is None
        ]
        raise ValueError(f"source question not found: {', '.join(missing)}")
    return question_specs  # type: ignore[return-value]


def _simple_categorical_column(question_spec: QuestionSpec) -> str:
    if question_spec.question_type not in SIMPLE_CATEGORICAL_TYPES:
        raise ValueError(
            "cross-tab expansion for multi-select questions is not implemented"
        )
    return question_spec.canonical_id


def _numeric_column(question_spec: QuestionSpec) -> str:
    if question_spec.question_type is QuestionType.NUMERIC_ALLOCATION:
        if not question_spec.raw_columns:
            raise ValueError("NUMERIC_ALLOCATION requires at least one raw column")
        return question_spec.raw_columns[0]
    return question_spec.canonical_id


def _require_columns(df: pd.DataFrame, columns: tuple[str, ...]) -> None:
    for column in columns:
        if column not in df.columns:
            raise ValueError(f"raw column not found in data: {column}")


def _parse_filter_expr(filter_expr: str, df: pd.DataFrame) -> pd.Series:
    match = FILTER_PATTERN.match(filter_expr)
    if match is None:
        raise ValueError(f"unsupported filter expression: {filter_expr}")
    filter_column = match.group(1)
    if filter_column not in df.columns:
        raise ValueError(f"filter column not found in data: {filter_column}")
    filter_value_str = match.group(2).strip("\"'")
    try:
        filter_value: int | str = int(filter_value_str)
    except ValueError:
        filter_value = filter_value_str
    return (df[filter_column] == filter_value).fillna(False)


def _compute_filtered_single_cut(
    target_spec: QuestionSpec,
    df: pd.DataFrame,
    log: CalculationLog,
    filter_mask: pd.Series,
    filter_expr: str,
) -> Any:
    if target_spec.question_type is QuestionType.SINGLE_SELECT:
        return compute_single_select(target_spec, df, log, filter_mask, filter_expr)
    if target_spec.question_type is QuestionType.MULTI_SELECT_BINARY:
        return compute_multi_select(target_spec, df, log, filter_mask, filter_expr)
    if target_spec.question_type in NUMERIC_TYPES:
        return compute_numeric(target_spec, df, log, filter_mask, filter_expr)
    if target_spec.question_type is QuestionType.GRID_SINGLE_SELECT:
        return compute_grid(target_spec, df, log, filter_mask, filter_expr)
    raise ValueError(
        f"unsupported target question type for segment profile: "
        f"{target_spec.question_type}"
    )


def _cross_cut_result(
    spec: CrossCutSpec,
    result_table: dict,
    audit_records: tuple[AuditRecord, ...],
    warnings: tuple[str, ...] = (),
) -> CrossCutResult:
    return CrossCutResult(
        cross_cut_id=spec.cross_cut_id,
        synthetic_question_title=spec.title,
        business_question=spec.title,
        source_question_ids=spec.source_question_ids,
        analysis_type=spec.analysis_type,
        result_table=result_table,
        ai_insight=None,
        ai_insight_was_template=False,
        audit_records=audit_records,
        warnings=warnings,
    )


def _audit(
    output_sheet: str,
    metric_name: str,
    source_question_id: str,
    source_columns: tuple[str, ...],
    filter_expr: str | None,
    formula: str,
    value_raw: float,
    valid_n: int,
    missing_n: int,
) -> AuditRecord:
    return AuditRecord(
        output_sheet=output_sheet,
        metric_name=metric_name,
        source_question_id=source_question_id,
        source_columns=source_columns,
        filter_expr=filter_expr,
        numerator=None,
        denominator=valid_n,
        formula=formula,
        value_raw=value_raw,
        valid_n=valid_n,
        missing_n=missing_n,
        timestamp=datetime.now(timezone.utc),
    )


def _frame_to_nested_dict(
    frame: pd.DataFrame,
    value_type: type[int] | type[float],
) -> dict[Any, dict[Any, int | float]]:
    result: dict[Any, dict[Any, int | float]] = {}
    for index_value, row in frame.iterrows():
        result[_python_scalar(index_value)] = {
            _python_scalar(column): value_type(value)
            for column, value in row.items()
        }
    return result


def _series_to_dict(
    series: pd.Series,
    value_type: type[int] | type[float],
) -> dict[Any, int | float]:
    return {
        _python_scalar(index_value): value_type(value)
        for index_value, value in series.items()
    }


def _python_scalar(value: Any) -> Any:
    if hasattr(value, "item"):
        try:
            return value.item()
        except ValueError:
            return value
    return value
