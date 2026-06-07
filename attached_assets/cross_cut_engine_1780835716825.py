"""Cross-question analysis engine for deterministic survey cuts."""

from __future__ import annotations

from dataclasses import asdict
from datetime import datetime, timezone
import re
from typing import Any

import pandas as pd

from src.calc_primitives import numeric_summary, selection_rate
from src.calculation_log import CalculationLog
from src.filter_options import _is_computed_multi_select_column, resolve_cross_cut_question
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
from src.single_cut.nps import compute_nps


FILTER_PATTERN = re.compile(r"^(\w+)\s*==\s*(\S+)$")
CATEGORICAL_TYPES = {
    QuestionType.SINGLE_SELECT,
    QuestionType.DEMOGRAPHIC_OR_SEGMENT,
    QuestionType.GRID_SINGLE_SELECT,
    QuestionType.NPS,
}
NUMERIC_TYPES = {
    QuestionType.DIRECT_NUMERIC,
    QuestionType.NUMERIC_ALLOCATION,
    QuestionType.GRID_RATED,
    QuestionType.NPS,
    QuestionType.MULTI_SELECT_BINARY,
    QuestionType.RANK_ORDER,
    QuestionType.GRID_BINARY_SELECT,
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
    if question_a.question_type is QuestionType.MULTI_SELECT_BINARY:
        raise ValueError(
            "CROSS_TAB does not yet support MULTI_SELECT_BINARY. "
            "Use SEGMENT_PROFILE with each sub-column instead."
        )
    if question_b.question_type is QuestionType.MULTI_SELECT_BINARY:
        raise ValueError(
            "CROSS_TAB does not yet support MULTI_SELECT_BINARY. "
            "Use SEGMENT_PROFILE with each sub-column instead."
        )
    if question_a.question_type not in CATEGORICAL_TYPES:
        raise ValueError(
            f"{question_a.canonical_id} is not categorical: {question_a.question_type}"
        )
    if question_b.question_type not in CATEGORICAL_TYPES:
        raise ValueError(
            f"{question_b.canonical_id} is not categorical: {question_b.question_type}"
        )
    series_a, row_label_map = _simple_categorical_column(question_a, df)
    series_b, column_label_map = _simple_categorical_column(question_b, df)
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
        "row_label_map": row_label_map,
        "column_label_map": column_label_map,
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
        source_columns=(
            question_a.canonical_id,
            question_b.canonical_id,
        ),
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

    filter_spec, target_spec = _source_specs(spec, schema, expected_count=2)
    filter_mask, filter_description = _filter_mask_from_expr(
        spec.filter_expr,
        filter_spec,
        df,
    )
    single_cut_result = _compute_filtered_single_cut(
        target_spec,
        df,
        log,
        filter_mask,
        spec.filter_mask_description or filter_description,
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
        "filter_mask_description": spec.filter_mask_description or filter_description,
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
    if segment_spec.question_type not in CATEGORICAL_TYPES:
        raise ValueError(
            f"{segment_spec.canonical_id} is not a supported segment question"
        )
    if metric_spec.question_type not in NUMERIC_TYPES:
        raise ValueError(
            f"{metric_spec.canonical_id} is not a supported metric question"
        )
    if metric_spec.question_type is QuestionType.NUMERIC_ALLOCATION:
        raise ValueError(
            "GROUP_COMPARISON does not yet support NUMERIC_ALLOCATION metrics. "
            "Use cross-cuts on individual allocation sub-columns instead."
        )
    if metric_spec.question_type is QuestionType.MULTI_SELECT_BINARY:
        return _compute_multi_select_group_comparison(
            spec,
            segment_spec,
            metric_spec,
            df,
            log,
        )
    if metric_spec.question_type is QuestionType.RANK_ORDER:
        return _compute_rank_group_comparison(
            spec,
            segment_spec,
            metric_spec,
            df,
            log,
        )
    if metric_spec.question_type is QuestionType.GRID_BINARY_SELECT:
        return _compute_grid_binary_group_comparison(
            spec,
            segment_spec,
            metric_spec,
            df,
            log,
        )
    if metric_spec.question_type is QuestionType.GRID_RATED:
        return _compute_grid_rated_group_comparison(
            spec,
            segment_spec,
            metric_spec,
            df,
            log,
        )
    if metric_spec.question_type is QuestionType.NPS:
        return _compute_nps_group_comparison(
            spec,
            segment_spec,
            metric_spec,
            df,
            log,
        )

    segment_series, segment_label_map = _simple_categorical_column(segment_spec, df)
    metric_column = _numeric_column(metric_spec)
    _require_columns(df, (metric_column,))

    per_segment: dict[Any, dict[str, Any]] = {}
    for segment_value in sorted(
        segment_series.dropna().unique(),
        key=lambda value: str(value),
    ):
        segment_mask = segment_series == segment_value
        summary, audit = numeric_summary(
            series=df.loc[segment_mask, metric_column],
            question_id=metric_spec.canonical_id,
            source_columns=(metric_column,),
            output_sheet=f"CC_{spec.cross_cut_id}",
            filter_expr=f"{segment_spec.canonical_id} == {segment_value}",
        )
        log.record(audit)
        segment_key = _python_scalar(segment_value)
        per_segment[segment_key] = {
            "label": segment_label_map.get(segment_key, str(segment_key)),
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
        source_columns=(segment_spec.canonical_id, metric_column),
        filter_expr=None,
        formula="numeric_summary(metric) per unique segment value + overall",
        value_raw=float(overall_summary["valid_n"]),
        valid_n=int(overall_summary["valid_n"]),
        missing_n=int(overall_summary["missing_n"]),
    )
    log.record(cross_audit)
    return _cross_cut_result(spec, result_table, (cross_audit,))


def _compute_grid_rated_group_comparison(
    spec: CrossCutSpec,
    segment_spec: QuestionSpec,
    metric_spec: QuestionSpec,
    df: pd.DataFrame,
    log: CalculationLog,
) -> CrossCutResult:
    segment_series, segment_label_map = _simple_categorical_column(segment_spec, df)
    source_columns = tuple(column for column in metric_spec.raw_columns if column in df.columns)
    if not source_columns:
        raise ValueError(f"{metric_spec.canonical_id} grid-rated columns not in raw data")

    segment_values = sorted(
        segment_series.dropna().unique(),
        key=lambda value: str(value),
    )
    grid_rows: dict[str, dict[str, Any]] = {}
    total_valid_n = 0
    total_missing_n = 0
    audit_records: list[AuditRecord] = []

    for source_column in source_columns:
        per_segment: dict[Any, dict[str, Any]] = {}
        for segment_value in segment_values:
            segment_mask = segment_series == segment_value
            summary, audit = numeric_summary(
                series=df.loc[segment_mask, source_column],
                question_id=f"{metric_spec.canonical_id}:{source_column}",
                source_columns=(source_column,),
                output_sheet=f"CC_{spec.cross_cut_id}",
                filter_expr=f"{segment_spec.canonical_id} == {segment_value}",
            )
            log.record(audit)
            audit_records.append(audit)
            segment_key = _python_scalar(segment_value)
            per_segment[segment_key] = {
                "label": segment_label_map.get(segment_key, str(segment_key)),
                "n": int(summary["valid_n"]),
                "missing_n": int(summary["missing_n"]),
                "mean": summary["mean"],
                "median": summary["median"],
                "std": summary["std"],
            }

        overall_summary, overall_audit = numeric_summary(
            series=df[source_column],
            question_id=f"{metric_spec.canonical_id}:{source_column}",
            source_columns=(source_column,),
            output_sheet=f"CC_{spec.cross_cut_id}",
            filter_expr=None,
        )
        log.record(overall_audit)
        audit_records.append(overall_audit)
        total_valid_n += int(overall_summary["valid_n"])
        total_missing_n += int(overall_summary["missing_n"])
        grid_rows[source_column] = {
            "label": _grid_row_label(metric_spec, source_column),
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

    result_table = {
        "segment_question_id": segment_spec.canonical_id,
        "metric_question_id": metric_spec.canonical_id,
        "metric_question_type": QuestionType.GRID_RATED.value,
        "grid_rows": grid_rows,
    }
    cross_audit = _audit(
        output_sheet=f"CC_{spec.cross_cut_id}",
        metric_name="grid_rated_group_comparison",
        source_question_id=f"{segment_spec.canonical_id} x {metric_spec.canonical_id}",
        source_columns=(segment_spec.canonical_id, *source_columns),
        filter_expr=None,
        formula="numeric_summary(each grid-rated row) per unique segment value + overall",
        value_raw=float(total_valid_n),
        valid_n=total_valid_n,
        missing_n=total_missing_n,
    )
    log.record(cross_audit)
    audit_records.append(cross_audit)
    return _cross_cut_result(spec, result_table, tuple(audit_records))


def _compute_multi_select_group_comparison(
    spec: CrossCutSpec,
    segment_spec: QuestionSpec,
    metric_spec: QuestionSpec,
    df: pd.DataFrame,
    log: CalculationLog,
) -> CrossCutResult:
    source_columns = tuple(
        column
        for column in metric_spec.raw_columns
        if column in df.columns and not _is_computed_multi_select_column(metric_spec, column)
    )
    if not source_columns:
        raise ValueError(f"{metric_spec.canonical_id} multi-select columns not in raw data")
    return _compute_binary_selection_group_comparison(
        spec=spec,
        segment_spec=segment_spec,
        metric_spec=metric_spec,
        df=df,
        log=log,
        source_columns=source_columns,
        row_key="selection_rate_rows",
        metric_question_type=QuestionType.MULTI_SELECT_BINARY.value,
        metric_name="multi_select_group_comparison",
        formula_label="selection_rate(each multi-select option) per unique segment value + overall",
    )


def _compute_grid_binary_group_comparison(
    spec: CrossCutSpec,
    segment_spec: QuestionSpec,
    metric_spec: QuestionSpec,
    df: pd.DataFrame,
    log: CalculationLog,
) -> CrossCutResult:
    source_columns = tuple(column for column in metric_spec.raw_columns if column in df.columns)
    if not source_columns:
        raise ValueError(f"{metric_spec.canonical_id} grid-binary columns not in raw data")
    return _compute_binary_selection_group_comparison(
        spec=spec,
        segment_spec=segment_spec,
        metric_spec=metric_spec,
        df=df,
        log=log,
        source_columns=source_columns,
        row_key="selection_rate_rows",
        metric_question_type=QuestionType.GRID_BINARY_SELECT.value,
        metric_name="grid_binary_group_comparison",
        formula_label="selection_rate(each grid-binary row) per unique segment value + overall",
    )


def _compute_binary_selection_group_comparison(
    *,
    spec: CrossCutSpec,
    segment_spec: QuestionSpec,
    metric_spec: QuestionSpec,
    df: pd.DataFrame,
    log: CalculationLog,
    source_columns: tuple[str, ...],
    row_key: str,
    metric_question_type: str,
    metric_name: str,
    formula_label: str,
) -> CrossCutResult:
    segment_series, segment_label_map = _simple_categorical_column(segment_spec, df)
    segment_values = sorted(
        segment_series.dropna().unique(),
        key=lambda value: str(value),
    )
    rows: dict[str, dict[str, Any]] = {}
    total_valid_n = 0
    total_missing_n = 0
    audit_records: list[AuditRecord] = []

    for source_column in source_columns:
        per_segment: dict[Any, dict[str, Any]] = {}
        for segment_value in segment_values:
            segment_mask = (segment_series == segment_value).fillna(False)
            result, audit = selection_rate(
                {source_column: df.loc[segment_mask, source_column]},
                question_id=f"{metric_spec.canonical_id}:{source_column}",
                output_sheet=f"CC_{spec.cross_cut_id}",
                denominator_policy="valid_responses",
                filter_expr=f"{segment_spec.canonical_id} == {segment_value}",
            )
            log.record(audit)
            audit_records.append(audit)
            payload = result[source_column]
            segment_key = _python_scalar(segment_value)
            per_segment[segment_key] = {
                "label": segment_label_map.get(segment_key, str(segment_key)),
                "n": int(audit.valid_n),
                "missing_n": int(audit.missing_n),
                "count": int(payload["count"]),
                "selection_rate": float(payload["selection_rate"]),
            }

        overall_result, overall_audit = selection_rate(
            {source_column: df[source_column]},
            question_id=f"{metric_spec.canonical_id}:{source_column}",
            output_sheet=f"CC_{spec.cross_cut_id}",
            denominator_policy="valid_responses",
            filter_expr=None,
        )
        log.record(overall_audit)
        audit_records.append(overall_audit)
        overall_payload = overall_result[source_column]
        total_valid_n += int(overall_audit.valid_n)
        total_missing_n += int(overall_audit.missing_n)
        rows[source_column] = {
            "label": _grid_row_label(metric_spec, source_column),
            "per_segment": per_segment,
            "overall": {
                "n": int(overall_audit.valid_n),
                "missing_n": int(overall_audit.missing_n),
                "count": int(overall_payload["count"]),
                "selection_rate": float(overall_payload["selection_rate"]),
            },
        }

    result_table = {
        "segment_question_id": segment_spec.canonical_id,
        "metric_question_id": metric_spec.canonical_id,
        "metric_question_type": metric_question_type,
        row_key: rows,
    }
    cross_audit = _audit(
        output_sheet=f"CC_{spec.cross_cut_id}",
        metric_name=metric_name,
        source_question_id=f"{segment_spec.canonical_id} x {metric_spec.canonical_id}",
        source_columns=(segment_spec.canonical_id, *source_columns),
        filter_expr=None,
        formula=formula_label,
        value_raw=float(total_valid_n),
        valid_n=total_valid_n,
        missing_n=total_missing_n,
    )
    log.record(cross_audit)
    audit_records.append(cross_audit)
    return _cross_cut_result(spec, result_table, tuple(audit_records))


def _compute_rank_group_comparison(
    spec: CrossCutSpec,
    segment_spec: QuestionSpec,
    metric_spec: QuestionSpec,
    df: pd.DataFrame,
    log: CalculationLog,
) -> CrossCutResult:
    segment_series, segment_label_map = _simple_categorical_column(segment_spec, df)
    source_columns = tuple(column for column in metric_spec.raw_columns if column in df.columns)
    if not source_columns:
        raise ValueError(f"{metric_spec.canonical_id} rank columns not in raw data")

    segment_values = sorted(
        segment_series.dropna().unique(),
        key=lambda value: str(value),
    )
    rank_rows: dict[str, dict[str, Any]] = {}
    total_valid_n = 0
    total_missing_n = 0
    audit_records: list[AuditRecord] = []

    for source_column in source_columns:
        per_segment: dict[Any, dict[str, Any]] = {}
        numeric_column = pd.to_numeric(df[source_column], errors="coerce")
        for segment_value in segment_values:
            segment_mask = (segment_series == segment_value).fillna(False)
            summary, audit = numeric_summary(
                series=numeric_column.loc[segment_mask],
                question_id=f"{metric_spec.canonical_id}:{source_column}",
                source_columns=(source_column,),
                output_sheet=f"CC_{spec.cross_cut_id}",
                filter_expr=f"{segment_spec.canonical_id} == {segment_value}",
            )
            log.record(audit)
            audit_records.append(audit)
            segment_key = _python_scalar(segment_value)
            per_segment[segment_key] = {
                "label": segment_label_map.get(segment_key, str(segment_key)),
                "n": int(summary["valid_n"]),
                "missing_n": int(summary["missing_n"]),
                "mean": summary["mean"],
                "median": summary["median"],
                "std": summary["std"],
            }

        overall_summary, overall_audit = numeric_summary(
            series=numeric_column,
            question_id=f"{metric_spec.canonical_id}:{source_column}",
            source_columns=(source_column,),
            output_sheet=f"CC_{spec.cross_cut_id}",
            filter_expr=None,
        )
        log.record(overall_audit)
        audit_records.append(overall_audit)
        total_valid_n += int(overall_summary["valid_n"])
        total_missing_n += int(overall_summary["missing_n"])
        rank_rows[source_column] = {
            "label": _grid_row_label(metric_spec, source_column),
            "per_segment": per_segment,
            "overall": {
                "n": int(overall_summary["valid_n"]),
                "missing_n": int(overall_summary["missing_n"]),
                "mean": overall_summary["mean"],
                "median": overall_summary["median"],
                "std": overall_summary["std"],
            },
        }

    result_table = {
        "segment_question_id": segment_spec.canonical_id,
        "metric_question_id": metric_spec.canonical_id,
        "metric_question_type": QuestionType.RANK_ORDER.value,
        "rank_rows": rank_rows,
    }
    cross_audit = _audit(
        output_sheet=f"CC_{spec.cross_cut_id}",
        metric_name="rank_order_group_comparison",
        source_question_id=f"{segment_spec.canonical_id} x {metric_spec.canonical_id}",
        source_columns=(segment_spec.canonical_id, *source_columns),
        filter_expr=None,
        formula="numeric_summary(each rank-order option) per unique segment value + overall",
        value_raw=float(total_valid_n),
        valid_n=total_valid_n,
        missing_n=total_missing_n,
    )
    log.record(cross_audit)
    audit_records.append(cross_audit)
    return _cross_cut_result(spec, result_table, tuple(audit_records))


def _compute_nps_group_comparison(
    spec: CrossCutSpec,
    segment_spec: QuestionSpec,
    metric_spec: QuestionSpec,
    df: pd.DataFrame,
    log: CalculationLog,
) -> CrossCutResult:
    segment_series, segment_label_map = _simple_categorical_column(segment_spec, df)
    source_columns = tuple(column for column in metric_spec.raw_columns if column in df.columns)
    if not source_columns:
        raise ValueError(f"{metric_spec.canonical_id} NPS columns not in raw data")

    segment_values = sorted(
        segment_series.dropna().unique(),
        key=lambda value: str(value),
    )
    nps_entities: dict[str, dict[str, Any]] = {}

    for segment_value in segment_values:
        segment_mask = (segment_series == segment_value).fillna(False)
        segment_result = compute_nps(
            metric_spec,
            df,
            CalculationLog(),
            filter_mask=segment_mask,
            filter_expr=f"{segment_spec.canonical_id} == {segment_value}",
        )
        segment_key = _python_scalar(segment_value)
        segment_label = segment_label_map.get(segment_key, str(segment_key))
        for source_column, entity in zip(source_columns, segment_result.entities):
            entity_payload = nps_entities.setdefault(
                source_column,
                {
                    "label": entity.entity_label,
                    "per_segment": {},
                },
            )
            entity_payload["per_segment"][segment_key] = {
                "label": segment_label,
                **_nps_entity_payload(entity),
            }

    overall_result = compute_nps(metric_spec, df, CalculationLog())
    total_valid_n = 0
    total_missing_n = 0
    for source_column, entity in zip(source_columns, overall_result.entities):
        entity_payload = nps_entities.setdefault(
            source_column,
            {
                "label": entity.entity_label,
                "per_segment": {},
            },
        )
        entity_payload["overall"] = _nps_entity_payload(entity)
        total_valid_n += int(entity.valid_n)
        total_missing_n += int(entity.missing_n)

    result_table = {
        "segment_question_id": segment_spec.canonical_id,
        "metric_question_id": metric_spec.canonical_id,
        "metric_question_type": QuestionType.NPS.value,
        "nps_entities": nps_entities,
    }
    cross_audit = _audit(
        output_sheet=f"CC_{spec.cross_cut_id}",
        metric_name="nps_group_comparison",
        source_question_id=f"{segment_spec.canonical_id} x {metric_spec.canonical_id}",
        source_columns=(segment_spec.canonical_id, *source_columns),
        filter_expr=None,
        formula="compute_nps(entity) per unique segment value + overall",
        value_raw=float(total_valid_n),
        valid_n=total_valid_n,
        missing_n=total_missing_n,
    )
    log.record(cross_audit)
    return _cross_cut_result(spec, result_table, (cross_audit,))


def _nps_entity_payload(entity: Any) -> dict[str, Any]:
    return {
        "n": int(entity.valid_n),
        "valid_n": int(entity.valid_n),
        "missing_n": int(entity.missing_n),
        "promoters": int(entity.promoters),
        "passives": int(entity.passives),
        "detractors": int(entity.detractors),
        "pct_promoters": float(entity.pct_promoters),
        "pct_passives": float(entity.pct_passives),
        "pct_detractors": float(entity.pct_detractors),
        "nps_score": float(entity.nps_score),
    }


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
        resolve_cross_cut_question(schema, source_id)
        for source_id in spec.source_question_ids
    )
    if any(question_spec is None for question_spec in question_specs):
        missing = [
            source_id
            for source_id, question_spec in zip(spec.source_question_ids, question_specs)
            if question_spec is None
        ]
        raise ValueError(f"source question not found: {', '.join(missing)}")
    return question_specs  # type: ignore[return-value]


def _simple_categorical_column(
    spec: QuestionSpec,
    df: pd.DataFrame,
) -> tuple[pd.Series, dict]:
    """Return a categorical Series for a question plus value labels."""

    if spec.question_type in (
        QuestionType.SINGLE_SELECT,
        QuestionType.DEMOGRAPHIC_OR_SEGMENT,
    ):
        if spec.canonical_id not in df.columns:
            raise ValueError(f"{spec.canonical_id} not in raw data")
        return df[spec.canonical_id], _label_map_with_decoded_values(spec.option_map)

    if spec.question_type is QuestionType.NPS:
        columns = [column for column in spec.raw_columns if column in df.columns]
        if not columns:
            raise ValueError(f"{spec.canonical_id} NPS columns not in raw data")
        result = pd.Series(pd.NA, index=df.index, dtype="object")
        for column in columns:
            bucket = df[column].map(_nps_bucket_label)
            result = result.where(result.notna(), bucket)
        return result, {
            "Promoter": "Promoter",
            "Passive": "Passive",
            "Detractor": "Detractor",
        }

    if spec.question_type is QuestionType.GRID_SINGLE_SELECT:
        if not spec.raw_columns:
            raise ValueError(f"{spec.canonical_id} grid has no raw columns")

        selected_value = _grid_selected_value(spec)
        selected_label = spec.option_map.get(selected_value)
        present_subs = [sub_column for sub_column in spec.raw_columns if sub_column in df.columns]
        if not present_subs:
            raise ValueError(
                f"{spec.canonical_id} grid sub-columns not in raw data"
            )

        sub_df = df[present_subs]
        mask = sub_df == selected_value
        if selected_label is not None:
            mask = mask | (sub_df == selected_label)
        has_any_selection = mask.any(axis=1)
        first_match = mask.idxmax(axis=1)
        result = first_match.where(has_any_selection)

        if spec.grid_row_labels:
            label_map = {
                sub_id: label
                for sub_id, label in spec.grid_row_labels.items()
                if sub_id in present_subs
            }
        else:
            label_map = {sub_column: sub_column for sub_column in present_subs}
        return result, label_map

    raise ValueError(
        f"{spec.canonical_id} is not categorical: "
        f"{spec.question_type}"
    )


def _grid_selected_value(spec: QuestionSpec) -> int:
    selected_value = 1
    if spec.option_map:
        positive_codes = [
            code
            for code in spec.option_map.keys()
            if isinstance(code, int) and code > 0
        ]
        if positive_codes:
            selected_value = min(positive_codes)
    return selected_value


def _grid_row_label(spec: QuestionSpec, source_column: str) -> str:
    if spec.grid_row_labels and source_column in spec.grid_row_labels:
        return str(spec.grid_row_labels[source_column])
    if source_column in spec.option_map:
        return str(spec.option_map[source_column])
    if source_column in spec.grid_column_labels:
        return str(spec.grid_column_labels[source_column])
    return source_column


def _nps_bucket_label(value: Any) -> str | None:
    score = _coerce_nps_score(value)
    if score is None:
        return None
    if score >= 9:
        return "Promoter"
    if score >= 7:
        return "Passive"
    return "Detractor"


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


def _filter_mask_from_expr(
    filter_expr: str,
    filter_spec: QuestionSpec,
    df: pd.DataFrame,
) -> tuple[pd.Series, str]:
    filter_column, filter_value = _parse_filter_expr_parts(filter_expr)
    if filter_column != filter_spec.canonical_id:
        raise ValueError(
            f"filter expression column {filter_column!r} does not match "
            f"source question {filter_spec.canonical_id!r}"
        )

    if filter_spec.question_type is QuestionType.GRID_SINGLE_SELECT:
        filter_value = str(filter_value)
        if filter_value not in df.columns:
            raise ValueError(f"grid sub-column {filter_value} not in data")
        selected_value = _grid_selected_value(filter_spec)
        selected_label = filter_spec.option_map.get(selected_value)
        label_map = (
            dict(filter_spec.grid_row_labels)
            if filter_spec.grid_row_labels
            else {filter_value: filter_value}
        )
        selected_mask = (df[filter_value] == selected_value).fillna(False)
        if selected_label is not None:
            selected_mask = selected_mask | (df[filter_value] == selected_label).fillna(False)
        return (
            selected_mask,
            f"{filter_spec.canonical_id} = "
            f"{label_map.get(filter_value, filter_value)}",
        )

    if filter_column not in df.columns:
        raise ValueError(f"filter column not found in data: {filter_column}")
    _, label_map = _simple_categorical_column(filter_spec, df)
    label = label_map.get(filter_value, str(filter_value))
    return _categorical_value_mask(
        df[filter_column],
        filter_value,
        filter_spec.option_map,
    ).fillna(False), (
        f"{filter_spec.canonical_id} = {label}"
    )


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
    filter_column, filter_value = _parse_filter_expr_parts(filter_expr)
    if filter_column not in df.columns:
        raise ValueError(f"filter column not found in data: {filter_column}")
    return (df[filter_column] == filter_value).fillna(False)


def _label_map_with_decoded_values(option_map: dict[int | str, str]) -> dict:
    label_map = dict(option_map)
    for label in option_map.values():
        label_map[label] = label
    return label_map


def _categorical_value_mask(
    series: pd.Series,
    value: int | str,
    option_map: dict[int | str, str],
) -> pd.Series:
    candidates = _expanded_filter_values(value, option_map)
    return series.isin(candidates)


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

    deduped: list[int | str] = []
    for candidate in candidates:
        if candidate not in deduped:
            deduped.append(candidate)
    return tuple(deduped)


def _parse_filter_expr_parts(filter_expr: str) -> tuple[str, int | str]:
    match = FILTER_PATTERN.match(filter_expr)
    if match is None:
        raise ValueError(f"unsupported filter expression: {filter_expr}")
    filter_column = match.group(1)
    filter_value_str = match.group(2).strip("\"'")
    try:
        filter_value: int | str = int(filter_value_str)
    except ValueError:
        filter_value = filter_value_str
    return filter_column, filter_value


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
    if target_spec.question_type is QuestionType.NPS:
        return compute_nps(target_spec, df, log, filter_mask, filter_expr)
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
        display_mode=spec.display_mode,
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
