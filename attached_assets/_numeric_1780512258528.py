"""Numeric single-cut calculators for direct numeric and allocation questions."""

from __future__ import annotations

from datetime import datetime, timezone

import numpy as np
import pandas as pd

from src.calc_primitives import allocation_summary, numeric_summary
from src.calculation_log import CalculationLog
from src.models import AuditRecord, NumericResult, QuestionSpec, QuestionType
from src.single_cut._conditional import apply_conditional_on_filter

try:
    from config import DEFAULT_ALLOCATION_TARGET, ALLOCATION_TOLERANCE
except ModuleNotFoundError as exc:
    if exc.name != "config":
        raise
    DEFAULT_ALLOCATION_TARGET = 100.0
    ALLOCATION_TOLERANCE = 2.0


def compute_numeric(
    question_spec: QuestionSpec,
    df: pd.DataFrame,
    log: CalculationLog,
    filter_mask: pd.Series | None = None,
    filter_expr: str | None = None,
) -> NumericResult:
    if question_spec.question_type is QuestionType.DIRECT_NUMERIC:
        return _compute_direct_numeric(
            question_spec=question_spec,
            df=df,
            log=log,
            filter_mask=filter_mask,
            filter_expr=filter_expr,
        )
    if question_spec.question_type is QuestionType.NUMERIC_ALLOCATION:
        return _compute_allocation(
            question_spec=question_spec,
            df=df,
            log=log,
            filter_mask=filter_mask,
            filter_expr=filter_expr,
        )
    raise ValueError(f"unsupported question_type: {question_spec.question_type}")


def _compute_direct_numeric(
    question_spec: QuestionSpec,
    df: pd.DataFrame,
    log: CalculationLog,
    filter_mask: pd.Series | None,
    filter_expr: str | None,
) -> NumericResult:
    column_name = question_spec.canonical_id
    if column_name not in df.columns:
        raise ValueError(f"raw column not found in data: {column_name}")

    working_df = df
    if filter_mask is not None:
        working_df = working_df[filter_mask]
        if filter_expr is None:
            filter_expr = "<unnamed filter>"
    else:
        filter_expr = None
    working_df, filter_expr = apply_conditional_on_filter(
        question_spec,
        working_df,
        filter_expr,
    )
    series = working_df[column_name]

    summary, audit = numeric_summary(
        series=series,
        question_id=question_spec.canonical_id,
        source_columns=(column_name,),
        output_sheet=f"SC_{question_spec.canonical_id}",
        filter_expr=filter_expr,
    )
    log.record(audit)
    stat_audits = _direct_numeric_stat_audits(
        question_spec=question_spec,
        series=series,
        summary=summary,
        filter_expr=filter_expr,
    )
    for stat_audit in stat_audits:
        log.record(stat_audit)

    warnings = ()
    if summary["valid_n"] == 0:
        warnings = ("all values are missing; no statistics computed",)

    return NumericResult(
        question_id=question_spec.canonical_id,
        question_type=QuestionType.DIRECT_NUMERIC,
        valid_n=int(summary["valid_n"]),
        missing_n=int(summary["missing_n"]),
        denominator_policy=question_spec.denominator_policy,
        mean=summary["mean"],
        median=summary["median"],
        std=summary["std"],
        min_val=summary["min"],
        max_val=summary["max"],
        percentiles={25: summary["p25"], 50: summary["p50"], 75: summary["p75"]},
        allocation_target=None,
        allocation_tolerance=None,
        allocation_excluded_n=None,
        per_option_stats=None,
        audit_records=(audit, *stat_audits),
        warnings=warnings,
    )


def _compute_allocation(
    question_spec: QuestionSpec,
    df: pd.DataFrame,
    log: CalculationLog,
    filter_mask: pd.Series | None,
    filter_expr: str | None,
) -> NumericResult:
    working_df = df
    if filter_mask is not None:
        working_df = working_df[filter_mask]
        if filter_expr is None:
            filter_expr = "<unnamed filter>"
    else:
        filter_expr = None
    working_df, filter_expr = apply_conditional_on_filter(
        question_spec,
        working_df,
        filter_expr,
    )

    columns: dict[str, pd.Series] = {}
    for column_name in question_spec.raw_columns:
        if column_name not in working_df.columns:
            raise ValueError(f"raw column not found in data: {column_name}")
        columns[column_name] = working_df[column_name]

    target_sum = float(DEFAULT_ALLOCATION_TARGET)
    tolerance = float(ALLOCATION_TOLERANCE)
    summary, audit = allocation_summary(
        columns=columns,
        question_id=question_spec.canonical_id,
        output_sheet=f"SC_{question_spec.canonical_id}",
        target_sum=target_sum,
        tolerance=tolerance,
        filter_expr=filter_expr,
    )
    log.record(audit)

    per_option_stats = _allocation_per_option_stats(columns)
    per_option_audits = _allocation_per_option_audits(
        question_spec=question_spec,
        columns=columns,
        per_option_stats=per_option_stats,
        filter_expr=filter_expr,
    )
    for per_option_audit in per_option_audits:
        log.record(per_option_audit)

    aggregate = _aggregate_per_option_means(per_option_stats)
    warnings: list[str] = []
    excluded_n = int(summary["excluded_tolerance_n"])
    if excluded_n:
        warnings.append(
            f"{excluded_n} respondents excluded for sum outside tolerance"
        )

    return NumericResult(
        question_id=question_spec.canonical_id,
        question_type=QuestionType.NUMERIC_ALLOCATION,
        valid_n=int(summary["included_n"]),
        missing_n=audit.missing_n,
        denominator_policy=question_spec.denominator_policy,
        mean=aggregate["mean"],
        median=aggregate["median"],
        std=aggregate["std"],
        min_val=aggregate["min"],
        max_val=aggregate["max"],
        percentiles={
            25: aggregate["p25"],
            50: aggregate["p50"],
            75: aggregate["p75"],
        },
        allocation_target=target_sum,
        allocation_tolerance=tolerance,
        allocation_excluded_n=excluded_n,
        per_option_stats=per_option_stats,
        audit_records=(audit, *per_option_audits),
        warnings=tuple(warnings),
    )


def _allocation_per_option_stats(
    columns: dict[str, pd.Series],
) -> dict[str, dict[str, float | int]]:
    stats: dict[str, dict[str, float | int]] = {}
    for column_id, series in columns.items():
        numeric = pd.to_numeric(series, errors="coerce")
        valid_values = numeric.dropna()
        valid_n = int(valid_values.count())
        missing_n = int(numeric.isna().sum())
        stats[column_id] = {
            "mean": float(valid_values.mean()) if valid_n else 0.0,
            "median": float(valid_values.median()) if valid_n else 0.0,
            "std": float(valid_values.std()) if valid_n > 1 else 0.0,
            "min_val": float(valid_values.min()) if valid_n else 0.0,
            "max_val": float(valid_values.max()) if valid_n else 0.0,
            "valid_n": valid_n,
            "missing_n": missing_n,
        }
    return stats


def _direct_numeric_stat_audits(
    question_spec: QuestionSpec,
    series: pd.Series,
    summary: dict[str, float],
    filter_expr: str | None,
) -> tuple[AuditRecord, ...]:
    numeric = pd.to_numeric(series, errors="coerce")
    valid_values = numeric.dropna()
    valid_n = int(summary["valid_n"])
    missing_n = int(summary["missing_n"])
    numerator = float(valid_values.sum()) if valid_n else 0.0
    stat_specs = (
        (
            "numeric_std",
            "std = series.dropna().std()",
            float(summary["std"]) if valid_n > 1 else 0.0,
        ),
        (
            "numeric_p25",
            "p25 = series.dropna().quantile(0.25)",
            float(summary["p25"]) if valid_n else 0.0,
        ),
        (
            "numeric_p50",
            "p50 = series.dropna().quantile(0.50)",
            float(summary["p50"]) if valid_n else 0.0,
        ),
        (
            "numeric_p75",
            "p75 = series.dropna().quantile(0.75)",
            float(summary["p75"]) if valid_n else 0.0,
        ),
    )
    return tuple(
        AuditRecord(
            output_sheet=f"SC_{question_spec.canonical_id}",
            metric_name=metric_name,
            source_question_id=question_spec.canonical_id,
            source_columns=question_spec.raw_columns,
            filter_expr=filter_expr,
            numerator=numerator,
            denominator=valid_n,
            formula=formula,
            value_raw=value,
            valid_n=valid_n,
            missing_n=missing_n,
            timestamp=datetime.now(timezone.utc),
        )
        for metric_name, formula, value in stat_specs
    )


def _allocation_per_option_audits(
    question_spec: QuestionSpec,
    columns: dict[str, pd.Series],
    per_option_stats: dict[str, dict[str, float | int]],
    filter_expr: str | None,
) -> tuple[AuditRecord, ...]:
    audits: list[AuditRecord] = []
    for column_id, series in columns.items():
        numeric = pd.to_numeric(series, errors="coerce")
        valid_values = numeric.dropna()
        valid_n = int(per_option_stats[column_id]["valid_n"])
        missing_n = int(per_option_stats[column_id]["missing_n"])
        numerator = float(valid_values.sum()) if valid_n else 0.0
        timestamp = datetime.now(timezone.utc)
        audits.append(
            AuditRecord(
                output_sheet=f"SC_{question_spec.canonical_id}",
                metric_name="numeric_allocation_mean",
                source_question_id=question_spec.canonical_id,
                source_columns=(column_id,),
                filter_expr=filter_expr,
                numerator=numerator,
                denominator=valid_n,
                formula="mean = series.dropna().mean()",
                value_raw=float(per_option_stats[column_id]["mean"]),
                valid_n=valid_n,
                missing_n=missing_n,
                timestamp=timestamp,
            )
        )
        audits.append(
            AuditRecord(
                output_sheet=f"SC_{question_spec.canonical_id}",
                metric_name="numeric_allocation_median",
                source_question_id=question_spec.canonical_id,
                source_columns=(column_id,),
                filter_expr=filter_expr,
                numerator=numerator,
                denominator=valid_n,
                formula="median = series.dropna().median()",
                value_raw=float(per_option_stats[column_id]["median"]),
                valid_n=valid_n,
                missing_n=missing_n,
                timestamp=timestamp,
            )
        )
    return tuple(audits)


def _aggregate_per_option_means(
    per_option_stats: dict[str, dict[str, float | int]]
) -> dict[str, float]:
    all_means = np.array(
        [payload["mean"] for payload in per_option_stats.values()],
        dtype=float,
    )
    if len(all_means) == 0:
        return {
            "mean": float("nan"),
            "median": float("nan"),
            "std": float("nan"),
            "min": float("nan"),
            "max": float("nan"),
            "p25": float("nan"),
            "p50": float("nan"),
            "p75": float("nan"),
        }

    percentiles = np.percentile(all_means, [25, 50, 75])
    return {
        "mean": float(np.mean(all_means)),
        "median": float(np.median(all_means)),
        "std": float(np.std(all_means, ddof=1)) if len(all_means) > 1 else 0.0,
        "min": float(np.min(all_means)),
        "max": float(np.max(all_means)),
        "p25": float(percentiles[0]),
        "p50": float(percentiles[1]),
        "p75": float(percentiles[2]),
    }
