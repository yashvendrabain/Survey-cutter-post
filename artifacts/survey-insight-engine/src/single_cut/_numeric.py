"""Numeric single-cut calculators for direct numeric and allocation questions."""

from __future__ import annotations

import numpy as np
import pandas as pd

from src.calc_primitives import allocation_summary, numeric_summary
from src.calculation_log import CalculationLog
from src.models import NumericResult, QuestionSpec, QuestionType

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

    series = df[column_name]
    if filter_mask is not None:
        series = series[filter_mask]
        if filter_expr is None:
            filter_expr = "<unnamed filter>"
    else:
        filter_expr = None

    summary, audit = numeric_summary(
        series=series,
        question_id=question_spec.canonical_id,
        source_columns=(column_name,),
        output_sheet=f"SC_{question_spec.canonical_id}",
        filter_expr=filter_expr,
    )
    log.record(audit)

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
        audit_records=(audit,),
        warnings=warnings,
    )


def _compute_allocation(
    question_spec: QuestionSpec,
    df: pd.DataFrame,
    log: CalculationLog,
    filter_mask: pd.Series | None,
    filter_expr: str | None,
) -> NumericResult:
    columns: dict[str, pd.Series] = {}
    for column_name in question_spec.raw_columns:
        if column_name not in df.columns:
            raise ValueError(f"raw column not found in data: {column_name}")
        columns[column_name] = df[column_name]

    if filter_mask is not None:
        columns = {
            column_name: series[filter_mask]
            for column_name, series in columns.items()
        }
        if filter_expr is None:
            filter_expr = "<unnamed filter>"
    else:
        filter_expr = None

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

    per_option_stats = summary["per_option"]
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
        audit_records=(audit,),
        warnings=tuple(warnings),
    )


def _aggregate_per_option_means(
    per_option_stats: dict[str, dict[str, float]]
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
