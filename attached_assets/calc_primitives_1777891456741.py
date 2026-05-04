"""Vectorised calculation primitives with audit records."""

from __future__ import annotations

from datetime import datetime, timezone
from typing import Any

import pandas as pd

from src.models import AuditRecord


def _make_audit_record(
    metric_name: str,
    question_id: str,
    source_columns: tuple[str, ...],
    filter_expr: str | None,
    numerator: float | int | None,
    denominator: float | int | None,
    formula: str,
    value_raw: float,
    valid_n: int,
    missing_n: int,
    output_sheet: str,
) -> AuditRecord:
    return AuditRecord(
        output_sheet=output_sheet,
        metric_name=metric_name,
        source_question_id=question_id,
        source_columns=source_columns,
        filter_expr=filter_expr,
        numerator=numerator,
        denominator=denominator,
        formula=formula,
        value_raw=float(value_raw),
        valid_n=int(valid_n),
        missing_n=int(missing_n),
        timestamp=datetime.now(timezone.utc),
    )


def count_by_value(
    series: pd.Series,
    question_id: str,
    source_columns: tuple[str, ...],
    output_sheet: str,
    filter_expr: str | None = None,
) -> tuple[dict[Any, int], AuditRecord]:
    valid_n = int(series.notna().sum())
    missing_n = int(series.isna().sum())
    counts = {
        _python_scalar(value): int(count)
        for value, count in series.value_counts(dropna=True).to_dict().items()
    }
    audit = _make_audit_record(
        metric_name="value_counts",
        question_id=question_id,
        source_columns=source_columns,
        filter_expr=filter_expr,
        numerator=None,
        denominator=valid_n,
        formula="count(non-NA values per distinct level)",
        value_raw=float(valid_n),
        valid_n=valid_n,
        missing_n=missing_n,
        output_sheet=output_sheet,
    )
    return counts, audit


def rate_per_value(
    series: pd.Series,
    question_id: str,
    source_columns: tuple[str, ...],
    output_sheet: str,
    filter_expr: str | None = None,
) -> tuple[dict[Any, dict], AuditRecord]:
    counts, _ = count_by_value(
        series=series,
        question_id=question_id,
        source_columns=source_columns,
        output_sheet=output_sheet,
        filter_expr=filter_expr,
    )
    valid_n = int(series.notna().sum())
    missing_n = int(series.isna().sum())
    result = {
        value: {
            "count": count,
            "rate": float(count / valid_n) if valid_n else float("nan"),
        }
        for value, count in counts.items()
    }
    audit = _make_audit_record(
        metric_name="rate_per_value",
        question_id=question_id,
        source_columns=source_columns,
        filter_expr=filter_expr,
        numerator=None,
        denominator=valid_n,
        formula="count(value) / count(non-NA values)",
        value_raw=float(valid_n),
        valid_n=valid_n,
        missing_n=missing_n,
        output_sheet=output_sheet,
    )
    return result, audit


def selection_rate(
    binary_columns: dict[str, pd.Series],
    question_id: str,
    output_sheet: str,
    denominator_policy: str,
    all_respondents_n: int | None = None,
    filter_expr: str | None = None,
) -> tuple[dict[str, dict], AuditRecord]:
    dataframe = pd.DataFrame(binary_columns)
    total_rows = int(len(dataframe))

    if denominator_policy == "valid_responses":
        denom = int(dataframe.notna().any(axis=1).sum())
        formula = "count(selected) / count(answered_any)"
    elif denominator_policy == "all_respondents":
        if all_respondents_n is None:
            raise ValueError(
                "all_respondents_n required for all_respondents policy"
            )
        n_rows = len(next(iter(binary_columns.values())))
        if all_respondents_n < n_rows:
            raise ValueError(
                f"all_respondents_n ({all_respondents_n}) cannot "
                f"be less than the number of rows in "
                f"binary_columns ({n_rows})"
            )
        denom = int(all_respondents_n)
        missing_n = int(all_respondents_n - n_rows)
        formula = "count(selected) / all_respondents"
    else:
        raise ValueError(f"unknown denominator_policy: {denominator_policy}")

    result = {}
    for column_id in binary_columns:
        series = dataframe[column_id]
        selected_mask = (series != 0) & series.notna()
        count = int(selected_mask.sum())
        result[column_id] = {
            "count": count,
            "selection_rate": float(count / denom) if denom else float("nan"),
        }

    audit = _make_audit_record(
        metric_name="selection_rate",
        question_id=question_id,
        source_columns=tuple(binary_columns.keys()),
        filter_expr=filter_expr,
        numerator=None,
        denominator=denom,
        formula=formula,
        value_raw=float(denom),
        valid_n=denom,
        missing_n=missing_n if denominator_policy == "all_respondents" else total_rows - denom,
        output_sheet=output_sheet,
    )
    return result, audit


def numeric_summary(
    series: pd.Series,
    question_id: str,
    source_columns: tuple[str, ...],
    output_sheet: str,
    filter_expr: str | None = None,
) -> tuple[dict[str, float], AuditRecord]:
    valid_values = series.dropna()
    valid_n = int(len(valid_values))
    missing_n = int(series.isna().sum())

    if valid_n == 0:
        summary = {
            "mean": float("nan"),
            "median": float("nan"),
            "std": float("nan"),
            "min": float("nan"),
            "max": float("nan"),
            "p25": float("nan"),
            "p50": float("nan"),
            "p75": float("nan"),
            "valid_n": 0,
            "missing_n": missing_n,
        }
        value_raw = float("nan")
    else:
        quantiles = valid_values.quantile([0.25, 0.5, 0.75])
        summary = {
            "mean": float(valid_values.mean()),
            "median": float(valid_values.median()),
            "std": float(valid_values.std(ddof=1)),
            "min": float(valid_values.min()),
            "max": float(valid_values.max()),
            "p25": float(quantiles.loc[0.25]),
            "p50": float(quantiles.loc[0.5]),
            "p75": float(quantiles.loc[0.75]),
            "valid_n": valid_n,
            "missing_n": missing_n,
        }
        value_raw = summary["mean"]

    audit = _make_audit_record(
        metric_name="numeric_summary",
        question_id=question_id,
        source_columns=source_columns,
        filter_expr=filter_expr,
        numerator=None,
        denominator=valid_n,
        formula="mean / median / std (ddof=1) / percentiles over non-NA values",
        value_raw=value_raw,
        valid_n=valid_n,
        missing_n=missing_n,
        output_sheet=output_sheet,
    )
    return summary, audit


def allocation_summary(
    columns: dict[str, pd.Series],
    question_id: str,
    output_sheet: str,
    target_sum: float,
    tolerance: float,
    filter_expr: str | None = None,
) -> tuple[dict, AuditRecord]:
    dataframe = pd.DataFrame(columns)
    total_rows = int(len(dataframe))
    respondent_sums = dataframe.sum(axis=1, skipna=True)
    respondent_n_options = dataframe.notna().sum(axis=1)
    answered_mask = respondent_n_options > 0
    tolerance_mask = answered_mask & (
        (respondent_sums - target_sum).abs() > tolerance
    )
    included_mask = answered_mask & ~tolerance_mask
    included_dataframe = dataframe[included_mask]

    per_option = {
        column_id: {
            "mean": float(included_dataframe[column_id].mean()),
            "median": float(included_dataframe[column_id].median()),
        }
        for column_id in columns
    }
    answered_n = int(answered_mask.sum())
    included_n = int(included_mask.sum())
    excluded_tolerance_n = int(answered_n - included_n)
    result = {
        "per_option": per_option,
        "answered_n": answered_n,
        "included_n": included_n,
        "excluded_tolerance_n": excluded_tolerance_n,
        "target_sum": float(target_sum),
        "tolerance": float(tolerance),
    }

    audit = _make_audit_record(
        metric_name="allocation_summary",
        question_id=question_id,
        source_columns=tuple(columns.keys()),
        filter_expr=filter_expr,
        numerator=None,
        denominator=answered_n,
        formula=(
            "mean per option over respondents whose row sum is within "
            "tolerance of target"
        ),
        value_raw=float(included_n),
        valid_n=included_n,
        missing_n=total_rows - included_n,
        output_sheet=output_sheet,
    )
    return result, audit


def _python_scalar(value: Any) -> Any:
    if hasattr(value, "item"):
        try:
            value = value.item()
        except ValueError:
            return value
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value
