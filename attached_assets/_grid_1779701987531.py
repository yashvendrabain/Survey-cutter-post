"""Grid single-select single-cut calculator."""

from __future__ import annotations

from datetime import datetime, timezone

import pandas as pd

from src.calculation_log import CalculationLog
from src.models import (
    AuditRecord,
    GridSingleSelectResult,
    QuestionSpec,
    QuestionType,
    SingleSelectResult,
)
from src.single_cut._single_select import compute_single_select


def compute_grid(
    question_spec: QuestionSpec,
    df: pd.DataFrame,
    log: CalculationLog,
    filter_mask: pd.Series | None = None,
    filter_expr: str | None = None,
) -> GridSingleSelectResult:
    if question_spec.question_type is not QuestionType.GRID_SINGLE_SELECT:
        raise ValueError(f"unsupported question_type for grid: {question_spec.question_type}")
    if not question_spec.grid_row_labels:
        raise ValueError("grid_row_labels must be non-empty")

    if filter_mask is not None and filter_expr is None:
        filter_expr = "<unnamed filter>"
    elif filter_mask is None:
        filter_expr = None

    warnings: list[str] = []
    rows = {}
    for sub_column_id, row_label in question_spec.grid_row_labels.items():
        if sub_column_id not in df.columns:
            warnings.append(
                f"row {sub_column_id} ({row_label}) not in raw data; skipped"
            )
            continue

        row_spec = QuestionSpec(
            question_id=f"{question_spec.canonical_id}.{sub_column_id}",
            canonical_id=sub_column_id,
            question_text=row_label,
            question_type=QuestionType.SINGLE_SELECT,
            raw_columns=(sub_column_id,),
            option_map=question_spec.option_map,
            value_range=question_spec.value_range,
            denominator_policy=question_spec.denominator_policy,
            theme_tags=question_spec.theme_tags,
            possible_role=question_spec.possible_role,
            analysis_eligible=True,
            parent_question_id=question_spec.canonical_id,
        )
        row_result = compute_single_select(
            row_spec,
            df,
            log,
            filter_mask=filter_mask,
            filter_expr=filter_expr,
        )
        row_result = _checked_only_row_result(row_result)
        if row_result is not None:
            rows[sub_column_id] = row_result

    present_columns = [
        sub_column_id
        for sub_column_id in question_spec.grid_row_labels
        if sub_column_id in df.columns
    ]
    filtered_n = int(filter_mask.sum()) if filter_mask is not None else len(df)
    if not present_columns:
        overall_valid_n = 0
        warnings.append("no grid rows present in raw data")
    else:
        sub_df = df[present_columns]
        if filter_mask is not None:
            sub_df = sub_df[filter_mask]
        overall_valid_n = int(sub_df.notna().any(axis=1).sum())
    missing_n = int(filtered_n - overall_valid_n)

    parent_audit = AuditRecord(
        output_sheet=f"SC_{question_spec.canonical_id}",
        metric_name="grid_overall",
        source_question_id=question_spec.canonical_id,
        source_columns=tuple(question_spec.grid_row_labels.keys()),
        filter_expr=filter_expr,
        numerator=None,
        denominator=overall_valid_n,
        formula="count(respondents who answered at least one row)",
        value_raw=float(overall_valid_n),
        valid_n=overall_valid_n,
        missing_n=missing_n,
        timestamp=datetime.now(timezone.utc),
    )
    log.record(parent_audit)

    return GridSingleSelectResult(
        question_id=question_spec.canonical_id,
        question_type=QuestionType.GRID_SINGLE_SELECT,
        valid_n=overall_valid_n,
        missing_n=missing_n,
        denominator_policy=question_spec.denominator_policy,
        rows=rows,
        overall_valid_n=overall_valid_n,
        audit_records=(parent_audit,),
        warnings=tuple(warnings),
    )


def _checked_only_row_result(
    row_result: SingleSelectResult,
) -> SingleSelectResult | None:
    filtered_distribution = {
        code: payload
        for code, payload in row_result.distribution.items()
        if not _is_unchecked_value(code)
        and (
            int(payload.get("count", 0)) > 0
            or float(payload.get("rate", 0.0)) > 0.0
        )
    }
    if not filtered_distribution:
        return None

    return SingleSelectResult(
        question_id=row_result.question_id,
        question_type=row_result.question_type,
        valid_n=row_result.valid_n,
        missing_n=row_result.missing_n,
        denominator_policy=row_result.denominator_policy,
        distribution=filtered_distribution,
        warnings=row_result.warnings,
        audit_records=row_result.audit_records,
    )


def _is_unchecked_value(value: object) -> bool:
    if value == 0 or value == "0":
        return True
    if isinstance(value, str):
        return value.strip().lower() in {"unchecked", "not selected"}
    return False
