"""Net Promoter Score single-cut calculator."""

from __future__ import annotations

from typing import Any

import pandas as pd

from src.calc_primitives import _make_audit_record, safe_share
from src.calculation_log import CalculationLog
from src.models import (
    DenominatorPolicy,
    NPSEntityResult,
    NPSResult,
    QuestionSpec,
    QuestionType,
)
from src.single_cut._conditional import apply_conditional_on_filter


def compute_nps(
    question_spec: QuestionSpec,
    df: pd.DataFrame,
    log: CalculationLog,
    filter_mask: pd.Series | None = None,
    filter_expr: str | None = None,
) -> NPSResult:
    """Compute NPS per entity column."""

    if question_spec.question_type is not QuestionType.NPS:
        raise ValueError(f"unsupported question_type for NPS: {question_spec.question_type}")

    working_df = df
    if filter_mask is not None:
        working_df = working_df[filter_mask]
        filter_expr = filter_expr or "<unnamed filter>"
    else:
        filter_expr = None
    working_df, filter_expr = apply_conditional_on_filter(
        question_spec,
        working_df,
        filter_expr,
    )

    columns = [column for column in question_spec.raw_columns if column in working_df.columns]
    if not columns:
        raise ValueError("NPS raw columns not found in dataframe")

    filtered_n = int(len(working_df))
    entity_results: list[NPSEntityResult] = []
    audit_records = []
    valid_columns: dict[str, pd.Series] = {}

    for column in columns:
        scores = working_df[column].map(_coerce_nps_score)
        valid_columns[column] = scores
        valid_mask = scores.notna()
        valid_n = int(valid_mask.sum())
        missing_n = int(filtered_n - valid_n)
        promoters = int(((scores >= 9) & (scores <= 10)).sum())
        passives = int(((scores >= 7) & (scores <= 8)).sum())
        detractors = int(((scores >= 0) & (scores <= 6)).sum())

        pct_promoters = safe_share(promoters, valid_n)
        pct_passives = safe_share(passives, valid_n)
        pct_detractors = safe_share(detractors, valid_n)
        nps_score = (pct_promoters - pct_detractors) * 100.0
        entity_label = _entity_label(question_spec, column)

        entity = NPSEntityResult(
            entity_label=entity_label,
            promoters=promoters,
            passives=passives,
            detractors=detractors,
            pct_promoters=pct_promoters,
            pct_passives=pct_passives,
            pct_detractors=pct_detractors,
            nps_score=nps_score,
            valid_n=valid_n,
            missing_n=missing_n,
        )
        entity_results.append(entity)
        audit_records.extend(
            _nps_audit_records(
                question_spec=question_spec,
                column=column,
                entity=entity,
                filter_expr=filter_expr,
            )
        )

    for audit in audit_records:
        log.record(audit)

    valid_frame = pd.DataFrame(valid_columns)
    overall_valid_n = int(valid_frame.notna().any(axis=1).sum())
    overall_missing_n = int(filtered_n - overall_valid_n)
    warnings = ()
    if overall_valid_n == 0:
        warnings = ("all NPS values are missing or outside 0-10",)

    return NPSResult(
        question_id=question_spec.canonical_id,
        question_type=QuestionType.NPS,
        valid_n=overall_valid_n,
        missing_n=overall_missing_n,
        denominator_policy=DenominatorPolicy.VALID_RESPONSES,
        entities=entity_results,
        warnings=warnings,
        audit_records=tuple(audit_records),
    )


def _coerce_nps_score(value: Any) -> int | None:
    if pd.isna(value):
        return None
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


def _entity_label(question_spec: QuestionSpec, column: str) -> str:
    if column in question_spec.grid_column_labels:
        return str(question_spec.grid_column_labels[column])
    if column in question_spec.option_map:
        return str(question_spec.option_map[column])
    if "|" in column:
        return column.rsplit("|", 1)[1].strip()
    if ":" in column:
        return column.split(":", 1)[1].strip()
    return question_spec.question_text or question_spec.canonical_id


def _nps_audit_records(
    *,
    question_spec: QuestionSpec,
    column: str,
    entity: NPSEntityResult,
    filter_expr: str | None,
):
    output_sheet = f"SC_{question_spec.canonical_id}"
    metric_prefix = _safe_metric_token(entity.entity_label)
    count_payloads = (
        ("promoters_count", entity.promoters, "count(scores in 9-10)"),
        ("passives_count", entity.passives, "count(scores in 7-8)"),
        ("detractors_count", entity.detractors, "count(scores in 0-6)"),
        ("valid_n", entity.valid_n, "count(integer scores in 0-10)"),
    )
    for metric_name, count_value, formula in count_payloads:
        yield _make_audit_record(
            metric_name=f"{metric_prefix}_{metric_name}",
            question_id=question_spec.canonical_id,
            source_columns=(column,),
            filter_expr=filter_expr,
            numerator=count_value,
            denominator=entity.valid_n,
            formula=formula,
            value_raw=float(count_value),
            valid_n=entity.valid_n,
            missing_n=entity.missing_n,
            output_sheet=output_sheet,
        )

    rate_payloads = (
        ("pct_promoters", entity.promoters, entity.pct_promoters, "promoters / valid_n"),
        ("pct_passives", entity.passives, entity.pct_passives, "passives / valid_n"),
        ("pct_detractors", entity.detractors, entity.pct_detractors, "detractors / valid_n"),
    )
    for metric_name, numerator, value, formula in rate_payloads:
        yield _make_audit_record(
            metric_name=f"{metric_prefix}_{metric_name}",
            question_id=question_spec.canonical_id,
            source_columns=(column,),
            filter_expr=filter_expr,
            numerator=numerator,
            denominator=entity.valid_n,
            formula=formula,
            value_raw=float(value),
            valid_n=entity.valid_n,
            missing_n=entity.missing_n,
            output_sheet=output_sheet,
        )

    yield _make_audit_record(
        metric_name=f"{metric_prefix}_nps_score",
        question_id=question_spec.canonical_id,
        source_columns=(column,),
        filter_expr=filter_expr,
        numerator=entity.promoters - entity.detractors,
        denominator=entity.valid_n,
        formula="(promoters / valid_n - detractors / valid_n) * 100",
        value_raw=entity.nps_score,
        valid_n=entity.valid_n,
        missing_n=entity.missing_n,
        output_sheet=output_sheet,
    )


def _safe_metric_token(value: str) -> str:
    import re

    token = re.sub(r"[^A-Za-z0-9]+", "_", value).strip("_").lower()
    return token or "entity"
