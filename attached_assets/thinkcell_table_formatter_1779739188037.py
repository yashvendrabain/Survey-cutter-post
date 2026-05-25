"""Format result data for think-cell chart datasheets."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from src.chart_recommender import ChartRecommendation, ChartType
from src.models import (
    GridBinaryPivotResult,
    GridSingleSelectResult,
    GridRatedResult,
    MultiSelectResult,
    NumericResult,
    RankOrderResult,
    SingleCutResult,
    SingleSelectResult,
)


@dataclass(frozen=True, slots=True)
class ThinkCellTablePayload:
    """Structured table data formatted for think-cell automation."""

    headers: list[str]
    rows: list[list[Any]]
    chart_type: ChartType
    title: str
    source_line: str

    def to_tsv(self) -> str:
        """Render tab-separated values for paste into a think-cell datasheet."""
        lines = ["\t".join(str(h) for h in [""] + self.headers)]
        for row in self.rows:
            lines.append("\t".join("" if value is None else str(value) for value in row))
        return "\n".join(lines)

    def to_json(self) -> dict[str, Any]:
        """Render a JSON-style table payload for future think-cell automation."""
        table: list[list[dict[str, Any] | None]] = [
            [None] + [{"string": str(header)} for header in self.headers]
        ]
        for row in self.rows:
            row_cells: list[dict[str, Any]] = []
            for index, cell in enumerate(row):
                if cell is None:
                    row_cells.append({})
                elif index == 0:
                    row_cells.append({"string": str(cell)})
                elif isinstance(cell, (int, float)) and not isinstance(cell, bool):
                    row_cells.append({"number": float(cell)})
                else:
                    row_cells.append({"string": str(cell)})
            table.append(row_cells)
        return {"table": table}


def format_for_thinkcell(
    result: SingleCutResult,
    recommendation: ChartRecommendation,
    question_text: str = "",
    survey_name: str = "Survey",
) -> ThinkCellTablePayload:
    """Return a think-cell-ready table payload for a result."""

    n_display = f"N={result.valid_n}"
    source = f"Source: Bain {survey_name} ({n_display})"
    title = question_text[:120] if question_text else f"Question {result.question_id}"

    if isinstance(result, SingleSelectResult):
        return _format_single_select(result, recommendation, title, source)
    if isinstance(result, MultiSelectResult):
        return _format_multi_select(result, recommendation, title, source)
    if isinstance(result, NumericResult):
        return _format_numeric(result, recommendation, title, source)
    if isinstance(result, RankOrderResult):
        return _format_rank_order(result, recommendation, title, source)
    if isinstance(result, GridRatedResult):
        return _format_grid_rated(result, recommendation, title, source)
    if isinstance(result, GridBinaryPivotResult):
        return _format_grid_binary_pivot(result, recommendation, title, source)
    if isinstance(result, GridSingleSelectResult):
        return _format_grid_single_select(result, recommendation, title, source)

    raise ValueError(f"Unsupported result type: {type(result).__name__}")


def _format_single_select(
    result: SingleSelectResult,
    recommendation: ChartRecommendation,
    title: str,
    source: str,
) -> ThinkCellTablePayload:
    items = sorted(result.distribution.items(), key=lambda item: -item[1]["rate"])
    rows = [[payload["label"], payload["rate"]] for _code, payload in items]
    return ThinkCellTablePayload(
        headers=["Respondents"],
        rows=rows,
        chart_type=recommendation.chart_type,
        title=title,
        source_line=source,
    )


def _format_multi_select(
    result: MultiSelectResult,
    recommendation: ChartRecommendation,
    title: str,
    source: str,
) -> ThinkCellTablePayload:
    items = sorted(
        result.selections.items(),
        key=lambda item: -item[1].get("selection_rate", item[1].get("rate", 0)),
    )
    rows = [
        [
            payload.get("label", option_id),
            payload.get("selection_rate", payload.get("rate", 0)),
        ]
        for option_id, payload in items
    ]
    return ThinkCellTablePayload(
        headers=["Selection rate"],
        rows=rows,
        chart_type=recommendation.chart_type,
        title=title,
        source_line=source,
    )


def _format_numeric(
    result: NumericResult,
    recommendation: ChartRecommendation,
    title: str,
    source: str,
) -> ThinkCellTablePayload:
    if result.per_option_stats:
        items = sorted(
            result.per_option_stats.items(),
            key=lambda item: -item[1].get("mean", 0),
        )
        rows = [
            [payload.get("label", option_id), payload.get("mean", 0)]
            for option_id, payload in items
        ]
    else:
        rows = [
            ["Mean", result.mean],
            ["Median", result.median],
            ["Min", result.min_val],
            ["Max", result.max_val],
        ]
    return ThinkCellTablePayload(
        headers=["Value"],
        rows=rows,
        chart_type=recommendation.chart_type,
        title=title,
        source_line=source,
    )


def _format_rank_order(
    result: RankOrderResult,
    recommendation: ChartRecommendation,
    title: str,
    source: str,
) -> ThinkCellTablePayload:
    rows_sorted = sorted(
        result.rows,
        key=lambda row: -row.counts_per_rank[0] if row.counts_per_rank else 0,
    )
    headers = [row.option_label for row in rows_sorted]
    rows = []
    for rank_index in range(result.K):
        rank_values = [row.counts_per_rank[rank_index] for row in rows_sorted]
        rows.append([f"Rank {rank_index + 1}"] + rank_values)
    return ThinkCellTablePayload(
        headers=headers,
        rows=rows,
        chart_type=recommendation.chart_type,
        title=title,
        source_line=source,
    )


def _format_grid_rated(
    result: GridRatedResult,
    recommendation: ChartRecommendation,
    title: str,
    source: str,
) -> ThinkCellTablePayload:
    if recommendation.show_delta and len(result.column_headers) == 2:
        sorted_rows = sorted(
            result.rows,
            key=lambda row: -(row.delta if row.delta is not None else 0),
        )
    else:
        sorted_rows = list(result.rows)

    rows = [[row.row_label] + list(row.means_per_column) for row in sorted_rows]
    return ThinkCellTablePayload(
        headers=list(result.column_headers),
        rows=rows,
        chart_type=recommendation.chart_type,
        title=title,
        source_line=source,
    )


def _format_grid_binary_pivot(
    result: GridBinaryPivotResult,
    recommendation: ChartRecommendation,
    title: str,
    source: str,
) -> ThinkCellTablePayload:
    rows = [[row.row_label] + list(row.pcts_per_column) for row in result.rows]
    return ThinkCellTablePayload(
        headers=list(result.column_headers),
        rows=rows,
        chart_type=recommendation.chart_type,
        title=title,
        source_line=source,
    )


def _format_grid_single_select(
    result: GridSingleSelectResult,
    recommendation: ChartRecommendation,
    title: str,
    source: str,
) -> ThinkCellTablePayload:
    rows: list[list[Any]] = []
    for row_id, sub_result in result.rows.items():
        for _code, payload in sub_result.distribution.items():
            rows.append([f"{row_id} - {payload['label']}", payload["rate"]])
    return ThinkCellTablePayload(
        headers=["Selection rate"],
        rows=rows,
        chart_type=recommendation.chart_type,
        title=title,
        source_line=source,
    )
