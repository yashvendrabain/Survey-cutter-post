"""Chart recommendation rules engine for survey-analysis results.

Maps deterministic result objects to Bain-style chart recommendations without
calling AI services. These recommendations are used both for inline Streamlit
rendering and for think-cell datasheet formatting.
"""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum

from src.models import (
    GridBinaryPivotResult,
    GridRatedResult,
    MultiSelectResult,
    NumericResult,
    QuestionSpec,
    QuestionType,
    RankOrderResult,
    SingleCutResult,
    SingleSelectResult,
)


class ChartType(str, Enum):
    COLUMN_STACKED = "COLUMN_STACKED"
    COLUMN_CLUSTERED = "COLUMN_CLUSTERED"
    BAR_STACKED = "BAR_STACKED"
    BAR_CLUSTERED = "BAR_CLUSTERED"
    LINE = "LINE"
    COMBO = "COMBO"
    HEATMAP_TABLE = "HEATMAP_TABLE"


BAIN_COLORS = {
    "bain_red": "#CC0000",
    "graphite": "#5C5C5C",
    "sky_blue": "#0070B9",
    "forest_green": "#2E7D32",
    "berry": "#6A1B9A",
    "sunset": "#E64A19",
    "beacon_violet": "#5C48D9",
    "white": "#FFFFFF",
    "stone": "#E3DAD4",
}

BAIN_SERIES_PALETTE = [
    BAIN_COLORS["graphite"],
    BAIN_COLORS["sky_blue"],
    BAIN_COLORS["forest_green"],
    BAIN_COLORS["berry"],
    BAIN_COLORS["sunset"],
]


@dataclass(frozen=True, slots=True)
class ChartRecommendation:
    """Recommendation produced by the chart rules engine."""

    chart_type: ChartType
    orientation: str | None
    primary_metric: str
    sort_order: str
    highlight_rule: str
    series_colors: list[str]
    data_label_format: str
    data_label_position: str
    axis_min: float | None = None
    axis_max: float | None = None
    show_delta: bool = False
    delta_format: str = "decimal_1"
    artifact_type: str = "chart"
    button_label_override: str | None = None
    notes: str = ""


def recommend_chart(
    result: SingleCutResult, spec: QuestionSpec | None = None
) -> ChartRecommendation:
    """Return a Bain-style chart recommendation for any single-cut result."""

    qtype = result.question_type

    if qtype is QuestionType.SINGLE_SELECT and isinstance(result, SingleSelectResult):
        return _recommend_single_select(result, spec)
    if qtype is QuestionType.MULTI_SELECT_BINARY and isinstance(
        result, MultiSelectResult
    ):
        return _recommend_multi_select(result, spec)
    if qtype in {QuestionType.NUMERIC_ALLOCATION, QuestionType.DIRECT_NUMERIC} and isinstance(
        result, NumericResult
    ):
        return _recommend_numeric_allocation(result, spec)
    if qtype is QuestionType.RANK_ORDER and isinstance(result, RankOrderResult):
        return _recommend_rank_order(result, spec)
    if qtype is QuestionType.GRID_RATED and isinstance(result, GridRatedResult):
        return _recommend_grid_rated(result, spec)
    if qtype is QuestionType.GRID_BINARY_SELECT and isinstance(
        result, GridBinaryPivotResult
    ):
        return _recommend_grid_binary_pivot(result, spec)
    if qtype is QuestionType.GRID_SINGLE_SELECT:
        return ChartRecommendation(
            chart_type=ChartType.BAR_CLUSTERED,
            orientation="horizontal",
            primary_metric="selection_rate",
            sort_order="descending",
            highlight_rule="top_1_to_3",
            series_colors=[BAIN_COLORS["bain_red"], BAIN_COLORS["graphite"]],
            data_label_format="percent_integer",
            data_label_position="outside_end",
            notes="Grid single-select treated as a multi-select profile.",
        )

    return ChartRecommendation(
        chart_type=ChartType.BAR_CLUSTERED,
        orientation="horizontal",
        primary_metric="count",
        sort_order="descending",
        highlight_rule="none",
        series_colors=[BAIN_COLORS["graphite"]],
        data_label_format="integer",
        data_label_position="outside_end",
        notes=f"Fallback recommendation for unhandled type {qtype}",
    )


def _recommend_single_select(
    result: SingleSelectResult, spec: QuestionSpec | None
) -> ChartRecommendation:
    return ChartRecommendation(
        chart_type=ChartType.COLUMN_STACKED,
        orientation="vertical",
        primary_metric="rate",
        sort_order="descending",
        highlight_rule="top_1",
        series_colors=[BAIN_COLORS["bain_red"], BAIN_COLORS["graphite"]],
        data_label_format="percent_integer",
        data_label_position="outside_end",
        notes="Single-select distribution. Stacked column per Bain practice.",
    )


def _recommend_multi_select(
    result: MultiSelectResult, spec: QuestionSpec | None
) -> ChartRecommendation:
    return ChartRecommendation(
        chart_type=ChartType.BAR_CLUSTERED,
        orientation="horizontal",
        primary_metric="selection_rate",
        sort_order="descending",
        highlight_rule="top_1_to_3",
        series_colors=[BAIN_COLORS["bain_red"], BAIN_COLORS["graphite"]],
        data_label_format="percent_integer",
        data_label_position="outside_end",
        notes="Multi-select: horizontal bar, top items highlighted. Never stack.",
    )


def _recommend_numeric_allocation(
    result: NumericResult, spec: QuestionSpec | None
) -> ChartRecommendation:
    return ChartRecommendation(
        chart_type=ChartType.BAR_CLUSTERED,
        orientation="horizontal",
        primary_metric="mean",
        sort_order="descending",
        highlight_rule="top_1",
        series_colors=[BAIN_COLORS["bain_red"], BAIN_COLORS["graphite"]],
        data_label_format="decimal_1",
        data_label_position="outside_end",
        notes="Numeric: horizontal bar on mean.",
    )


def _recommend_rank_order(
    result: RankOrderResult, spec: QuestionSpec | None
) -> ChartRecommendation:
    rank_colors = BAIN_SERIES_PALETTE[: result.K]

    return ChartRecommendation(
        chart_type=ChartType.BAR_STACKED,
        orientation="horizontal",
        primary_metric="counts_per_rank",
        sort_order="rank_1_descending",
        highlight_rule="none",
        series_colors=rank_colors,
        data_label_format="percent_integer",
        data_label_position="inside_base",
        notes=f"Rank-order: stacked bar, K={result.K} segments per option.",
    )


def _recommend_grid_rated(
    result: GridRatedResult, spec: QuestionSpec | None
) -> ChartRecommendation:
    n_columns = len(result.column_headers)

    if n_columns == 2 and result.show_delta:
        return ChartRecommendation(
            chart_type=ChartType.BAR_CLUSTERED,
            orientation="horizontal",
            primary_metric="means_per_column",
            sort_order="delta_descending",
            highlight_rule="none",
            series_colors=[BAIN_COLORS["bain_red"], BAIN_COLORS["graphite"]],
            data_label_format="decimal_1",
            data_label_position="outside_end",
            axis_min=0,
            axis_max=10,
            show_delta=True,
            delta_format="decimal_1",
            notes="2-entity rated grid: clustered bar sorted by delta.",
        )

    return ChartRecommendation(
        chart_type=ChartType.LINE,
        orientation=None,
        primary_metric="means_per_column",
        sort_order="preserve",
        highlight_rule="none",
        series_colors=BAIN_SERIES_PALETTE[:n_columns],
        data_label_format="decimal_1",
        data_label_position="right_end",
        axis_min=0,
        axis_max=10,
        notes=f"{n_columns}-entity rated grid: line chart preserving criterion order.",
    )


def _recommend_grid_binary_pivot(
    result: GridBinaryPivotResult, spec: QuestionSpec | None
) -> ChartRecommendation:
    return ChartRecommendation(
        chart_type=ChartType.HEATMAP_TABLE,
        orientation=None,
        primary_metric="pcts_per_column",
        sort_order="preserve",
        highlight_rule="none",
        series_colors=[BAIN_COLORS["bain_red"], BAIN_COLORS["white"]],
        data_label_format="percent_integer",
        data_label_position="inside_base",
        artifact_type="formatted_table",
        button_label_override="Generate formatted table",
        notes="2D binary pivot: formatted table with conditional cell coloring.",
    )
