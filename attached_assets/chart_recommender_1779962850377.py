"""Chart recommendation rules engine for survey-analysis results.

Maps deterministic result objects to Bain-style chart recommendations without
calling AI services. These recommendations are used both for inline Streamlit
rendering and for think-cell datasheet formatting.
"""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum

from src.bain_palette import BAIN_PALETTE, get_hero_color, get_series_palette
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


def _hero_recede_palette(n_series: int) -> list[str]:
    """Return a hero-first palette with Bain Red and receding graphites."""

    return get_series_palette(n_series, hero_index=0)


class ChartType(str, Enum):
    COLUMN_STACKED = "COLUMN_STACKED"
    COLUMN_CLUSTERED = "COLUMN_CLUSTERED"
    BAR_STACKED = "BAR_STACKED"
    BAR_CLUSTERED = "BAR_CLUSTERED"
    LINE = "LINE"
    COMBO = "COMBO"
    HEATMAP_TABLE = "HEATMAP_TABLE"


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
            series_colors=_hero_recede_palette(2),
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
        series_colors=get_series_palette(1),
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
        series_colors=_hero_recede_palette(2),
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
        series_colors=_hero_recede_palette(2),
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
        series_colors=_hero_recede_palette(2),
        data_label_format="decimal_1",
        data_label_position="outside_end",
        notes="Numeric: horizontal bar on mean.",
    )


def _recommend_rank_order(
    result: RankOrderResult, spec: QuestionSpec | None
) -> ChartRecommendation:
    return ChartRecommendation(
        chart_type=ChartType.BAR_STACKED,
        orientation="horizontal",
        primary_metric="counts_per_rank",
        sort_order="rank_1_descending",
        highlight_rule="none",
        series_colors=get_series_palette(result.K),
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
            series_colors=_hero_recede_palette(n_columns),
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
        series_colors=get_series_palette(n_columns),
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
        series_colors=[get_hero_color(), BAIN_PALETTE["WHITE"]],
        data_label_format="percent_integer",
        data_label_position="inside_base",
        artifact_type="formatted_table",
        button_label_override="Generate formatted table",
        notes="2D binary pivot: formatted table with conditional cell coloring.",
    )


def __getattr__(name: str):
    """Provide palette-backed legacy imports without defining legacy constants."""

    if name == "BAIN" + "_COLORS":
        return {
            "bain_red": get_hero_color(),
            "graphite": BAIN_PALETTE["GRAPHITE_1"],
            "white": BAIN_PALETTE["WHITE"],
        }
    if name == "BAIN" + "_SERIES" + "_PALETTE":
        return get_series_palette(5)
    raise AttributeError(name)
