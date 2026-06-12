"""Render recommended charts as Plotly figures for the Streamlit UI."""

from __future__ import annotations

try:
    import plotly.graph_objects as go
except ModuleNotFoundError:  # pragma: no cover - optional UI dependency.
    go = None  # type: ignore[assignment]

from src.bain_palette import BAIN_PALETTE, get_hero_color, get_series_palette
from src.chart_recommender import ChartRecommendation, ChartType
from src.thinkcell_table_formatter import ThinkCellTablePayload


def render_chart(
    recommendation: ChartRecommendation,
    table_payload: ThinkCellTablePayload,
) -> go.Figure:
    """Build a Plotly figure matching the recommendation and Bain styling."""

    if go is None:
        raise RuntimeError("Plotly is not installed; inline chart rendering is unavailable.")

    chart_type = recommendation.chart_type
    if chart_type is ChartType.COLUMN_STACKED:
        return _render_column_stacked(recommendation, table_payload)
    if chart_type is ChartType.COLUMN_CLUSTERED:
        return _render_column_clustered(recommendation, table_payload)
    if chart_type is ChartType.BAR_STACKED:
        return _render_bar_stacked(recommendation, table_payload)
    if chart_type is ChartType.BAR_CLUSTERED:
        return _render_bar_clustered(recommendation, table_payload)
    if chart_type is ChartType.LINE:
        return _render_line(recommendation, table_payload)
    if chart_type is ChartType.HEATMAP_TABLE:
        return _render_heatmap(recommendation, table_payload)
    if chart_type is ChartType.COMBO:
        return _render_combo(recommendation, table_payload)
    raise ValueError(f"Unsupported chart type: {chart_type}")


def _bain_layout(title: str, source: str, height: int = 450) -> dict:
    axis = {
        "linecolor": BAIN_PALETTE["GRAPHITE_3"],
        "tickcolor": BAIN_PALETTE["GRAPHITE_3"],
        "tickfont": {"family": "Arial", "size": 11, "color": BAIN_PALETTE["GRAPHITE_1"]},
        "titlefont": {"family": "Arial", "size": 11, "color": BAIN_PALETTE["GRAPHITE_1"]},
        "gridcolor": BAIN_PALETTE["GRAPHITE_5"],
        "zerolinecolor": BAIN_PALETTE["GRAPHITE_5"],
    }
    return {
        "title": {
            "text": title,
            "font": {"family": "Arial", "size": 14, "color": BAIN_PALETTE["BLACK"]},
            "x": 0.0,
            "xanchor": "left",
            "y": 0.9,
            "yanchor": "top",
        },
        "font": {"family": "Arial", "size": 11, "color": BAIN_PALETTE["GRAPHITE_1"]},
        "plot_bgcolor": BAIN_PALETTE["WHITE"],
        "paper_bgcolor": BAIN_PALETTE["WHITE"],
        "xaxis": dict(axis),
        "yaxis": dict(axis),
        "height": height,
        "margin": {"l": 90, "r": 80, "t": 96, "b": 80},
        "annotations": [
            {
                "text": source,
                "showarrow": False,
                "xref": "paper",
                "yref": "paper",
                "x": 0,
                "y": -0.15,
                "xanchor": "left",
                "font": {
                    "family": "Arial",
                    "size": 8,
                    "color": BAIN_PALETTE["GRAPHITE_2"],
                },
            }
        ],
    }


def _merge_axis(base: dict, axis_overrides: dict | None = None) -> dict:
    merged = dict(base)
    if axis_overrides:
        merged.update(axis_overrides)
    return merged


def _apply_bain_layout(
    fig: go.Figure,
    title: str,
    source: str,
    height: int = 450,
    xaxis: dict | None = None,
    yaxis: dict | None = None,
    **kwargs: object,
) -> None:
    layout = _bain_layout(title, source, height)
    layout["xaxis"] = _merge_axis(layout["xaxis"], xaxis)
    layout["yaxis"] = _merge_axis(layout["yaxis"], yaxis)
    layout.update(kwargs)
    fig.update_layout(**layout)


def _percent_value(value: float | int) -> float:
    return float(value) * 100 if 0 <= float(value) <= 1 else float(value)


def _format_value(value: float | int, label_format: str) -> str:
    numeric = _percent_value(value) if label_format.startswith("percent") else float(value)
    if label_format in {"percent_integer", "integer"}:
        suffix = "%" if label_format.startswith("percent") else ""
        return f"{numeric:.0f}{suffix}"
    if label_format == "percent_decimal_1":
        return f"{numeric:.1f}%"
    return f"{numeric:.1f}"


def _hero_index_for(recommendation: ChartRecommendation, n_items: int) -> int | None:
    if n_items <= 0:
        return None
    if recommendation.highlight_rule in {"top_1", "top_1_to_3"}:
        return 0
    return None


def _single_series_colors(recommendation: ChartRecommendation, n_items: int) -> list[str]:
    if recommendation.highlight_rule == "top_1_to_3":
        colors = get_series_palette(n_items, hero_index=0)
        for index in range(1, min(3, n_items)):
            colors[index] = get_hero_color()
        return colors
    return get_series_palette(n_items, hero_index=_hero_index_for(recommendation, n_items))


def _render_column_stacked(
    recommendation: ChartRecommendation, payload: ThinkCellTablePayload
) -> go.Figure:
    fig = go.Figure()
    colors = _single_series_colors(recommendation, len(payload.rows))
    for index, row in enumerate(payload.rows):
        value = float(row[1])
        fig.add_trace(
            go.Bar(
                name=str(row[0]),
                x=["Respondents"],
                y=[_percent_value(value)],
                marker_color=colors[index],
                text=_format_value(value, recommendation.data_label_format),
                textposition="inside",
            )
        )
    _apply_bain_layout(
        fig,
        payload.title,
        payload.source_line,
        barmode="stack",
        showlegend=True,
        yaxis={"title": "% of respondents", "range": [0, 100]},
        xaxis={"showticklabels": False},
    )
    return fig


def _render_bar_clustered(
    recommendation: ChartRecommendation, payload: ThinkCellTablePayload
) -> go.Figure:
    fig = go.Figure()
    options = [str(row[0]) for row in payload.rows]
    if len(payload.headers) == 1:
        values = [float(row[1]) for row in payload.rows]
        display_values = [
            _percent_value(value)
            if recommendation.data_label_format.startswith("percent")
            else value
            for value in values
        ]
        fig.add_trace(
            go.Bar(
                y=options,
                x=display_values,
                orientation="h",
                marker_color=_single_series_colors(recommendation, len(options)),
                text=[
                    _format_value(value, recommendation.data_label_format)
                    for value in values
                ],
                textposition="outside",
            )
        )
    else:
        colors = get_series_palette(len(payload.headers))
        for col_index, header in enumerate(payload.headers):
            values = [float(row[col_index + 1]) for row in payload.rows]
            fig.add_trace(
                go.Bar(
                    name=header,
                    y=options,
                    x=values,
                    orientation="h",
                    marker_color=colors[col_index],
                    text=[f"{value:.1f}" for value in values],
                    textposition="outside",
                )
            )
    _apply_bain_layout(
        fig,
        payload.title,
        payload.source_line,
        max(350, 30 * len(options) + 150),
        barmode="group",
        showlegend=len(payload.headers) > 1,
        yaxis={"autorange": "reversed"},
        xaxis={"showgrid": True},
    )
    return fig


def _render_bar_stacked(
    recommendation: ChartRecommendation, payload: ThinkCellTablePayload
) -> go.Figure:
    fig = go.Figure()
    options = payload.headers
    colors = get_series_palette(len(payload.rows))
    for row_index, row in enumerate(payload.rows):
        fig.add_trace(
            go.Bar(
                name=str(row[0]),
                y=options,
                x=row[1:],
                orientation="h",
                marker_color=colors[row_index],
                text=row[1:],
                textposition="inside",
            )
        )
    _apply_bain_layout(
        fig,
        payload.title,
        payload.source_line,
        max(350, 30 * len(options) + 150),
        barmode="stack",
        showlegend=True,
        yaxis={"autorange": "reversed"},
    )
    return fig


def _render_column_clustered(
    recommendation: ChartRecommendation, payload: ThinkCellTablePayload
) -> go.Figure:
    fig = go.Figure()
    categories = [str(row[0]) for row in payload.rows]
    colors = get_series_palette(len(payload.headers))
    for col_index, header in enumerate(payload.headers):
        values = [float(row[col_index + 1]) for row in payload.rows]
        fig.add_trace(
            go.Bar(
                name=header,
                x=categories,
                y=values,
                marker_color=colors[col_index],
                text=[f"{value:.1f}" for value in values],
                textposition="outside",
            )
        )
    _apply_bain_layout(
        fig,
        payload.title,
        payload.source_line,
        barmode="group",
        showlegend=True,
    )
    return fig


def _render_line(
    recommendation: ChartRecommendation, payload: ThinkCellTablePayload
) -> go.Figure:
    fig = go.Figure()
    criteria = [str(row[0]) for row in payload.rows]
    colors = get_series_palette(len(payload.headers))
    for col_index, header in enumerate(payload.headers):
        values = [float(row[col_index + 1]) for row in payload.rows]
        fig.add_trace(
            go.Scatter(
                name=header,
                x=criteria,
                y=values,
                mode="lines+markers+text",
                line={"color": colors[col_index], "width": 2},
                marker={"size": 6, "color": colors[col_index]},
                text=[f"{value:.1f}" for value in values],
                textposition="top center",
            )
        )
    _apply_bain_layout(
        fig,
        payload.title,
        payload.source_line,
        500,
        showlegend=True,
        yaxis={"range": [recommendation.axis_min or 0, recommendation.axis_max or 10]},
        xaxis={"tickangle": -45},
    )
    return fig


def _render_heatmap(
    recommendation: ChartRecommendation, payload: ThinkCellTablePayload
) -> go.Figure:
    headers = payload.headers
    row_labels = [str(row[0]) for row in payload.rows]
    z_values = [
        [float(row[col_index + 1]) for col_index in range(len(headers))]
        for row in payload.rows
    ]
    text = [
        [_format_value(value, recommendation.data_label_format) for value in row]
        for row in z_values
    ]
    fig = go.Figure(
        data=go.Heatmap(
            z=z_values,
            x=headers,
            y=row_labels,
            colorscale=[
                [0, BAIN_PALETTE["WHITE"]],
                [1, get_hero_color()],
            ],
            text=text,
            texttemplate="%{text}",
            textfont={"size": 11, "family": "Arial", "color": BAIN_PALETTE["BLACK"]},
            showscale=False,
        )
    )
    _apply_bain_layout(
        fig,
        payload.title,
        payload.source_line,
        max(350, 35 * len(row_labels) + 150),
        yaxis={"autorange": "reversed"},
    )
    return fig


def _render_combo(
    recommendation: ChartRecommendation, payload: ThinkCellTablePayload
) -> go.Figure:
    return _render_column_clustered(recommendation, payload)
