"""Render recommended charts as Plotly figures for the Streamlit UI."""

from __future__ import annotations

try:
    import plotly.graph_objects as go
except ModuleNotFoundError:  # pragma: no cover - optional UI dependency.
    go = None  # type: ignore[assignment]

from src.chart_recommender import BAIN_COLORS, ChartRecommendation, ChartType
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
    raise ValueError(f"Unsupported chart type: {chart_type}")


def _bain_layout(title: str, source: str, height: int = 450) -> dict:
    return {
        "title": {
            "text": title,
            "font": {"family": "Arial", "size": 14, "color": BAIN_COLORS["graphite"]},
            "x": 0.0,
            "xanchor": "left",
        },
        "font": {"family": "Arial", "size": 11, "color": BAIN_COLORS["graphite"]},
        "plot_bgcolor": BAIN_COLORS["white"],
        "paper_bgcolor": BAIN_COLORS["white"],
        "height": height,
        "margin": {"l": 90, "r": 80, "t": 60, "b": 80},
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
                    "color": BAIN_COLORS["graphite"],
                },
            }
        ],
    }


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


def _render_column_stacked(
    recommendation: ChartRecommendation, payload: ThinkCellTablePayload
) -> go.Figure:
    fig = go.Figure()
    for index, row in enumerate(payload.rows):
        option = str(row[0])
        value = float(row[1])
        color = (
            BAIN_COLORS["bain_red"]
            if recommendation.highlight_rule == "top_1" and index == 0
            else BAIN_COLORS["graphite"]
        )
        fig.add_trace(
            go.Bar(
                name=option,
                x=["Respondents"],
                y=[_percent_value(value)],
                marker_color=color,
                text=_format_value(value, recommendation.data_label_format),
                textposition="inside",
            )
        )
    fig.update_layout(
        barmode="stack",
        showlegend=True,
        yaxis={"title": "% of respondents", "range": [0, 100]},
        xaxis={"showticklabels": False},
        **_bain_layout(payload.title, payload.source_line),
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
        colors = []
        for index in range(len(options)):
            if recommendation.highlight_rule == "top_1" and index == 0:
                colors.append(BAIN_COLORS["bain_red"])
            elif recommendation.highlight_rule == "top_1_to_3" and index < 3:
                colors.append(BAIN_COLORS["bain_red"])
            else:
                colors.append(BAIN_COLORS["graphite"])
        fig.add_trace(
            go.Bar(
                y=options,
                x=display_values,
                orientation="h",
                marker_color=colors,
                text=[
                    _format_value(value, recommendation.data_label_format)
                    for value in values
                ],
                textposition="outside",
            )
        )
    else:
        for col_index, header in enumerate(payload.headers):
            values = [float(row[col_index + 1]) for row in payload.rows]
            color = (
                recommendation.series_colors[col_index]
                if col_index < len(recommendation.series_colors)
                else BAIN_COLORS["graphite"]
            )
            fig.add_trace(
                go.Bar(
                    name=header,
                    y=options,
                    x=values,
                    orientation="h",
                    marker_color=color,
                    text=[f"{value:.1f}" for value in values],
                    textposition="outside",
                )
            )
        fig.update_layout(barmode="group")
    fig.update_layout(
        showlegend=len(payload.headers) > 1,
        yaxis={"autorange": "reversed"},
        xaxis={"showgrid": True, "gridcolor": "#F0F0F0"},
        **_bain_layout(payload.title, payload.source_line, max(350, 30 * len(options) + 150)),
    )
    return fig


def _render_bar_stacked(
    recommendation: ChartRecommendation, payload: ThinkCellTablePayload
) -> go.Figure:
    fig = go.Figure()
    options = payload.headers
    for row_index, row in enumerate(payload.rows):
        color = (
            recommendation.series_colors[row_index]
            if row_index < len(recommendation.series_colors)
            else BAIN_COLORS["graphite"]
        )
        fig.add_trace(
            go.Bar(
                name=str(row[0]),
                y=options,
                x=row[1:],
                orientation="h",
                marker_color=color,
                text=row[1:],
                textposition="inside",
            )
        )
    fig.update_layout(
        barmode="stack",
        showlegend=True,
        yaxis={"autorange": "reversed"},
        **_bain_layout(payload.title, payload.source_line, max(350, 30 * len(options) + 150)),
    )
    return fig


def _render_column_clustered(
    recommendation: ChartRecommendation, payload: ThinkCellTablePayload
) -> go.Figure:
    fig = go.Figure()
    categories = [str(row[0]) for row in payload.rows]
    for col_index, header in enumerate(payload.headers):
        values = [float(row[col_index + 1]) for row in payload.rows]
        color = (
            recommendation.series_colors[col_index]
            if col_index < len(recommendation.series_colors)
            else BAIN_COLORS["graphite"]
        )
        fig.add_trace(
            go.Bar(
                name=header,
                x=categories,
                y=values,
                marker_color=color,
                text=[f"{value:.1f}" for value in values],
                textposition="outside",
            )
        )
    fig.update_layout(
        barmode="group",
        showlegend=True,
        **_bain_layout(payload.title, payload.source_line),
    )
    return fig


def _render_line(
    recommendation: ChartRecommendation, payload: ThinkCellTablePayload
) -> go.Figure:
    fig = go.Figure()
    criteria = [str(row[0]) for row in payload.rows]
    for col_index, header in enumerate(payload.headers):
        values = [float(row[col_index + 1]) for row in payload.rows]
        color = (
            recommendation.series_colors[col_index]
            if col_index < len(recommendation.series_colors)
            else BAIN_COLORS["graphite"]
        )
        fig.add_trace(
            go.Scatter(
                name=header,
                x=criteria,
                y=values,
                mode="lines+markers+text",
                line={"color": color, "width": 2},
                marker={"size": 6},
                text=[f"{value:.1f}" for value in values],
                textposition="top center",
            )
        )
    fig.update_layout(
        showlegend=True,
        yaxis={"range": [recommendation.axis_min or 0, recommendation.axis_max or 10]},
        xaxis={"tickangle": -45},
        **_bain_layout(payload.title, payload.source_line, 500),
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
            colorscale=[[0, BAIN_COLORS["white"]], [1, BAIN_COLORS["bain_red"]]],
            text=text,
            texttemplate="%{text}",
            textfont={"size": 11, "family": "Arial"},
            showscale=False,
        )
    )
    fig.update_layout(
        yaxis={"autorange": "reversed"},
        **_bain_layout(payload.title, payload.source_line, max(350, 35 * len(row_labels) + 150)),
    )
    return fig
