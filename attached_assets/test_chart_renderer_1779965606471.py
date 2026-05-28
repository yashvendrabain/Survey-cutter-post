"""Tests for Plotly chart rendering palette behavior."""

from __future__ import annotations

import unittest

from src.bain_palette import BAIN_PALETTE, get_hero_color, get_series_palette
from src.chart_recommender import ChartRecommendation, ChartType
from src.chart_renderer import go, render_chart
from src.thinkcell_table_formatter import ThinkCellTablePayload


def _recommendation(chart_type: ChartType, highlight_rule: str) -> ChartRecommendation:
    return ChartRecommendation(
        chart_type=chart_type,
        orientation="vertical" if chart_type is ChartType.COLUMN_STACKED else "horizontal",
        primary_metric="rate",
        sort_order="descending",
        highlight_rule=highlight_rule,
        series_colors=get_series_palette(5),
        data_label_format="percent_integer",
        data_label_position="outside_end",
    )


def _assert_figure_uses_only_bain_palette(testcase: unittest.TestCase, figure) -> None:
    allowed = set(BAIN_PALETTE.values())
    for trace in figure.data:
        marker = getattr(trace, "marker", None)
        marker_color = getattr(marker, "color", None)
        if marker_color is not None:
            colors = marker_color if isinstance(marker_color, (list, tuple)) else [marker_color]
            for color in colors:
                testcase.assertIn(color, allowed)
        line = getattr(trace, "line", None)
        line_color = getattr(line, "color", None)
        if line_color is not None:
            testcase.assertIn(line_color, allowed)
        colorscale = getattr(trace, "colorscale", None)
        if colorscale is not None:
            for _stop, color in colorscale:
                testcase.assertIn(color, allowed)
    layout = figure.layout
    testcase.assertIn(layout.paper_bgcolor, allowed)
    testcase.assertIn(layout.plot_bgcolor, allowed)
    testcase.assertIn(layout.font.color, allowed)
    testcase.assertIn(layout.title.font.color, allowed)
    testcase.assertIn(layout.xaxis.gridcolor, allowed)
    testcase.assertIn(layout.yaxis.gridcolor, allowed)


@unittest.skipIf(go is None, "Plotly is not installed in this environment.")
class TestChartRendererPalette(unittest.TestCase):
    def test_single_select_stacked_column_uses_bain_palette(self):
        payload = ThinkCellTablePayload(
            headers=["Respondents"],
            rows=[[f"Option {index}", 0.125] for index in range(8)],
            chart_type=ChartType.COLUMN_STACKED,
            title="Revenue distribution",
            source_line="Source: Test (N=100)",
        )
        figure = render_chart(_recommendation(ChartType.COLUMN_STACKED, "top_1"), payload)

        colors = [trace.marker.color for trace in figure.data]

        self.assertEqual(colors[0], get_hero_color())
        self.assertEqual(colors[1:5], get_series_palette(5, hero_index=0)[1:])
        _assert_figure_uses_only_bain_palette(self, figure)

    def test_multi_select_bar_highlights_top_three_and_uses_bain_palette(self):
        payload = ThinkCellTablePayload(
            headers=["Selection rate"],
            rows=[[f"Option {index}", 0.5 - index * 0.05] for index in range(8)],
            chart_type=ChartType.BAR_CLUSTERED,
            title="Multi select",
            source_line="Source: Test (N=100)",
        )
        figure = render_chart(_recommendation(ChartType.BAR_CLUSTERED, "top_1_to_3"), payload)

        colors = list(figure.data[0].marker.color)

        self.assertEqual(colors[:3], [get_hero_color()] * 3)
        self.assertTrue(set(colors[3:]).issubset(set(BAIN_PALETTE.values())))
        _assert_figure_uses_only_bain_palette(self, figure)

    def test_line_chart_uses_only_bain_palette(self):
        payload = ThinkCellTablePayload(
            headers=["A", "B", "C"],
            rows=[["Metric 1", 1, 2, 3], ["Metric 2", 2, 3, 4]],
            chart_type=ChartType.LINE,
            title="Line",
            source_line="Source: Test",
        )
        figure = render_chart(_recommendation(ChartType.LINE, "none"), payload)

        _assert_figure_uses_only_bain_palette(self, figure)

    def test_heatmap_uses_only_bain_palette(self):
        payload = ThinkCellTablePayload(
            headers=["A", "B"],
            rows=[["Row 1", 0.2, 0.8], ["Row 2", 0.4, 0.6]],
            chart_type=ChartType.HEATMAP_TABLE,
            title="Heatmap",
            source_line="Source: Test",
        )
        figure = render_chart(_recommendation(ChartType.HEATMAP_TABLE, "none"), payload)

        _assert_figure_uses_only_bain_palette(self, figure)


if __name__ == "__main__":
    unittest.main()
