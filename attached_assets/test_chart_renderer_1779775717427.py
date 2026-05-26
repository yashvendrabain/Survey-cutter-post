"""Tests for Plotly chart rendering palette behavior."""

from __future__ import annotations

import unittest

from src.chart_recommender import (
    BAIN_COLORS,
    BAIN_SERIES_PALETTE,
    ChartRecommendation,
    ChartType,
)
from src.chart_renderer import go, render_chart
from src.thinkcell_table_formatter import ThinkCellTablePayload


def _recommendation(chart_type: ChartType, highlight_rule: str) -> ChartRecommendation:
    return ChartRecommendation(
        chart_type=chart_type,
        orientation="vertical" if chart_type is ChartType.COLUMN_STACKED else "horizontal",
        primary_metric="rate",
        sort_order="descending",
        highlight_rule=highlight_rule,
        series_colors=[BAIN_COLORS["bain_red"], *BAIN_SERIES_PALETTE],
        data_label_format="percent_integer",
        data_label_position="outside_end",
    )


@unittest.skipIf(go is None, "Plotly is not installed in this environment.")
class TestChartRendererPalette(unittest.TestCase):
    def test_single_select_stacked_column_uses_distinct_segment_colors(self):
        payload = ThinkCellTablePayload(
            headers=["Respondents"],
            rows=[[f"Option {index}", 0.125] for index in range(8)],
            chart_type=ChartType.COLUMN_STACKED,
            title="Revenue distribution",
            source_line="Source: Test (N=100)",
        )
        figure = render_chart(_recommendation(ChartType.COLUMN_STACKED, "top_1"), payload)

        colors = [trace.marker.color for trace in figure.data]

        self.assertEqual(colors[0], BAIN_COLORS["bain_red"])
        self.assertEqual(colors[1:], [BAIN_SERIES_PALETTE[index % len(BAIN_SERIES_PALETTE)] for index in range(7)])
        self.assertEqual(len(set(colors)), 6)

    def test_multi_select_bar_highlights_top_three_and_cycles_remaining_palette(self):
        payload = ThinkCellTablePayload(
            headers=["Selection rate"],
            rows=[[f"Option {index}", 0.5 - index * 0.05] for index in range(8)],
            chart_type=ChartType.BAR_CLUSTERED,
            title="Multi select",
            source_line="Source: Test (N=100)",
        )
        figure = render_chart(_recommendation(ChartType.BAR_CLUSTERED, "top_1_to_3"), payload)

        colors = list(figure.data[0].marker.color)

        self.assertEqual(colors[:3], [BAIN_COLORS["bain_red"]] * 3)
        self.assertEqual(colors[3:], [BAIN_SERIES_PALETTE[index % len(BAIN_SERIES_PALETTE)] for index in range(5)])
        self.assertEqual(len(set(colors)), 6)


if __name__ == "__main__":
    unittest.main()
