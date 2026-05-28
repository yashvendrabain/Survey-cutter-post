"""Tests for screen-only chart type overrides in the Streamlit app."""

from __future__ import annotations

import importlib.util
from pathlib import Path
import unittest

from src.chart_recommender import ChartRecommendation, ChartType


APP_PATH = Path(__file__).resolve().parents[1] / "artifacts" / "survey-insight-engine" / "app.py"


def load_app_module():
    spec = importlib.util.spec_from_file_location("survey_app_chart_override", APP_PATH)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


class FakeStreamlit:
    def __init__(self, selected: ChartType | None = None) -> None:
        self.session_state: dict[str, object] = {}
        self.selected = selected
        self.last_selectbox: dict[str, object] | None = None

    def selectbox(self, label, options, format_func=None, index=0, key=None):
        selected = self.selected if self.selected is not None else options[index]
        self.session_state[key] = selected
        self.last_selectbox = {
            "label": label,
            "options": options,
            "format_func": format_func,
            "index": index,
            "key": key,
            "selected": selected,
        }
        return selected


def make_recommendation(chart_type: ChartType = ChartType.COLUMN_STACKED) -> ChartRecommendation:
    return ChartRecommendation(
        chart_type=chart_type,
        orientation="vertical",
        primary_metric="rate",
        sort_order="descending",
        highlight_rule="top_1",
        series_colors=["#CC0000"],
        data_label_format="percent_integer",
        data_label_position="outside_end",
    )


class TestChartTypeOverride(unittest.TestCase):
    def test_selecting_non_default_chart_type_updates_session_state(self) -> None:
        app_module = load_app_module()
        fake_app = FakeStreamlit(selected=ChartType.BAR_CLUSTERED)
        app_module.st = fake_app

        selected = app_module._render_chart_type_override_control(
            "Q1",
            ChartType.COLUMN_STACKED,
            "recommended",
        )

        self.assertEqual(selected, ChartType.BAR_CLUSTERED)
        self.assertEqual(fake_app.session_state["chart_type_override_Q1"], ChartType.BAR_CLUSTERED)
        self.assertEqual(fake_app.last_selectbox["label"], "Chart type")
        self.assertEqual(fake_app.last_selectbox["options"], app_module._chart_type_options())

    def test_render_path_receives_screen_override_type(self) -> None:
        app_module = load_app_module()
        recommendation = make_recommendation(ChartType.COLUMN_STACKED)

        screen_recommendation = app_module._chart_recommendation_for_screen(
            recommendation,
            ChartType.BAR_CLUSTERED,
        )

        self.assertEqual(screen_recommendation.chart_type, ChartType.BAR_CLUSTERED)
        self.assertEqual(recommendation.chart_type, ChartType.COLUMN_STACKED)
        app_text = APP_PATH.read_text(encoding="utf-8")
        self.assertIn("fig = render_chart(screen_recommendation, payload)", app_text)

    def test_switching_back_to_recommended_type_sets_override_equal_to_default(self) -> None:
        app_module = load_app_module()
        fake_app = FakeStreamlit(selected=ChartType.COLUMN_STACKED)
        fake_app.session_state["chart_type_override_Q1"] = ChartType.BAR_CLUSTERED
        app_module.st = fake_app

        selected = app_module._render_chart_type_override_control(
            "Q1",
            ChartType.COLUMN_STACKED,
            "recommended",
        )

        self.assertEqual(selected, ChartType.COLUMN_STACKED)
        self.assertEqual(fake_app.session_state["chart_type_override_Q1"], ChartType.COLUMN_STACKED)

    def test_ppttc_generator_still_uses_recommended_type_not_override(self) -> None:
        app_text = APP_PATH.read_text(encoding="utf-8")

        self.assertIn("payload = format_for_thinkcell(\n            result,\n            recommendation,", app_text)
        self.assertIn("recommendation=recommendation,", app_text)
        self.assertNotIn("recommendation=screen_recommendation", app_text)

    def test_label_formatter_is_human_readable(self) -> None:
        app_module = load_app_module()

        self.assertEqual(app_module._format_chart_type_label(ChartType.COLUMN_STACKED), "Stacked column")
        self.assertEqual(app_module._format_chart_type_label(ChartType.BAR_CLUSTERED), "Clustered bar")
        self.assertEqual(app_module._format_chart_type_label(ChartType.HEATMAP_TABLE), "Heatmap table")


if __name__ == "__main__":
    unittest.main()
