"""Tests for screen-only chart type overrides in the Streamlit app."""

from __future__ import annotations

import importlib.util
from pathlib import Path
import unittest

from src.chart_recommender import ChartRecommendation, ChartType
from src.models import InsightResult


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

    def selectbox(self, label, options, format_func=None, index=None, key=None):
        if self.selected is not None:
            selected = self.selected
        elif key in self.session_state:
            selected = self.session_state[key]
        elif index is not None:
            selected = options[index]
        else:
            selected = options[0]
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

    def test_widget_state_is_source_of_truth_after_previous_override(self) -> None:
        app_module = load_app_module()
        fake_app = FakeStreamlit()
        fake_app.session_state["chart_type_override_Q1"] = ChartType.COLUMN_STACKED
        fake_app.session_state["chart_type_override_Q1_recommended"] = ChartType.BAR_CLUSTERED
        app_module.st = fake_app

        selected = app_module._render_chart_type_override_control(
            "Q1",
            ChartType.COLUMN_STACKED,
            "recommended",
        )

        self.assertEqual(selected, ChartType.BAR_CLUSTERED)
        self.assertEqual(fake_app.session_state["chart_type_override_Q1"], ChartType.BAR_CLUSTERED)
        self.assertIsNone(fake_app.last_selectbox["index"])

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

    def test_outcome_diff_insight_with_wrong_numbers_uses_python_fallback(self) -> None:
        app_module = load_app_module()
        payload = {
            "option_label": "GenAI tools",
            "winner_rate": 0.60,
            "loser_rate": 0.464286,
            "lift": 1.2923,
            "rate_gap": 0.135714,
            "winner_label": "Winner",
            "laggard_label": "Laggard",
        }
        generated = InsightResult(
            title="",
            insight="Winners are 9.9x more likely (99% vs 12%)",
        )

        validated = app_module._validated_outcome_diff_insight(generated, payload)

        self.assertTrue(validated.was_template)
        self.assertEqual(
            validated.insight,
            "Winners are 1.3\u00d7 more likely than Laggards "
            "to select \u201cGenAI tools\u201d (60% vs 46%)",
        )

    def test_outcome_diff_insight_framing_rejects_magnitude_phrasing(self) -> None:
        app_module = load_app_module()

        self.assertFalse(
            app_module._outcome_diff_insight_framing_is_safe(
                "Winners allocate 13.5% of their budget to sales and marketing "
                "technology vs 2.1% of Laggards"
            )
        )
        self.assertFalse(
            app_module._outcome_diff_insight_framing_is_safe(
                "Winners expect 20.7% growth in revenue vs 7.7% of Laggards"
            )
        )

    def test_outcome_diff_insight_with_allocation_framing_uses_fallback(self) -> None:
        app_module = load_app_module()
        payload = {
            "option_label": "sales and marketing technology",
            "winner_rate": 0.135,
            "loser_rate": 0.021,
            "lift": 6.428,
            "rate_gap": 0.114,
            "winner_label": "Winners",
            "laggard_label": "Laggards",
        }
        generated = InsightResult(
            title="",
            insight=(
                "Winners allocate 13.5% of their budget to sales and marketing "
                "technology vs 2.1% of Laggards"
            ),
        )

        validated = app_module._validated_outcome_diff_insight(generated, payload)

        self.assertTrue(validated.was_template)
        self.assertEqual(
            validated.insight,
            "Winners are 6.4\u00d7 more likely than Laggards to select "
            "\u201csales and marketing technology\u201d (14% vs 2%)",
        )

    def test_outcome_diff_insight_with_growth_framing_uses_fallback(self) -> None:
        app_module = load_app_module()
        payload = {
            "option_label": "20%+ growth expectation",
            "winner_rate": 0.207,
            "loser_rate": 0.077,
            "lift": 2.688,
            "rate_gap": 0.130,
            "winner_label": "Winner",
            "laggard_label": "Laggard",
        }
        generated = InsightResult(
            title="",
            insight="Winners expect 20.7% growth in revenue vs 7.7% of Laggards",
        )

        validated = app_module._validated_outcome_diff_insight(generated, payload)

        self.assertTrue(validated.was_template)
        self.assertEqual(
            validated.insight,
            "Winners are 2.7\u00d7 more likely than Laggards to select "
            "\u201c20%+ growth expectation\u201d (21% vs 8%)",
        )

    def test_outcome_diff_insight_with_matching_numbers_is_kept(self) -> None:
        app_module = load_app_module()
        payload = {
            "option_label": "GenAI tools",
            "winner_rate": 0.60,
            "loser_rate": 0.464286,
            "lift": 1.2923,
            "rate_gap": 0.135714,
            "winner_label": "Winners",
            "laggard_label": "Laggards",
        }
        generated = InsightResult(
            title="",
            insight="Winners are 1.29x more likely to select GenAI tools (60.0% vs 46.4%).",
        )

        validated = app_module._validated_outcome_diff_insight(generated, payload)

        self.assertIs(validated, generated)

    def test_outcome_diff_insight_with_safe_more_likely_framing_is_kept(self) -> None:
        app_module = load_app_module()
        payload = {
            "option_label": "dedicated pricing team",
            "winner_rate": 0.54,
            "loser_rate": 0.23,
            "lift": 2.3478,
            "rate_gap": 0.31,
            "winner_label": "Winners",
            "laggard_label": "Laggards",
        }
        generated = InsightResult(
            title="",
            insight=(
                "Winners are 2.35x more likely to select dedicated pricing "
                "team (54.0% vs 23.0%)."
            ),
        )

        validated = app_module._validated_outcome_diff_insight(generated, payload)

        self.assertIs(validated, generated)


if __name__ == "__main__":
    unittest.main()
