"""Tests for think-cell .ppttc generation."""

from __future__ import annotations

from pathlib import Path
import io
import json
import tempfile
import unittest
import zipfile

from src.chart_recommender import (
    BAIN_COLORS,
    ChartRecommendation,
    ChartType,
    recommend_chart,
)
from src.models import (
    DenominatorPolicy,
    GridBinaryPivotResult,
    GridBinaryPivotRow,
    GridRatedResult,
    GridRatedRow,
    MultiSelectResult,
    QuestionType,
)
from src.ppttc_generator import (
    build_ppttc_bundle,
    build_ppttc,
    resolve_chart_anchor,
)
from src.thinkcell_table_formatter import ThinkCellTablePayload, format_for_thinkcell
from tests.test_chart_recommender import make_single_select


def _tmp_template() -> Path:
    handle = tempfile.NamedTemporaryFile(suffix=".pptx", delete=False)
    handle.write(b"fake pptx template")
    handle.close()
    return Path(handle.name)


class TestPpttcGenerator(unittest.TestCase):
    def test_anchor_resolution_single_select(self):
        result = make_single_select()
        rec = recommend_chart(result)
        payload = format_for_thinkcell(result, rec)
        anchor, warnings = resolve_chart_anchor(rec, payload)
        self.assertEqual(anchor, "Chart_stacked_column")
        self.assertEqual(warnings, [])

    def test_anchor_resolution_bar_clustered_single_series(self):
        result = MultiSelectResult(
            question_id="Q2",
            question_type=QuestionType.MULTI_SELECT_BINARY,
            valid_n=100,
            missing_n=0,
            denominator_policy=DenominatorPolicy.VALID_RESPONSES,
            selections={
                "Q2r1": {"label": "A", "count": 60, "selection_rate": 0.6},
                "Q2r2": {"label": "B", "count": 40, "selection_rate": 0.4},
            },
            respondents_who_answered_any=100,
        )
        rec = recommend_chart(result)
        payload = format_for_thinkcell(result, rec)
        anchor, warnings = resolve_chart_anchor(rec, payload)
        self.assertEqual(anchor, "Chart_horizontal_bar")
        self.assertEqual(warnings, [])

    def test_anchor_resolution_bar_clustered_multi_series(self):
        result = GridRatedResult(
            question_id="Q30",
            question_type=QuestionType.GRID_RATED,
            valid_n=200,
            missing_n=0,
            denominator_policy=DenominatorPolicy.VALID_RESPONSES,
            question_text="Rate vendors",
            column_headers=["Winner", "Other"],
            rows=[
                GridRatedRow(
                    row_id="Q30r1",
                    row_label="Criterion",
                    means_per_column=[8.5, 7.2],
                    valid_n_per_column=[100, 100],
                    delta=1.3,
                )
            ],
            total_respondents=200,
            total_responses=200,
            show_delta=True,
        )
        rec = recommend_chart(result)
        payload = format_for_thinkcell(result, rec)
        anchor, warnings = resolve_chart_anchor(rec, payload)
        self.assertEqual(anchor, "Chart_clustered_horizontal")
        self.assertEqual(warnings, [])

    def test_heatmap_returns_none(self):
        result = GridBinaryPivotResult(
            question_id="Q26",
            question_type=QuestionType.GRID_BINARY_SELECT,
            valid_n=100,
            missing_n=0,
            denominator_policy=DenominatorPolicy.VALID_RESPONSES,
            question_text="Roles",
            column_headers=["Blocked", "Scored"],
            rows=[
                GridBinaryPivotRow(
                    row_id="Q26r1",
                    row_label="Future users",
                    counts_per_column=[5, 10],
                    pcts_per_column=[0.05, 0.10],
                )
            ],
            total_respondents=100,
            total_responses=15,
        )
        rec = recommend_chart(result)
        payload = format_for_thinkcell(result, rec)
        self.assertIsNone(build_ppttc(rec, payload, Path("missing.pptx"), "Q26"))

    def test_combo_uses_fallback_with_warning(self):
        rec = ChartRecommendation(
            chart_type=ChartType.COMBO,
            orientation="vertical",
            primary_metric="value",
            sort_order="preserve",
            highlight_rule="none",
            series_colors=[BAIN_COLORS["graphite"]],
            data_label_format="decimal_1",
            data_label_position="outside_end",
        )
        payload = ThinkCellTablePayload(
            headers=["A"],
            rows=[["Series", 1.0]],
            chart_type=ChartType.COMBO,
            title="Combo",
            source_line="Source",
        )
        artifact = build_ppttc(rec, payload, _tmp_template(), "Q1", "Combo chart")
        self.assertIsNotNone(artifact)
        assert artifact is not None
        self.assertEqual(artifact.chart_anchor, "Chart_clustered_column")
        self.assertTrue(any("COMBO" in warning for warning in artifact.warnings))

    def test_filename_sanitization(self):
        result = make_single_select()
        rec = recommend_chart(result)
        payload = format_for_thinkcell(result, rec)
        artifact = build_ppttc(
            rec,
            payload,
            _tmp_template(),
            "Q/12",
            'growth levers / "quoted" very long label that keeps going',
        )
        self.assertIsNotNone(artifact)
        assert artifact is not None
        self.assertTrue(artifact.filename.endswith(".ppttc"))
        self.assertNotIn("/", artifact.filename)
        self.assertNotIn('"', artifact.filename)
        self.assertTrue(artifact.filename.startswith("Q_12_"))
        label_segment = Path(artifact.filename).stem[len("Q_12_") :]
        self.assertLessEqual(len(label_segment), 40)

    def test_ppttc_json_is_valid_json(self):
        result = make_single_select()
        rec = recommend_chart(result)
        payload = format_for_thinkcell(result, rec)
        template = _tmp_template()
        artifact = build_ppttc(rec, payload, template, "Q1", "Question")
        self.assertIsNotNone(artifact)
        assert artifact is not None
        parsed = json.loads(artifact.content_bytes.decode("utf-8"))
        self.assertEqual(parsed[0]["template"], template.name)
        self.assertEqual(parsed[1]["data"][0]["name"], "Chart_stacked_column")
        self.assertIn("table", parsed[1]["data"][0])

    def test_ppttc_shape_matches_thinkcell_spec(self):
        result = make_single_select()
        rec = recommend_chart(result)
        payload = format_for_thinkcell(result, rec)
        template = _tmp_template()
        artifact = build_ppttc(rec, payload, template, "Q1", "test_q")
        assert artifact is not None
        parsed = json.loads(artifact.content_bytes)
        self.assertIsInstance(parsed, list)
        self.assertEqual(len(parsed), 2)
        self.assertIn("template", parsed[0])
        self.assertNotIn("data", parsed[0])
        self.assertIn("data", parsed[1])
        self.assertNotIn("template", parsed[1])
        self.assertEqual(parsed[0]["template"], template.name)
        self.assertEqual(parsed[1]["data"][0]["name"], "Chart_stacked_column")

    def test_blank_cells_serialize_as_empty_object(self):
        result = make_single_select()
        rec = recommend_chart(result)
        payload = format_for_thinkcell(result, rec)
        template = _tmp_template()
        artifact = build_ppttc(rec, payload, template, "Q1", "test_q")
        assert artifact is not None
        parsed = json.loads(artifact.content_bytes)
        table = parsed[1]["data"][0]["table"]
        self.assertEqual(table[0][0], {})
        for row in table:
            for cell in row:
                self.assertIsNotNone(cell, "blank cells must be {}, never null")

    def test_zip_bundle_contains_both_files(self):
        result = make_single_select()
        rec = recommend_chart(result)
        payload = format_for_thinkcell(result, rec)
        template = _tmp_template()
        artifact = build_ppttc(rec, payload, template, "Q1", "test_q")
        assert artifact is not None
        zip_bytes, zip_filename = build_ppttc_bundle(artifact, template)
        self.assertTrue(zip_filename.endswith("_thinkcell.zip"))
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as archive:
            names = set(archive.namelist())
            self.assertIn(template.name, names)
            self.assertTrue(any(name.endswith(".ppttc") for name in names))
