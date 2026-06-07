from __future__ import annotations

from pathlib import Path
import unittest


APP_PATH = Path(__file__).resolve().parents[1] / "artifacts" / "survey-insight-engine" / "app.py"


class TestOutcomeLaggardUi(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app_text = APP_PATH.read_text(encoding="utf-8")

    def test_laggard_cohort_ui_copy_is_present(self) -> None:
        self.assertIn("Laggard cohort", self.app_text)
        self.assertIn("Laggard option values", self.app_text)
        self.assertIn("Laggard threshold", self.app_text)

    def test_override_outcome_checkbox_and_dropdown_are_present(self) -> None:
        self.assertIn("Override outcome question for laggards", self.app_text)
        self.assertIn("laggard_outcome_variable_selector", self.app_text)

    def test_grid_override_subquestion_dropdown_is_present(self) -> None:
        self.assertIn("Laggard grid sub-question", self.app_text)
        self.assertIn("laggard_outcome_sub_question_selector", self.app_text)

    def test_outcome_segmented_download_button_is_present(self) -> None:
        self.assertIn("Outcome Segmented Workbook (Winners vs Laggards)", self.app_text)
        self.assertIn("export_winners_vs_laggards_workbook", self.app_text)

    def test_protective_app_container_and_run_pipeline_import_position(self) -> None:
        self.assertIn("with app.container():", self.app_text)
        run_pipeline = self.app_text.index("def _run_pipeline(")
        body = self.app_text[run_pipeline:self.app_text.index("status.update", run_pipeline)]
        self.assertIn("app = _require_streamlit()", body)

    def test_manual_cross_cut_routes_numeric_metric_pairs(self) -> None:
        start = self.app_text.index("def _render_manual_cross_cut(")
        end = self.app_text.index("# ---------------------------------------------------------------------------", start)
        body = self.app_text[start:end]

        self.assertIn("QuestionType.DIRECT_NUMERIC", body)
        self.assertIn("QuestionType.NUMERIC_ALLOCATION", body)
        self.assertIn("QuestionType.GRID_RATED", body)
        self.assertIn("QuestionType.NPS", body)
        self.assertIn("QuestionType.MULTI_SELECT_BINARY", body)
        self.assertIn("QuestionType.RANK_ORDER", body)
        self.assertIn("QuestionType.GRID_BINARY_SELECT", body)
        self.assertIn("QuestionType.DEMOGRAPHIC_OR_SEGMENT", body)
        self.assertIn("QuestionType.GRID_SINGLE_SELECT", body)
        self.assertIn("analysis_type = AnalysisType.GROUP_COMPARISON", body)
        self.assertIn("source_ids = (segment_id, metric_id)", body)
        self.assertIn("source_question_ids=source_ids", body)
        self.assertIn(
            "GROUP_COMPARISON does not yet support NUMERIC_ALLOCATION metrics",
            body,
        )
        self.assertIn("CROSS_TAB requires two categorical questions", body)
        self.assertIn("EXPECTED_VS_REALIZED requires two direct numeric questions", body)

    def test_pipeline_exports_reconciled_schema(self) -> None:
        run_pipeline = self.app_text.index("def _run_pipeline(")
        body = self.app_text[
            run_pipeline:self.app_text.index("app.session_state[\"manual_cohort_input\"]", run_pipeline)
        ]

        self.assertIn("classified_schema = _run_with_status_heartbeat", body)
        self.assertIn(
            "schema = reconcile_multiselect_value_subtypes(classified_schema, dataframe)",
            body,
        )
        compute_index = body.index("work=lambda: compute_single_cuts(schema, dataframe, log)")
        export_index = body.index("work=lambda: export_single_cuts(")
        self.assertLess(compute_index, export_index)
        export_body = body[export_index:body.index("complete_label=\"Stage 5/5", export_index)]
        self.assertIn("schema=schema", export_body)

    def test_wizard_pipeline_overrides_preserve_current_schema(self) -> None:
        start = self.app_text.index("def _wizard_pipeline_overrides(")
        end = self.app_text.index("def _wizard_rerun", start)
        body = self.app_text[start:end]

        self.assertIn("apply_wizard_schema_overrides(", body)
        self.assertIn("schema,", body)
        self.assertIn("themes_from_wizard_assignments(schema, assignments)", body)
        self.assertNotIn('app.session_state.get("wiz_schema_override") or schema', body)


if __name__ == "__main__":
    unittest.main()
