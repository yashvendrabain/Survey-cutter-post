from __future__ import annotations

from datetime import datetime, timezone
import math
import os
import unittest
from unittest.mock import patch

import numpy as np
import pandas as pd

from src.calculation_log import CalculationLog
import src.hypothesis_validator as hv
from src.hypothesis_validator import (
    _check_sample_size_floor,
    _classify_question_role,
    _compute_cohen_d,
    _compute_eta_squared,
    _compute_pearson_ci,
    _run_anova,
    _run_chi_square,
    _run_pearson,
    _run_spearman,
    _run_welch_t,
    _select_test,
    _verdict_from_statistic,
    explain_verdict,
    parse_freetext_hypothesis,
    validate_hypothesis,
)
from src.models import (
    DenominatorPolicy,
    HypothesisGroup,
    HypothesisSpec,
    HypothesisStatistic,
    QuestionSpec,
    QuestionType,
    SurveySchema,
)


def _question(
    qid: str,
    qtype: QuestionType,
    raw_columns: tuple[str, ...] | None = None,
    option_map: dict[int | str, str] | None = None,
    **kwargs,
) -> QuestionSpec:
    grid_labels = None
    if qtype in {
        QuestionType.GRID_SINGLE_SELECT,
        QuestionType.GRID_RATED,
        QuestionType.GRID_BINARY_SELECT,
    }:
        grid_labels = {column: column for column in (raw_columns or (qid,))}
    return QuestionSpec(
        question_id=qid,
        canonical_id=qid,
        question_text=f"{qid} text",
        question_type=qtype,
        raw_columns=raw_columns or (qid,),
        option_map=option_map or {},
        denominator_policy=DenominatorPolicy.VALID_RESPONSES,
        grid_row_labels=grid_labels,
        **kwargs,
    )


def _schema(*questions: QuestionSpec) -> SurveySchema:
    return SurveySchema(
        questions=tuple(questions),
        respondent_id_column="respondent_id",
        total_respondents=200,
        source_datamap_path="datamap.xlsx",
        source_rawdata_path="raw.xlsx",
        parsed_at=datetime.now(timezone.utc),
    )


class TestHypothesisValidator(unittest.TestCase):
    def test_classify_question_role_for_supported_types(self) -> None:
        ordinal = _question(
            "Q_ORD",
            QuestionType.SINGLE_SELECT,
            option_map={1: "1 low", 2: "2 high"},
            label_to_numeric_value={"1 low": 1.0, "2 high": 2.0},
            allowed_numeric_range=(1.0, 2.0),
        )
        cases = [
            (_question("Q_NUM", QuestionType.DIRECT_NUMERIC), "continuous"),
            (_question("Q_ALLOC", QuestionType.NUMERIC_ALLOCATION, ("Q_ALLOCr1", "Q_ALLOCr2")), "continuous"),
            (_question("Q_GRID", QuestionType.GRID_RATED, ("Q_GRIDr1c1", "Q_GRIDr1c2")), "continuous"),
            (_question("Q_CAT", QuestionType.SINGLE_SELECT, option_map={1: "A"}), "categorical"),
            (_question("Q_MS", QuestionType.MULTI_SELECT_BINARY, ("Q_MSr1",)), "categorical"),
            (ordinal, "ordinal"),
            (_question("Q_RANK", QuestionType.RANK_ORDER, ("Q_RANKr1",)), "ordinal"),
            (_question("Q_TEXT", QuestionType.OPEN_TEXT), "unsupported"),
        ]
        for question, expected in cases:
            self.assertEqual(_classify_question_role(question), expected)

    def test_select_test_matrix(self) -> None:
        self.assertEqual(_select_test("continuous", "continuous", 0), "pearson_r")
        self.assertEqual(_select_test("continuous", "ordinal", 0), "spearman_rho")
        self.assertEqual(_select_test("ordinal", "ordinal", 0), "spearman_rho")
        self.assertEqual(_select_test("continuous", "categorical", 2), "welch_t")
        self.assertEqual(_select_test("continuous", "categorical", 3), "anova_oneway")
        self.assertEqual(_select_test("categorical", "categorical", 2), "chi_square_cramer_v")
        with self.assertRaises(ValueError):
            _select_test("categorical", "continuous", 0)

    def test_run_pearson_on_known_data(self) -> None:
        stat = _run_pearson(np.arange(80), np.arange(80) * 2)
        self.assertAlmostEqual(stat.effect_size, 1.0, places=6)
        self.assertLess(stat.p_value, 0.05)
        self.assertEqual(stat.sample_size, 80)
        self.assertIsNotNone(stat.confidence_interval_low)

    def test_run_spearman_on_known_data(self) -> None:
        stat = _run_spearman([1, 2, 3, 4, 5], [10, 20, 30, 40, 50])
        self.assertAlmostEqual(stat.effect_size, 1.0, places=6)

    def test_run_welch_t_with_two_groups(self) -> None:
        stat = _run_welch_t(np.repeat(10.0, 60), np.repeat(5.0, 60))
        self.assertEqual(stat.test_name, "welch_t")
        self.assertGreater(stat.effect_size, 0)
        self.assertEqual(stat.sample_sizes_per_group, (60, 60))

    def test_run_anova_with_three_groups(self) -> None:
        rng = np.random.default_rng(seed=42)
        group_a = 1.0 + rng.normal(0, 0.5, 60)
        group_b = 3.0 + rng.normal(0, 0.5, 60)
        group_c = 6.0 + rng.normal(0, 0.5, 60)
        stat = _run_anova([group_a, group_b, group_c])
        self.assertEqual(stat.test_name, "anova_oneway")
        self.assertGreaterEqual(stat.effect_size, 0.06)

    def test_run_chi_square_cramer_v(self) -> None:
        table = np.array([[80, 20], [20, 80]])
        stat = _run_chi_square(table)
        expected_v = math.sqrt(float(stat.raw_test_statistic) / (table.sum() * 1))
        self.assertAlmostEqual(stat.effect_size, expected_v, places=6)
        self.assertEqual(stat.sample_size, 200)

    def test_compute_pearson_ci_reference_shape(self) -> None:
        low, high = _compute_pearson_ci(0.5, 100)
        self.assertLess(low, 0.5)
        self.assertGreater(high, 0.5)

    def test_compute_effect_helpers(self) -> None:
        self.assertAlmostEqual(_compute_cohen_d(2.0, 50, 50), 0.4)
        self.assertAlmostEqual(_compute_eta_squared(3.0, 2.0, 97.0), 6.0 / 103.0)

    def test_eta_squared_handles_infinite_f(self) -> None:
        """Zero within-group variance produces F=inf; eta-squared should be 1.0."""
        self.assertEqual(_compute_eta_squared(float("inf"), 2.0, 0.0), 1.0)
        self.assertEqual(_compute_eta_squared(0.0, 0.0, 0.0), 0.0)

    def test_validate_continuous_vs_continuous_pearson_path(self) -> None:
        df = pd.DataFrame({"QY": np.arange(120), "QX": np.arange(120)})
        schema = _schema(_question("QY", QuestionType.DIRECT_NUMERIC), _question("QX", QuestionType.DIRECT_NUMERIC))
        log = CalculationLog()
        result = validate_hypothesis(
            HypothesisSpec("Positive correlation", "QY", "QX", "correlated_positive"),
            df,
            schema,
            log,
            cross_cut_lookup={("QY", "QX"): "CC_01"},
        )
        self.assertEqual(result.statistic.test_name, "pearson_r")
        self.assertEqual(result.verdict, "CONFIRMED")
        self.assertEqual(result.related_cross_cut_id, "CC_01")
        self.assertGreaterEqual(len(result.calculation_log_entry_ids), 4)

    def test_validate_continuous_vs_categorical_three_group_anova(self) -> None:
        df = pd.DataFrame(
            {
                "QY": [1] * 60 + [3] * 60 + [6] * 60,
                "QG": [1] * 60 + [2] * 60 + [3] * 60,
            }
        )
        schema = _schema(
            _question("QY", QuestionType.DIRECT_NUMERIC),
            _question("QG", QuestionType.SINGLE_SELECT, option_map={1: "A", 2: "B", 3: "C"}),
        )
        result = validate_hypothesis(
            HypothesisSpec("Groups differ", "QY", "QG", "different"),
            df,
            schema,
            CalculationLog(),
        )
        self.assertEqual(result.statistic.test_name, "anova_oneway")
        self.assertEqual(result.verdict, "CONFIRMED")

    def test_validate_continuous_vs_categorical_two_group_welch(self) -> None:
        df = pd.DataFrame({"QY": [10.0] * 60 + [1.0] * 60, "QG": [1] * 60 + [2] * 60})
        schema = _schema(
            _question("QY", QuestionType.DIRECT_NUMERIC),
            _question("QG", QuestionType.SINGLE_SELECT, option_map={1: "A", 2: "B"}),
        )
        result = validate_hypothesis(
            HypothesisSpec("Two groups differ", "QY", "QG", "different"),
            df,
            schema,
            CalculationLog(),
        )
        self.assertEqual(result.statistic.test_name, "welch_t")
        self.assertEqual(result.verdict, "CONFIRMED")

    def test_validate_categorical_vs_categorical_chi_square(self) -> None:
        df = pd.DataFrame(
            {
                "QO": [1] * 80 + [1] * 20 + [2] * 20 + [2] * 80,
                "QP": [1] * 100 + [2] * 100,
            }
        )
        schema = _schema(
            _question("QO", QuestionType.SINGLE_SELECT, option_map={1: "Yes", 2: "No"}),
            _question("QP", QuestionType.SINGLE_SELECT, option_map={1: "A", 2: "B"}),
        )
        result = validate_hypothesis(
            HypothesisSpec("Different shares", "QO", "QP", "different"),
            df,
            schema,
            CalculationLog(),
        )
        self.assertEqual(result.statistic.test_name, "chi_square_cramer_v")
        self.assertEqual(result.verdict, "CONFIRMED")

    def test_unknown_question_returns_could_not_classify(self) -> None:
        result = validate_hypothesis(
            HypothesisSpec("Unknown", "QY", "Q_MISSING", "correlated_positive"),
            pd.DataFrame({"QY": range(60)}),
            _schema(_question("QY", QuestionType.DIRECT_NUMERIC)),
            CalculationLog(),
        )
        self.assertEqual(result.verdict, "INCONCLUSIVE")
        self.assertEqual(result.verdict_reason, "could_not_classify_questions")
        self.assertIsNone(result.statistic)

    def test_sample_size_floor_returns_inconclusive(self) -> None:
        stat = HypothesisStatistic("pearson_r", 0.9, "Pearson r", 0.001, None, None, 49)
        self.assertFalse(_check_sample_size_floor([49]))
        verdict, reason = _verdict_from_statistic(stat, "correlated_positive")
        self.assertEqual(verdict, "CONFIRMED")
        df = pd.DataFrame({"QY": range(49), "QX": range(49)})
        result = validate_hypothesis(
            HypothesisSpec("Small n", "QY", "QX", "correlated_positive"),
            df,
            _schema(_question("QY", QuestionType.DIRECT_NUMERIC), _question("QX", QuestionType.DIRECT_NUMERIC)),
            CalculationLog(),
        )
        self.assertEqual(result.verdict, "INCONCLUSIVE")
        self.assertEqual(result.verdict_reason, "insufficient_sample_size")

    def test_verdict_confirmed_refuted_tiny_and_non_significant(self) -> None:
        strong_pos = HypothesisStatistic("pearson_r", 0.5, "Pearson r", 0.001, None, None, 100)
        self.assertEqual(_verdict_from_statistic(strong_pos, "correlated_positive"), ("CONFIRMED", "effect_meets_threshold"))
        self.assertEqual(_verdict_from_statistic(strong_pos, "correlated_negative"), ("REFUTED", "direction_opposite"))
        tiny = HypothesisStatistic("pearson_r", 0.1, "Pearson r", 0.001, None, None, 100)
        self.assertEqual(_verdict_from_statistic(tiny, "correlated_positive")[0], "INCONCLUSIVE")
        weak_p = HypothesisStatistic("pearson_r", 0.5, "Pearson r", 0.2, None, None, 100)
        self.assertEqual(_verdict_from_statistic(weak_p, "correlated_positive")[1], "not_statistically_significant")

    def test_cross_cut_lookup_missing_adds_warning(self) -> None:
        df = pd.DataFrame({"QY": np.arange(60), "QX": np.arange(60)})
        result = validate_hypothesis(
            HypothesisSpec("No cc", "QY", "QX", "correlated_positive"),
            df,
            _schema(_question("QY", QuestionType.DIRECT_NUMERIC), _question("QX", QuestionType.DIRECT_NUMERIC)),
            CalculationLog(),
        )
        self.assertIsNone(result.related_cross_cut_id)
        self.assertIn("no_matching_cross_cut_found", result.warnings)

    def test_cross_cut_lookup_reverse_string_pair_populates_id(self) -> None:
        df = pd.DataFrame({"QY": np.arange(60), "QX": np.arange(60)})
        result = validate_hypothesis(
            HypothesisSpec("Reverse cc", "QY", "QX", "correlated_positive"),
            df,
            _schema(_question("QY", QuestionType.DIRECT_NUMERIC), _question("QX", QuestionType.DIRECT_NUMERIC)),
            CalculationLog(),
            cross_cut_lookup={"QX|QY": "Cross_Cut_Reverse"},
        )
        self.assertEqual(result.related_cross_cut_id, "Cross_Cut_Reverse")
        self.assertNotIn("no_matching_cross_cut_found", result.warnings)

    def test_pairwise_nan_drop(self) -> None:
        df = pd.DataFrame({"QY": [1.0, None, 3.0, 4.0], "QX": [1.0, 2.0, None, 4.0]})
        prepared = hv._prepare_test_data(
            HypothesisSpec("NaNs", "QY", "QX", "correlated_positive"),
            df,
            _schema(_question("QY", QuestionType.DIRECT_NUMERIC), _question("QX", QuestionType.DIRECT_NUMERIC)),
        )
        self.assertEqual(prepared["pairwise_n"], 2)

    def test_grid_rated_outcome_row_mean(self) -> None:
        df = pd.DataFrame({"QG1": [8, 6], "QG2": [10, 8], "QX": [1, 2]})
        qg = _question("QG", QuestionType.GRID_RATED, ("QG1", "QG2"))
        prepared = hv._prepare_test_data(
            HypothesisSpec("Grid", "QG", "QX", "correlated_positive"),
            df,
            _schema(qg, _question("QX", QuestionType.DIRECT_NUMERIC)),
        )
        self.assertEqual(prepared["outcome"].tolist(), [9.0, 7.0])

    def test_multi_select_predictor_collapses_to_binary_t_test(self) -> None:
        df = pd.DataFrame({"QY": [10] * 60 + [1] * 60, "QMSr1": [1] * 60 + [0] * 60})
        schema = _schema(
            _question("QY", QuestionType.DIRECT_NUMERIC),
            _question("QMS", QuestionType.MULTI_SELECT_BINARY, ("QMSr1",), option_map={"QMSr1": "Selected"}),
        )
        result = validate_hypothesis(
            HypothesisSpec("Selected higher", "QY", "QMS", "different"),
            df,
            schema,
            CalculationLog(),
        )
        self.assertEqual(result.statistic.test_name, "welch_t")

    def test_calculation_log_entries_emitted(self) -> None:
        df = pd.DataFrame({"QY": np.arange(80), "QX": np.arange(80)})
        log = CalculationLog()
        result = validate_hypothesis(
            HypothesisSpec("Log", "QY", "QX", "correlated_positive"),
            df,
            _schema(_question("QY", QuestionType.DIRECT_NUMERIC), _question("QX", QuestionType.DIRECT_NUMERIC)),
            log,
        )
        self.assertGreaterEqual(len(log), 4)
        self.assertGreaterEqual(len(result.calculation_log_entry_ids), 4)

    def test_ai_functions_noop_without_key(self) -> None:
        with patch.dict(os.environ, {"PORTKEY_API_KEY": ""}, clear=False):
            schema = _schema(_question("QY", QuestionType.DIRECT_NUMERIC), _question("QX", QuestionType.DIRECT_NUMERIC))
            self.assertIsNone(parse_freetext_hypothesis("x", schema))
            result = validate_hypothesis(
                HypothesisSpec("x", "QY", "QX", "correlated_positive"),
                pd.DataFrame({"QY": range(60), "QX": range(60)}),
                schema,
                CalculationLog(),
            )
            self.assertEqual(explain_verdict(result), "")

    def test_parse_freetext_hypothesis_mock_returns_spec(self) -> None:
        captured: dict[str, str] = {}

        class _FakeMessage:
            content = (
                '{"title":"Parsed","outcome_question_id":"QY",'
                '"predictor_question_id":"QX","outcome_direction":"correlated_positive"}'
            )

        class _FakeChoice:
            message = _FakeMessage()

        class _FakeResponse:
            choices = [_FakeChoice()]

        class _FakeCompletions:
            def create(self, **kwargs):
                captured["system"] = kwargs["messages"][0]["content"]
                return _FakeResponse()

        class _FakeChat:
            completions = _FakeCompletions()

        class _FakeClient:
            def __init__(self, **_kwargs):
                self.chat = _FakeChat()

        schema = _schema(_question("QY", QuestionType.DIRECT_NUMERIC), _question("QX", QuestionType.DIRECT_NUMERIC))
        with patch.dict(os.environ, {"PORTKEY_API_KEY": "test"}, clear=False):
            with patch.object(hv, "OpenAI", _FakeClient):
                spec = parse_freetext_hypothesis("Y rises with X", schema)
        self.assertIsNotNone(spec)
        self.assertEqual(spec.outcome_question_id, "QY")
        self.assertIn("Return ONLY JSON", captured["system"])

    def test_explain_verdict_prompt_forbids_number_invention(self) -> None:
        captured: dict[str, str] = {}

        class _FakeMessage:
            content = "Explanation"

        class _FakeChoice:
            message = _FakeMessage()

        class _FakeResponse:
            choices = [_FakeChoice()]

        class _FakeCompletions:
            def create(self, **kwargs):
                captured["system"] = kwargs["messages"][0]["content"]
                return _FakeResponse()

        class _FakeChat:
            completions = _FakeCompletions()

        class _FakeClient:
            def __init__(self, **_kwargs):
                self.chat = _FakeChat()

        stat = HypothesisStatistic("pearson_r", 0.5, "Pearson r", 0.01, 0.2, 0.7, 100)
        result = hv.HypothesisResult(
            spec=HypothesisSpec("x", "QY", "QX", "correlated_positive"),
            verdict="CONFIRMED",
            verdict_reason="effect_meets_threshold",
            statistic=stat,
            cohort_n=100,
            pairwise_n=100,
            active_filters_summary="(all respondents)",
            related_cross_cut_id=None,
            calculation_log_entry_ids=(),
        )
        with patch.dict(os.environ, {"PORTKEY_API_KEY": "test"}, clear=False):
            with patch.object(hv, "OpenAI", _FakeClient):
                self.assertEqual(explain_verdict(result), "Explanation")
        self.assertIn("Do NOT invent values", captured["system"])
        self.assertIn("Do NOT round or change numbers", captured["system"])


if __name__ == "__main__":
    unittest.main()
