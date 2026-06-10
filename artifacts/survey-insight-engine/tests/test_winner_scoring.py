import math
import sys
import unittest

import pandas as pd

sys.path.insert(0, "/home/claude/seng")
from src.winner_scoring import (  # noqa: E402
    BalanceStrategy, MetricDirection, WinnerMetricSpec, WinnerScoringConfig,
    compute_winner_scoring, suggest_band_midpoints,
)


def m(col, direction=MetricDirection.HIGHER_IS_BETTER, w=1.0):
    return WinnerMetricSpec(question_id=col, weight=w, direction=direction, column=col)


def pct(result, rid, qid):
    sc = next(s for s in result.respondent_scores if s.respondent_id == rid)
    return dict(sc.metric_percentiles)[qid]


def comp(result, rid):
    return next(s for s in result.respondent_scores if s.respondent_id == rid).composite_score


class TestWinnerScoring(unittest.TestCase):

    def test_T1_band_parser(self):
        got = suggest_band_midpoints(
            ["10-20%", "20-30%", "30-40%", "40%+", "35%", "Don't know"])
        self.assertEqual(got, {"10-20%": 15.0, "20-30%": 25.0, "30-40%": 35.0,
                               "40%+": None, "35%": 35.0, "Don't know": None})

    def test_T2_core(self):
        df = pd.DataFrame({"rev": [100, 80, 60, 10], "margin": [10, 80, 60, 40]},
                          index=[1, 2, 3, 4])
        cfg = WinnerScoringConfig(metrics=(m("rev"), m("margin")),
                                  winner_pct=0.25, laggard_pct=0.25)
        r = compute_winner_scoring(df, cfg)
        self.assertEqual([pct(r, i, "rev") for i in (1, 2, 3, 4)], [100, 75, 50, 25])
        self.assertEqual([pct(r, i, "margin") for i in (1, 2, 3, 4)], [25, 100, 75, 50])
        self.assertEqual([comp(r, i) for i in (1, 2, 3, 4)], [62.5, 87.5, 62.5, 37.5])
        self.assertEqual(r.winner_ids, (2,))
        self.assertEqual(r.laggard_ids, (4,))
        self.assertEqual(r.middle_ids, (1, 3))  # high-rev/low-margin R1 NOT a winner

    def test_T3_missing(self):
        df = pd.DataFrame({"rev": [100, 80, 60, 10, math.nan, 90],
                           "margin": [10, 80, 60, 40, math.nan, math.nan]},
                          index=[1, 2, 3, 4, 5, 6])
        cfg = WinnerScoringConfig(metrics=(m("rev"), m("margin")),
                                  winner_pct=0.25, laggard_pct=0.25)
        r = compute_winner_scoring(df, cfg)
        self.assertEqual(comp(r, 2), 80.0)
        self.assertEqual(comp(r, 6), 80.0)  # rev-only, renormalized
        self.assertIsNone(comp(r, 5))
        self.assertIn(5, r.excluded_ids)
        self.assertEqual(r.winner_ids, (2,))   # tie 80 broken by lower id
        self.assertEqual(r.laggard_ids, (4,))

    def test_T4_direction(self):
        df = pd.DataFrame({"cost": [10, 20, 30, 40]}, index=[1, 2, 3, 4])
        cfg = WinnerScoringConfig(metrics=(m("cost", MetricDirection.LOWER_IS_BETTER),),
                                  winner_pct=0.25, laggard_pct=0.25)
        r = compute_winner_scoring(df, cfg)
        self.assertEqual(pct(r, 1, "cost"), 100.0)  # lowest cost -> best
        self.assertEqual(pct(r, 4, "cost"), 25.0)
        self.assertEqual(r.winner_ids, (1,))

    def test_T5_stratified(self):
        df = pd.DataFrame({
            "rev": [90, 10, 80, 20, 70, 30, 60, 40],
            "sector": ["A", "A", "A", "A", "B", "B", "B", "B"],
        }, index=range(1, 9))
        cfg = WinnerScoringConfig(
            metrics=(m("rev"),), winner_pct=0.25, laggard_pct=0.25,
            balance_strategy=BalanceStrategy.STRATIFIED, stratify_dimension="sector",
            balance_dimensions=("sector",))
        r = compute_winner_scoring(df, cfg)
        # 4 per sector, 25% -> 1 winner + 1 laggard each -> 2 winners, 2 laggards
        self.assertEqual(r.winner_count, 2)
        self.assertEqual(r.laggard_count, 2)
        win_sectors = {df.loc[i, "sector"] for i in r.winner_ids}
        self.assertEqual(win_sectors, {"A", "B"})  # one winner from each sector

    def test_T6_composition(self):
        df = pd.DataFrame({
            "rev": [100, 90, 80, 70, 10, 20, 30, 40],
            "region": ["US", "US", "US", "US", "EU", "EU", "EU", "EU"],
        }, index=range(1, 9))
        cfg = WinnerScoringConfig(metrics=(m("rev"),), winner_pct=0.5, laggard_pct=0.5,
                                  balance_dimensions=("region",))
        r = compute_winner_scoring(df, cfg)
        us = next(c for c in r.composition if c.category == "US")
        # top 50% are all US -> winner_share 1.0 vs pop_share 0.5 -> ratio 2.0
        self.assertEqual(us.winner_share, 1.0)
        self.assertEqual(us.population_share, 0.5)
        self.assertEqual(us.index_ratio, 2.0)
        self.assertTrue(us.over_indexed)


if __name__ == "__main__":
    unittest.main(verbosity=2)
