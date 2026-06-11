"""Advanced Outcome Segmentation — deterministic scoring core.

Self-contained on purpose so it drops into the repo with zero edits to
models.py / calc_primitives.py. Move the dataclasses to models.py and the
percentile helper to calc_primitives.py later if you prefer; the math lives
in ONE place here (_percentile_rank) and is not reimplemented elsewhere.

NO Streamlit, NO Portkey/OpenAI imports. Pure pandas/numpy.
AI never computes here — this is the deterministic engine the AI mode
(Round 3) will only *configure*.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from enum import Enum
from typing import Optional

import numpy as np
import pandas as pd


# --------------------------------------------------------------------------
# Models  (frozen + slots)
# --------------------------------------------------------------------------
class MetricDirection(Enum):
    HIGHER_IS_BETTER = "higher_is_better"
    LOWER_IS_BETTER = "lower_is_better"


class BalanceStrategy(Enum):
    NONE = "none"
    STRATIFIED = "stratified"


@dataclass(frozen=True, slots=True)
class WinnerMetricSpec:
    question_id: str
    weight: float = 1.0
    direction: MetricDirection = MetricDirection.HIGHER_IS_BETTER
    # tuple of (label, midpoint) pairs so the spec stays hashable/frozen
    band_midpoints: Optional[tuple[tuple[str, float], ...]] = None
    column: Optional[str] = None  # explicit df column override (UI / tests)


@dataclass(frozen=True, slots=True)
class WinnerScoringConfig:
    metrics: tuple[WinnerMetricSpec, ...]
    winner_pct: float = 0.25
    laggard_pct: float = 0.25
    balance_strategy: BalanceStrategy = BalanceStrategy.NONE
    balance_dimensions: tuple[str, ...] = ()
    stratify_dimension: Optional[str] = None
    min_metrics_present: int = 1
    combination_mode: str = "blend"  # "blend" (weighted composite) | "intersection" (top X% on every metric)


@dataclass(frozen=True, slots=True)
class RespondentScore:
    respondent_id: object
    metric_percentiles: tuple[tuple[str, float], ...]
    metrics_present: int
    composite_score: Optional[float]
    cohort: str  # "winner" | "laggard" | "middle" | "excluded"


@dataclass(frozen=True, slots=True)
class CompositionStat:
    dimension_id: str
    category: str
    winner_count: int
    winner_share: float
    population_count: int
    population_share: float
    index_ratio: Optional[float]
    over_indexed: bool


@dataclass(frozen=True, slots=True)
class WinnerScoringResult:
    config: WinnerScoringConfig
    winner_ids: tuple
    laggard_ids: tuple
    middle_ids: tuple
    excluded_ids: tuple
    winner_count: int
    laggard_count: int
    middle_count: int
    excluded_count: int
    balance_ratio: Optional[float]
    is_balanced: bool
    balance_warning: Optional[str]
    respondent_scores: tuple[RespondentScore, ...]
    composition: tuple[CompositionStat, ...]
    calc_log: tuple[dict, ...]


# --------------------------------------------------------------------------
# Band-label -> midpoint
# --------------------------------------------------------------------------
_OPEN_TOKENS = ("+", ">", "<", "≥", "≤", "more", "less", "above", "below",
                "under", "over", "up to", "at least", "at most", "plus",
                "greater", "fewer")


def suggest_band_midpoints(labels) -> dict:
    """Map each ordinal band label to a numeric midpoint (or None).

    "30-40%" -> 35.0 ; "10 to 20" -> 15.0 ; "35%" -> 35.0
    open-ended ("40%+", "<10%", "up to 10") or unparseable -> None
    """
    out: dict[str, Optional[float]] = {}
    for raw in labels:
        out[raw] = _midpoint_for_label(raw)
    return out


def _midpoint_for_label(raw) -> Optional[float]:
    if raw is None:
        return None
    label = str(raw)
    low = label.lower()
    cleaned = (label.replace("\u2013", "-").replace("\u2014", "-")
               .replace(" to ", "-").replace(",", ""))
    cleaned = cleaned.replace("%", "").replace("$", "")
    nums = re.findall(r"\d+(?:\.\d+)?", cleaned)
    is_open = any(tok in low for tok in _OPEN_TOKENS)
    if len(nums) >= 2:
        a, b = float(nums[0]), float(nums[1])
        return (a + b) / 2.0
    if len(nums) == 1 and not is_open:
        return float(nums[0])
    return None


# --------------------------------------------------------------------------
# Percentile (SINGLE math source — do not reimplement elsewhere)
# --------------------------------------------------------------------------
def _percentile_rank(series: pd.Series, direction: MetricDirection) -> pd.Series:
    s = pd.to_numeric(series, errors="coerce")
    ascending = direction == MetricDirection.HIGHER_IS_BETTER
    return s.rank(method="average", pct=True, ascending=ascending) * 100.0


def _to_float(value) -> float:
    if value is None:
        return math.nan
    if isinstance(value, (int, float)):
        return float(value) if not (isinstance(value, float) and math.isnan(value)) else math.nan
    text = str(value).strip()
    if text == "" or text.lower() in {"n/a", "na", "none", "null", "nan", "-", "don't know", "dont know"}:
        return math.nan
    text = text.replace("%", "").replace("$", "").replace(",", "").strip()
    try:
        return float(text)
    except ValueError:
        return math.nan


# --------------------------------------------------------------------------
# Metric resolution
# --------------------------------------------------------------------------
def resolve_metric_series(df: pd.DataFrame, spec: WinnerMetricSpec,
                          schema=None) -> pd.Series:
    col = spec.column
    if col is None and schema is not None:
        col = _resolve_column_from_schema(schema, spec.question_id, df)
    if col is None or col not in df.columns:
        raise ValueError(f"metric column not found for {spec.question_id}")

    raw = df[col]
    if spec.band_midpoints:
        mapping = {str(k): v for k, v in spec.band_midpoints}
        out = raw.map(lambda v: mapping.get(str(v), math.nan)
                      if v is not None else math.nan)
        return pd.to_numeric(out, errors="coerce")
    return raw.map(_to_float)


def _resolve_column_from_schema(schema, question_id, df) -> Optional[str]:
    """Defensive: try the obvious attributes a QuestionSpec might expose."""
    try:
        q = schema.get_question(question_id)
    except Exception:
        q = None
    candidates = [question_id]
    if q is not None:
        for attr in ("canonical_id", "question_id"):
            v = getattr(q, attr, None)
            if isinstance(v, str):
                candidates.append(v)
        raw_cols = getattr(q, "raw_columns", None)
        if raw_cols:
            candidates.extend([c for c in raw_cols if isinstance(c, str)])
    for c in candidates:
        if c in df.columns:
            return c
    return None


# --------------------------------------------------------------------------
# Core
# --------------------------------------------------------------------------
def _round_count(n: int, pct: float) -> int:
    return int(round(n * pct))


def compute_winner_scoring(df: pd.DataFrame, config: WinnerScoringConfig,
                           schema=None) -> WinnerScoringResult:
    calc_log: list[dict] = []
    index = list(df.index)

    # 1. metric percentiles
    pct_by_metric: dict[str, pd.Series] = {}
    weights: dict[str, float] = {}
    for spec in config.metrics:
        try:
            series = resolve_metric_series(df, spec, schema=schema)
        except Exception as exc:  # noqa: BLE001
            calc_log.append({"stage": "metric_skipped",
                             "question_id": spec.question_id,
                             "reason": f"{type(exc).__name__}: {exc}"})
            continue
        pct = _percentile_rank(series, spec.direction)
        pct_by_metric[spec.question_id] = pct
        weights[spec.question_id] = float(spec.weight)
        calc_log.append({"stage": "metric_percentile",
                         "question_id": spec.question_id,
                         "direction": spec.direction.value,
                         "weight": float(spec.weight),
                         "valid_n": int(series.notna().sum()),
                         "formula": "rank(pct=True, method=average)*100"
                                    + (" ascending=False" if spec.direction == MetricDirection.LOWER_IS_BETTER else "")})

    # 2. composite per respondent (renormalize weights over present metrics)
    composite: dict[object, Optional[float]] = {}
    present_count: dict[object, int] = {}
    metric_pct_pairs: dict[object, list[tuple[str, float]]] = {rid: [] for rid in index}
    for rid in index:
        num = 0.0
        wsum = 0.0
        present = 0
        for qid, pct_series in pct_by_metric.items():
            val = pct_series.loc[rid]
            if pd.notna(val):
                w = weights[qid]
                num += w * float(val)
                wsum += w
                present += 1
                metric_pct_pairs[rid].append((qid, float(val)))
        present_count[rid] = present
        if present >= config.min_metrics_present and wsum > 0:
            composite[rid] = num / wsum
        else:
            composite[rid] = None
    calc_log.append({"stage": "composite",
                     "formula": "sum(weight*percentile)/sum(weight) over present metrics",
                     "scored_n": sum(1 for v in composite.values() if v is not None)})

    # 3. cohorts
    scored = [(rid, composite[rid]) for rid in index if composite[rid] is not None]
    excluded_ids = tuple(rid for rid in index if composite[rid] is None)

    if config.combination_mode == "intersection":
        # Winner = in the top X% on EVERY metric individually (strict intersection).
        # Laggard = in the bottom X% on every metric. Composite still used only for
        # ranking/tie-breaks and the middle group.
        winner_set, laggard_set = _intersection_cohorts(pct_by_metric, config, calc_log)
        # keep only respondents who actually have a composite (present metrics)
        scored_ids = {rid for rid, _ in scored}
        winner_set &= scored_ids
        laggard_set = (laggard_set & scored_ids) - winner_set
    elif config.balance_strategy == BalanceStrategy.STRATIFIED and config.stratify_dimension:
        winner_set, laggard_set = _stratified_cohorts(df, scored, config, calc_log)
    else:
        winner_set, laggard_set = _global_cohorts(scored, config)

    winner_ids = tuple(rid for rid, _ in sorted(scored, key=lambda t: (-t[1], _sk(t[0]))) if rid in winner_set)
    laggard_ids = tuple(rid for rid, _ in sorted(scored, key=lambda t: (t[1], _sk(t[0]))) if rid in laggard_set)
    middle_ids = tuple(rid for rid, _ in sorted(scored, key=lambda t: (-t[1], _sk(t[0])))
                       if rid not in winner_set and rid not in laggard_set)

    cohort_of: dict[object, str] = {}
    for rid in winner_ids:
        cohort_of[rid] = "winner"
    for rid in laggard_ids:
        cohort_of[rid] = "laggard"
    for rid in middle_ids:
        cohort_of[rid] = "middle"
    for rid in excluded_ids:
        cohort_of[rid] = "excluded"

    # 4. balance
    wc, lc = len(winner_ids), len(laggard_ids)
    balance_ratio = (wc / lc) if lc else None
    is_balanced = bool(lc > 0 and balance_ratio is not None and 0.67 <= balance_ratio <= 1.5)
    balance_warning = _balance_warning(wc, lc, balance_ratio, is_balanced)
    calc_log.append({"stage": "cohorts", "winner_count": wc, "laggard_count": lc,
                     "middle_count": len(middle_ids), "excluded_count": len(excluded_ids),
                     "balance_ratio": balance_ratio, "is_balanced": is_balanced})

    # 5. composition
    composition = _composition(df, winner_ids, scored, config, calc_log)

    # 6. respondent scores
    scores = tuple(
        RespondentScore(
            respondent_id=rid,
            metric_percentiles=tuple(metric_pct_pairs[rid]),
            metrics_present=present_count[rid],
            composite_score=composite[rid],
            cohort=cohort_of[rid],
        )
        for rid in index
    )

    return WinnerScoringResult(
        config=config, winner_ids=winner_ids, laggard_ids=laggard_ids,
        middle_ids=middle_ids, excluded_ids=excluded_ids,
        winner_count=wc, laggard_count=lc, middle_count=len(middle_ids),
        excluded_count=len(excluded_ids), balance_ratio=balance_ratio,
        is_balanced=is_balanced, balance_warning=balance_warning,
        respondent_scores=scores, composition=composition,
        calc_log=tuple(calc_log),
    )


def _intersection_cohorts(pct_by_metric, config, calc_log):
    """Strict intersection: winner iff percentile >= (100 - X) on EVERY metric.

    pct_by_metric: {question_id: pd.Series of 0-100 percentiles}.
    A respondent must have a non-null percentile on every present metric and clear
    the threshold on all of them. Laggard is the mirror (<= X on every metric).
    """
    import pandas as pd
    win_cut = (1.0 - config.winner_pct) * 100.0   # e.g. top 25% -> >= 75th pct
    lag_cut = config.laggard_pct * 100.0          # bottom 25% -> <= 25th pct
    if not pct_by_metric:
        return set(), set()
    metric_ids = list(pct_by_metric)
    # Align all metric percentile series on a common index.
    common_index = None
    for s in pct_by_metric.values():
        common_index = s.index if common_index is None else common_index.intersection(s.index)
    winner_set, laggard_set = set(), set()
    for rid in common_index:
        vals = [pct_by_metric[q].loc[rid] for q in metric_ids]
        if any(pd.isna(v) for v in vals):
            continue
        if all(float(v) >= win_cut for v in vals):
            winner_set.add(rid)
        elif all(float(v) <= lag_cut for v in vals):
            laggard_set.add(rid)
    calc_log.append({"stage": "intersection_cohorts",
                     "winner_cut_pct": win_cut, "laggard_cut_pct": lag_cut,
                     "metrics": metric_ids,
                     "winner_count": len(winner_set), "laggard_count": len(laggard_set),
                     "formula": "winner iff percentile>=100-X on ALL metrics"})
    return winner_set, laggard_set


def _sk(rid):
    """Sort key for tie-breaking by respondent id (numeric if possible)."""
    try:
        return (0, float(rid))
    except (TypeError, ValueError):
        return (1, str(rid))


def _global_cohorts(scored, config):
    n = len(scored)
    n_win = _round_count(n, config.winner_pct)
    n_lag = _round_count(n, config.laggard_pct)
    if n_win + n_lag > n:  # shrink proportionally
        total = n_win + n_lag
        n_win = int(n_win * n / total)
        n_lag = n - n_win if (n - n_win) >= 0 else 0
    desc = [rid for rid, _ in sorted(scored, key=lambda t: (-t[1], _sk(t[0])))]
    asc = [rid for rid, _ in sorted(scored, key=lambda t: (t[1], _sk(t[0])))]
    winner_set = set(desc[:n_win])
    laggard_set = set(asc[:n_lag]) - winner_set
    return winner_set, laggard_set


def _stratified_cohorts(df, scored, config, calc_log):
    dim = config.stratify_dimension
    comp = dict(scored)
    winner_set: set = set()
    laggard_set: set = set()
    per_stratum = {}
    if dim not in df.columns:
        calc_log.append({"stage": "stratified", "dimension": dim,
                         "reason": "dimension column missing; fell back to global"})
        return _global_cohorts(scored, config)
    cats = df.loc[[r for r, _ in scored], dim]
    for cat in sorted(set(str(c) for c in cats.dropna())):
        members = [(rid, comp[rid]) for rid in comp if str(df.loc[rid, dim]) == cat]
        nm = len(members)
        n_win = _round_count(nm, config.winner_pct)
        n_lag = _round_count(nm, config.laggard_pct)
        desc = [rid for rid, _ in sorted(members, key=lambda t: (-t[1], _sk(t[0])))]
        asc = [rid for rid, _ in sorted(members, key=lambda t: (t[1], _sk(t[0])))]
        w = set(desc[:n_win])
        l = set(asc[:n_lag]) - w
        winner_set |= w
        laggard_set |= l
        per_stratum[cat] = {"n": nm, "winners": len(w), "laggards": len(l)}
    laggard_set -= winner_set
    calc_log.append({"stage": "stratified", "dimension": dim,
                     "per_stratum_counts": per_stratum})
    return winner_set, laggard_set


def _balance_warning(wc, lc, ratio, is_balanced):
    msgs = []
    if lc == 0 or wc == 0:
        msgs.append("One cohort is empty — widen the percentile or check the metrics.")
    elif not is_balanced:
        msgs.append(f"Cohorts are lopsided ({wc} winners vs {lc} laggards, "
                    f"ratio {ratio:.2f}). Aim for a ratio near 1.0.")
    if 0 < wc < 10 or 0 < lc < 10:
        msgs.append(f"Small cohort(s): {wc} winners / {lc} laggards. "
                    "Results may be unstable.")
    return " ".join(msgs) if msgs else None


def _composition(df, winner_ids, scored, config, calc_log):
    out: list[CompositionStat] = []
    total_winners = len(winner_ids)
    scored_ids = [rid for rid, _ in scored]
    total_scored = len(scored_ids)
    if total_scored == 0:
        return tuple(out)
    win_set = set(winner_ids)
    for dim in config.balance_dimensions:
        if dim not in df.columns:
            calc_log.append({"stage": "composition_skipped", "dimension": dim,
                             "reason": "column missing"})
            continue
        cats = sorted(set(str(c) for c in df.loc[scored_ids, dim].dropna()))
        for cat in cats:
            pop_count = int(sum(1 for rid in scored_ids if str(df.loc[rid, dim]) == cat))
            win_count = int(sum(1 for rid in win_set if str(df.loc[rid, dim]) == cat))
            win_share = (win_count / total_winners) if total_winners else 0.0
            pop_share = (pop_count / total_scored) if total_scored else 0.0
            index_ratio = (win_share / pop_share) if pop_share else None
            over = bool(index_ratio is not None and index_ratio >= 1.5)
            out.append(CompositionStat(dimension_id=dim, category=cat,
                                       winner_count=win_count, winner_share=win_share,
                                       population_count=pop_count, population_share=pop_share,
                                       index_ratio=index_ratio, over_indexed=over))
            calc_log.append({"stage": "composition", "dimension": dim, "category": cat,
                             "winner_count": win_count, "winner_share": win_share,
                             "population_share": pop_share, "index_ratio": index_ratio,
                             "over_indexed": over})
    return tuple(out)


def winner_mask(df: pd.DataFrame, result: WinnerScoringResult) -> pd.Series:
    """Boolean Series aligned to df.index: True for winners."""
    return pd.Series(df.index.isin(result.winner_ids), index=df.index)


def laggard_mask(df: pd.DataFrame, result: WinnerScoringResult) -> pd.Series:
    return pd.Series(df.index.isin(result.laggard_ids), index=df.index)
