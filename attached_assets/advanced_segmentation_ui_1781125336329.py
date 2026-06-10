"""Advanced Outcome Segmentation dashboard (Round 3).

Manual mode: analyst picks metrics/weights/cutoff/balance.
AI mode: GPT recommends the config + written reasoning; Python computes the
actual cohorts. Both modes show live counts, a winners-vs-laggards comparison
table + chart, and a sector/region composition table + chart.

Reads session_state["active_df"]. AI never computes a number — it only proposes
the configuration; every figure shown is computed by winner_scoring in Python.
"""

from __future__ import annotations

import streamlit as st

from src.winner_scoring import (
    BalanceStrategy, MetricDirection, WinnerMetricSpec, WinnerScoringConfig,
    compute_winner_scoring, compute_cohort_comparison, suggest_band_midpoints,
)

_METRIC_HINTS = ("revenue", "rev", "growth", "margin", "gross", "ebitda", "profit")
_DIM_HINTS = ("sector", "industry", "region", "country", "geography", "geo", "market")


def _guess(cols, hints, n=2):
    return [c for c in cols if any(h in str(c).lower() for h in hints)][:n]


def _is_band_column(series) -> bool:
    import pandas as pd
    num = pd.to_numeric(series, errors="coerce")
    if num.notna().mean() > 0.8:
        return False
    return 0 < series.nunique(dropna=True) <= 20


def _build_candidates(work_df, cols):
    """Auto-split columns into metric candidates (numeric / ordinal band) and
    dimension candidates (categorical) for the AI recommender."""
    import pandas as pd
    metrics, dims = [], []
    for col in cols:
        s = work_df[col]
        num = pd.to_numeric(s, errors="coerce")
        if num.notna().mean() > 0.8 and num.notna().any():
            metrics.append({"column": col, "kind": "numeric",
                            "detail": f"numeric (mean~{round(float(num.dropna().mean()), 2)})"})
            continue
        nun = s.nunique(dropna=True)
        if nun == 0 or nun > 20:
            continue
        labels = [str(v) for v in s.dropna().unique()]
        mids = suggest_band_midpoints(labels)
        if sum(1 for v in mids.values() if v is not None) >= max(2, len(labels) // 2):
            metrics.append({"column": col, "kind": "band", "detail": f"ordinal bands: {labels[:6]}"})
        else:
            dims.append({"column": col, "categories": labels[:12]})
    return metrics, dims


def _scope(df, cols):
    with st.expander("Scope (optional) — restrict to one sector / region", expanded=False):
        scope_dim = st.selectbox("Scope dimension", ["(whole population)"] + cols,
                                 key="adv_seg_scope_dim")
        if scope_dim != "(whole population)":
            vals = sorted(str(v) for v in df[scope_dim].dropna().unique())
            scope_val = st.selectbox("Scope value", vals, key="adv_seg_scope_val")
            work = df[df[scope_dim].astype(str) == scope_val]
            st.caption(f"Scoped to {scope_dim} = {scope_val}: {len(work)} respondents.")
            return work
    return df


def render_advanced_outcome_segmentation() -> None:
    df = st.session_state.get("active_df")
    if df is None or getattr(df, "empty", True):
        st.info("Run an analysis first — the advanced dashboard needs loaded data.")
        return

    cols = list(df.columns)
    st.markdown("#### Advanced winner / laggard builder")
    st.caption("Define winners across several metrics at once. A respondent only "
               "wins if they score well on the weighted blend — not on one axis alone.")

    mode = st.radio("Mode", ["Manual", "AI recommendation"], horizontal=True, key="adv_seg_mode")
    work_df = _scope(df, cols)

    if mode.startswith("AI"):
        _render_ai_mode(work_df, cols)
    else:
        _render_manual_mode(work_df, cols)


def _render_manual_mode(work_df, cols) -> None:
    st.markdown("**Performance metrics** (drive the score)")
    chosen = st.multiselect("Metrics", cols, default=_guess(cols, _METRIC_HINTS),
                            key="adv_seg_metrics",
                            help="Columns that define performance (e.g. revenue growth, gross margin).")
    specs = []
    for col in chosen:
        with st.expander(f"settings — {col}", expanded=False):
            c1, c2 = st.columns(2)
            with c1:
                direction = st.radio("Direction", ["Higher is better", "Lower is better"],
                                     key=f"adv_dir_{col}")
            with c2:
                weight = st.slider("Weight", 0.0, 3.0, 1.0, 0.5, key=f"adv_w_{col}")
            band_pairs = None
            if _is_band_column(work_df[col]):
                labels = [str(v) for v in work_df[col].dropna().unique()]
                suggested = suggest_band_midpoints(labels)
                st.caption("Ordinal bands detected — set a numeric mid-value per band:")
                pairs = []
                for lab in sorted(labels):
                    d = suggested.get(lab)
                    val = st.number_input(lab, value=float(d) if d is not None else 0.0,
                                          key=f"adv_band_{col}_{lab}")
                    pairs.append((lab, float(val)))
                band_pairs = tuple(pairs)
            specs.append(WinnerMetricSpec(
                question_id=col, column=col, weight=float(weight),
                direction=(MetricDirection.HIGHER_IS_BETTER if direction.startswith("Higher")
                           else MetricDirection.LOWER_IS_BETTER),
                band_midpoints=band_pairs))

    st.markdown("**Balance dimensions** (keep winners representative)")
    balance_dims = st.multiselect("Sector / region", cols, default=_guess(cols, _DIM_HINTS),
                                  key="adv_seg_dims")
    stratify_dim = None
    if st.checkbox("Balanced pick — draw winners proportionally within each category",
                   key="adv_seg_stratify") and balance_dims:
        stratify_dim = st.selectbox("Stratify by", balance_dims, key="adv_seg_stratdim")
    cutoff = st.slider("Winners = top X% / Laggards = bottom X%", 5, 50, 25, 5, key="adv_seg_cutoff")

    if not specs:
        st.warning("Pick at least one performance metric to compute cohorts.")
        return
    config = WinnerScoringConfig(
        metrics=tuple(specs), winner_pct=cutoff / 100.0, laggard_pct=cutoff / 100.0,
        balance_strategy=(BalanceStrategy.STRATIFIED if stratify_dim else BalanceStrategy.NONE),
        balance_dimensions=tuple(balance_dims), stratify_dimension=stratify_dim)
    _render_results(work_df, config, balance_dims)


def _render_ai_mode(work_df, cols) -> None:
    metric_cands, dim_cands = _build_candidates(work_df, cols)
    if not metric_cands:
        st.warning("No numeric or ordinal-band columns found to score winners on. "
                   "This survey may be purely categorical — load one with a "
                   "revenue / margin / growth-type question.")
        return
    st.caption(f"{len(metric_cands)} candidate metric(s), {len(dim_cands)} balance dimension(s) detected.")

    run = st.button("Get AI recommendation", type="primary", key="adv_ai_run")
    if run:
        with st.spinner("Asking the model to recommend a balanced configuration..."):
            try:
                from src.ai_insights import recommend_winner_config
                st.session_state["adv_ai_rec"] = recommend_winner_config(metric_cands, dim_cands)
            except Exception as exc:  # noqa: BLE001
                st.error(f"AI recommendation failed: {type(exc).__name__}: {exc}")
                return

    rec = st.session_state.get("adv_ai_rec")
    if not rec:
        st.info("Click **Get AI recommendation** — GPT proposes the metrics, weights, cutoff "
                "and balance strategy with reasoning; Python then computes the actual cohorts.")
        return

    if rec.get("was_template"):
        st.warning(rec.get("reasoning", "AI unavailable — used a heuristic configuration."))
    else:
        st.markdown("**AI reasoning**")
        st.info(rec.get("reasoning", ""))

    import pandas as pd
    cfg_table = pd.DataFrame([{
        "Metric": m["column"],
        "Direction": "Higher = better" if m["direction"] == "higher_is_better" else "Lower = better",
        "Weight": m["weight"],
    } for m in rec["metrics"]])
    st.markdown("**Recommended configuration**")
    st.dataframe(cfg_table, use_container_width=True, hide_index=True)
    st.caption(f"Cutoff: top/bottom {rec['cutoff_pct']}%  -  "
               f"Balance: {', '.join(rec['balance_dimensions']) or 'none'}"
               + (f"  -  stratified by {rec['stratify_dimension']}" if rec.get('stratify_dimension') else ""))
    st.caption("Switch to **Manual** above to fine-tune any of these picks.")

    specs = []
    for m in rec["metrics"]:
        col = m["column"]
        band_pairs = None
        if _is_band_column(work_df[col]):
            labels = [str(v) for v in work_df[col].dropna().unique()]
            band_pairs = tuple((lab, float(v)) for lab, v in suggest_band_midpoints(labels).items()
                               if v is not None)
        specs.append(WinnerMetricSpec(
            question_id=col, column=col, weight=float(m["weight"]),
            direction=(MetricDirection.HIGHER_IS_BETTER if m["direction"] == "higher_is_better"
                       else MetricDirection.LOWER_IS_BETTER),
            band_midpoints=band_pairs))
    strat = rec.get("stratify_dimension")
    config = WinnerScoringConfig(
        metrics=tuple(specs), winner_pct=rec["cutoff_pct"] / 100.0,
        laggard_pct=rec["cutoff_pct"] / 100.0,
        balance_strategy=(BalanceStrategy.STRATIFIED if strat else BalanceStrategy.NONE),
        balance_dimensions=tuple(rec["balance_dimensions"]), stratify_dimension=strat)
    _render_results(work_df, config, rec["balance_dimensions"])


def _render_results(work_df, config, balance_dims) -> None:
    try:
        result = compute_winner_scoring(work_df, config)
    except Exception as exc:  # noqa: BLE001
        st.error(f"Scoring failed: {type(exc).__name__}: {exc}")
        return
    import pandas as pd

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Winners", result.winner_count)
    m2.metric("Laggards", result.laggard_count)
    m3.metric("Middle", result.middle_count)
    m4.metric("Excluded (no data)", result.excluded_count)
    if result.balance_warning:
        st.warning(result.balance_warning)
    elif result.is_balanced:
        st.success(f"Balanced cohorts (ratio {result.balance_ratio:.2f}).")

    # ---- winners vs laggards comparison ----
    comp = compute_cohort_comparison(work_df, config, result)
    if comp:
        st.markdown("**Winners vs Laggards — by metric**")
        comp_df = pd.DataFrame([{
            "Metric": r["metric"],
            "Winners (avg)": round(r["winner_mean"], 2),
            "Laggards (avg)": round(r["laggard_mean"], 2),
            "Gap": round(r["gap"], 2),
            "Lift": (f"{r['lift']:.2f}x" if r["lift"] is not None else "-"),
        } for r in comp])
        st.dataframe(comp_df, use_container_width=True, hide_index=True)
        try:
            import plotly.graph_objects as go
            mets = [r["metric"] for r in comp]
            fig = go.Figure()
            fig.add_bar(name="Winners", x=mets, y=[r["winner_mean"] for r in comp], marker_color="#CC0000")
            fig.add_bar(name="Laggards", x=mets, y=[r["laggard_mean"] for r in comp], marker_color="#999999")
            fig.update_layout(barmode="group", height=320, title="Average metric value by cohort",
                              font_family="Arial", margin=dict(t=44, b=20))
            st.plotly_chart(fig, use_container_width=True)
        except Exception:
            pass

    # ---- composition ----
    if result.composition:
        st.markdown("**Winner composition vs population**")
        cdf = pd.DataFrame([{
            "Dimension": c.dimension_id, "Category": c.category,
            "Winner %": round(c.winner_share * 100, 1),
            "Population %": round(c.population_share * 100, 1),
            "Index": (round(c.index_ratio, 2) if c.index_ratio is not None else None),
            "Over-indexed": "warn" if c.over_indexed else "",
        } for c in result.composition])
        st.dataframe(cdf, use_container_width=True, hide_index=True)
        try:
            import plotly.graph_objects as go
            for dim in balance_dims:
                sub = [c for c in result.composition if c.dimension_id == dim]
                if not sub:
                    continue
                cats = [c.category for c in sub]
                fig = go.Figure()
                fig.add_bar(name="Winners", x=cats, y=[c.winner_share * 100 for c in sub], marker_color="#CC0000")
                fig.add_bar(name="Population", x=cats, y=[c.population_share * 100 for c in sub], marker_color="#BBBBBB")
                fig.update_layout(barmode="group", height=300, title=dim,
                                  font_family="Arial", margin=dict(t=40, b=20))
                st.plotly_chart(fig, use_container_width=True)
        except Exception:
            pass
        if any(c.over_indexed for c in result.composition):
            st.caption("Over-indexed categories: winners are concentrated there. "
                       "Turn on the balanced pick to spread them across categories.")

    if st.button("Lock these cohorts", type="primary", key="adv_seg_lock"):
        st.session_state["advanced_winner_result"] = result
        st.session_state["advanced_winner_ids"] = result.winner_ids
        st.session_state["advanced_laggard_ids"] = result.laggard_ids
        st.success(f"Locked: {result.winner_count} winners / {result.laggard_count} laggards.")
