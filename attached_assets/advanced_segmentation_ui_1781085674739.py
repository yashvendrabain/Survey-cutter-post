"""Round 2 — Advanced Outcome Segmentation dashboard (drop-in).

Wire with ONE line inside the existing Outcome Segmentation screen, behind an
"Open Advanced" toggle (see WIRING note at the bottom). Leaves the current
Winners-vs-Laggards flow untouched.

Operates on real df COLUMNS (from session_state["active_df"]) so it needs no
guesswork about how the schema maps questions -> columns. Performance metrics
drive the composite score; balance dimensions (sector/region) are
representativeness checks + an optional stratified pick.
"""

from __future__ import annotations

import streamlit as st

from src.winner_scoring import (
    BalanceStrategy, MetricDirection, WinnerMetricSpec, WinnerScoringConfig,
    compute_winner_scoring, suggest_band_midpoints,
)

_METRIC_HINTS = ("revenue", "rev", "growth", "margin", "gross", "ebitda", "profit")
_DIM_HINTS = ("sector", "industry", "region", "country", "geography", "geo", "market")


def _guess(cols, hints, n=1):
    out = [c for c in cols if any(h in str(c).lower() for h in hints)]
    return out[:n]


def _is_band_column(series) -> bool:
    """Non-numeric column with a small set of distinct labels -> ordinal band."""
    import pandas as pd
    num = pd.to_numeric(series, errors="coerce")
    if num.notna().mean() > 0.8:
        return False
    return series.nunique(dropna=True) <= 20


def render_advanced_outcome_segmentation() -> None:
    df = st.session_state.get("active_df")
    if df is None or getattr(df, "empty", True):
        st.info("Run an analysis first — the advanced dashboard needs loaded data.")
        return

    cols = list(df.columns)
    st.markdown("#### Advanced winner / laggard builder")
    st.caption("Define winners across several metrics at once. A respondent only "
               "wins if they score well on the weighted blend — not on one axis alone.")

    mode = st.radio("Mode", ["Manual", "AI (Round 3)"], horizontal=True,
                    key="adv_seg_mode")
    if mode.startswith("AI"):
        st.info("AI mode arrives in Round 3: GPT will recommend the metric set, "
                "weights, cutoff and balance strategy with written reasoning, then "
                "Python computes the actual cohort list. Requires the Portkey key "
                "rotated into Replit Secrets first.")
        return

    # ---- optional scope ----
    with st.expander("Scope (optional) — restrict to one sector / region", expanded=False):
        scope_dim = st.selectbox("Scope dimension", ["(whole population)"] + cols,
                                 key="adv_seg_scope_dim")
        work_df = df
        if scope_dim != "(whole population)":
            vals = sorted(str(v) for v in df[scope_dim].dropna().unique())
            scope_val = st.selectbox("Scope value", vals, key="adv_seg_scope_val")
            work_df = df[df[scope_dim].astype(str) == scope_val]
            st.caption(f"Scoped to {scope_dim} = {scope_val}: {len(work_df)} respondents.")

    # ---- performance metrics ----
    st.markdown("**Performance metrics** (drive the score)")
    default_metrics = _guess(cols, _METRIC_HINTS, n=2)
    chosen_metrics = st.multiselect(
        "Metrics", cols, default=default_metrics, key="adv_seg_metrics",
        help="Pick the columns that define performance (e.g. revenue growth, gross margin). "
             "Add a custom one here too.")

    specs: list[WinnerMetricSpec] = []
    for col in chosen_metrics:
        with st.expander(f"⚙ {col}", expanded=False):
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
                    default = suggested.get(lab)
                    val = st.number_input(
                        lab, value=float(default) if default is not None else 0.0,
                        key=f"adv_band_{col}_{lab}")
                    pairs.append((lab, float(val)))
                band_pairs = tuple(pairs)
            specs.append(WinnerMetricSpec(
                question_id=col, column=col, weight=float(weight),
                direction=(MetricDirection.HIGHER_IS_BETTER if direction.startswith("Higher")
                           else MetricDirection.LOWER_IS_BETTER),
                band_midpoints=band_pairs))

    # ---- balance / representativeness ----
    st.markdown("**Balance dimensions** (keep winners representative)")
    default_dims = _guess(cols, _DIM_HINTS, n=2)
    balance_dims = st.multiselect("Sector / region", cols, default=default_dims,
                                  key="adv_seg_dims")
    do_stratify = st.checkbox("Balanced pick — draw winners proportionally within each "
                              "category (avoids all-one-sector cohorts)",
                              key="adv_seg_stratify")
    stratify_dim = None
    if do_stratify and balance_dims:
        stratify_dim = st.selectbox("Stratify by", balance_dims, key="adv_seg_stratdim")

    # ---- cutoff ----
    cutoff = st.slider("Winners = top X% · Laggards = bottom X%", 5, 50, 25, 5,
                       key="adv_seg_cutoff")
    pct = cutoff / 100.0

    if not specs:
        st.warning("Pick at least one performance metric to compute cohorts.")
        return

    config = WinnerScoringConfig(
        metrics=tuple(specs), winner_pct=pct, laggard_pct=pct,
        balance_strategy=(BalanceStrategy.STRATIFIED if stratify_dim else BalanceStrategy.NONE),
        balance_dimensions=tuple(balance_dims), stratify_dimension=stratify_dim,
        min_metrics_present=1)

    try:
        result = compute_winner_scoring(work_df, config)
    except Exception as exc:  # noqa: BLE001
        st.error(f"Scoring failed: {type(exc).__name__}: {exc}")
        return

    # ---- live counts ----
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Winners", result.winner_count)
    m2.metric("Laggards", result.laggard_count)
    m3.metric("Middle", result.middle_count)
    m4.metric("Excluded (no data)", result.excluded_count)
    if result.balance_warning:
        st.warning(result.balance_warning)
    elif result.is_balanced:
        st.success(f"Balanced cohorts (ratio {result.balance_ratio:.2f}).")

    # ---- composition ----
    if result.composition:
        st.markdown("**Winner composition vs population**")
        import pandas as pd
        comp_df = pd.DataFrame([{
            "Dimension": c.dimension_id, "Category": c.category,
            "Winner %": round(c.winner_share * 100, 1),
            "Population %": round(c.population_share * 100, 1),
            "Index": (round(c.index_ratio, 2) if c.index_ratio is not None else None),
            "Over-indexed": "⚠" if c.over_indexed else "",
        } for c in result.composition])
        st.dataframe(comp_df, use_container_width=True, hide_index=True)
        try:
            import plotly.graph_objects as go
            for dim in balance_dims:
                sub = [c for c in result.composition if c.dimension_id == dim]
                if not sub:
                    continue
                cats = [c.category for c in sub]
                fig = go.Figure()
                fig.add_bar(name="Winners", x=cats, y=[c.winner_share * 100 for c in sub],
                            marker_color="#CC0000")
                fig.add_bar(name="Population", x=cats, y=[c.population_share * 100 for c in sub],
                            marker_color="#BBBBBB")
                fig.update_layout(barmode="group", height=300, title=dim,
                                  font_family="Arial", margin=dict(t=40, b=20))
                st.plotly_chart(fig, use_container_width=True)
        except Exception:
            pass
        if any(c.over_indexed for c in result.composition):
            st.caption("⚠ Over-indexed categories: winners are concentrated here. "
                       "Turn on the balanced pick to spread them across categories.")

    # ---- lock ----
    if st.button("Lock these cohorts", type="primary", key="adv_seg_lock"):
        st.session_state["advanced_winner_result"] = result
        st.session_state["advanced_winner_ids"] = result.winner_ids
        st.session_state["advanced_laggard_ids"] = result.laggard_ids
        st.success(f"Locked: {result.winner_count} winners / {result.laggard_count} "
                   "laggards. The differentiator + AI-insight pipeline can now run on "
                   "this cohort.")
        # WIRING (one line, depends on your existing engine signature):
        #   feed result.winner_ids / result.laggard_ids (df-index labels) — or
        #   winner_mask(work_df, result) / laggard_mask(work_df, result) — into
        #   your existing compute_outcome_segmentation() as the manual cohort split.


# ---------------------------------------------------------------------------
# WIRING — add to the Outcome Segmentation screen in app.py:
#
#   from src.advanced_segmentation_ui import render_advanced_outcome_segmentation
#   ...
#   if st.toggle("Open Advanced", key="outcome_seg_advanced"):
#       render_advanced_outcome_segmentation()
#   else:
#       <existing winners-vs-laggards rendering, unchanged>
# ---------------------------------------------------------------------------
