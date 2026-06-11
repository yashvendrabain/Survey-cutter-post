"""Advanced Outcome Segmentation dashboard (v2 — generic Top group vs Rest).

Drop-in replacement for the prior advanced view. Changes vs v1:
  - Generic framing throughout: "Top group" vs "Rest" (no winners/laggards wording
    in the UI). The underlying engine is unchanged.
  - Three NAMED metric slots — Revenue, Gross Margin, Custom — each with its own
    question picker, direction, and an editable top-X% threshold. Open-ended bands
    like "40%+" that have no parseable midpoint default to 45% (editable).
  - Combination mode: "All three at once (intersection)" [DEFAULT] or
    "Weighted blend". Intersection = top X% on EVERY chosen metric.
  - Sector / Region do double duty: FILTER the pool, and a BREAKDOWN that splits
    the result charts by category.
  - Per-metric charts: what options the Top group chose vs the Rest, Top group in
    red — as grouped bars.

Calculation-first: all cohort math is in src.winner_scoring; this file only builds
config, calls compute_winner_scoring, and renders what it returns.
"""

from __future__ import annotations

import streamlit as st

from src.winner_scoring import (
    BalanceStrategy, MetricDirection, WinnerMetricSpec, WinnerScoringConfig,
    compute_winner_scoring, suggest_band_midpoints, winner_mask, laggard_mask,
)

_REVENUE_HINTS = ("revenue", "rev", "growth", "sales", "topline", "top line")
_MARGIN_HINTS = ("margin", "gross", "ebitda", "profit", "gm")
_DIM_HINTS = ("sector", "industry", "region", "country", "geography", "geo", "market")

_OPEN_BAND_DEFAULT = 45.0  # "40%+" and similar open bands seed to 45% (editable)
_TOP = "Top group"
_REST = "Rest"


def _guess(cols, hints, n=1):
    out = [c for c in cols if any(h in str(c).lower() for h in hints)]
    return out[:n]


def _is_band_column(series) -> bool:
    import pandas as pd
    num = pd.to_numeric(series, errors="coerce")
    if num.notna().mean() > 0.8:
        return False
    return series.nunique(dropna=True) <= 20


def _metric_slot(label_name: str, cols, work_df, default_col, key_prefix: str):
    """Render one named metric slot. Returns a WinnerMetricSpec or None if 'None'."""
    import pandas as pd
    st.markdown(f"**{label_name}**")
    options = ["(not used)"] + cols
    idx = options.index(default_col) if default_col in options else 0
    chosen = st.selectbox(
        f"{label_name} — survey question",
        options, index=idx, key=f"{key_prefix}_col",
        help=f"Pick the question that represents {label_name.lower()}.",
    )
    if chosen == "(not used)":
        return None
    c1, c2 = st.columns(2)
    with c1:
        direction = st.radio(
            "Direction", ["Higher is better", "Lower is better"],
            key=f"{key_prefix}_dir", horizontal=True,
        )
    with c2:
        weight = st.slider("Weight (blend mode)", 0.0, 3.0, 1.0, 0.5, key=f"{key_prefix}_w")

    band_pairs = None
    if _is_band_column(work_df[chosen]):
        labels = [str(v) for v in work_df[chosen].dropna().unique()]
        suggested = suggest_band_midpoints(labels)
        st.caption(
            "This question uses ranges (bands). Set a numeric value per band. "
            "Open-ended bands like '40%+' default to 45% — edit if needed."
        )
        pairs = []
        for lab in sorted(labels):
            mid = suggested.get(lab)
            seed = float(mid) if mid is not None else _OPEN_BAND_DEFAULT
            val = st.number_input(
                lab, value=seed, key=f"{key_prefix}_band_{lab}",
                help="Default 45% shown for open-ended bands; override anytime.",
            )
            pairs.append((lab, float(val)))
        band_pairs = tuple(pairs)

    return WinnerMetricSpec(
        question_id=chosen, column=chosen, weight=float(weight),
        direction=(MetricDirection.HIGHER_IS_BETTER if direction.startswith("Higher")
                   else MetricDirection.LOWER_IS_BETTER),
        band_midpoints=band_pairs,
    )


def _option_choice_chart(work_df, mask_top, mask_rest, question_col, title):
    """Grouped bar: % of Top group vs Rest choosing each option of question_col.

    Top group bars red, Rest grey. Percentages within each cohort.
    """
    try:
        import plotly.graph_objects as go
        import pandas as pd
    except Exception:
        return
    col = work_df[question_col]
    options = [str(v) for v in col.dropna().unique()]
    if not options or len(options) > 25:
        return
    top_n = max(1, int(mask_top.sum()))
    rest_n = max(1, int(mask_rest.sum()))
    top_pct, rest_pct = [], []
    for opt in options:
        is_opt = col.astype(str) == opt
        top_pct.append(round(float((is_opt & mask_top).sum()) / top_n * 100, 1))
        rest_pct.append(round(float((is_opt & mask_rest).sum()) / rest_n * 100, 1))
    fig = go.Figure()
    fig.add_bar(name=_TOP, x=options, y=top_pct, marker_color="#CC0000",
                text=[f"{v:.0f}%" for v in top_pct], textposition="outside")
    fig.add_bar(name=_REST, x=options, y=rest_pct, marker_color="#BBBBBB",
                text=[f"{v:.0f}%" for v in rest_pct], textposition="outside")
    fig.update_layout(barmode="group", height=320, title=title,
                      font_family="Arial", margin=dict(t=46, b=30),
                      yaxis_title="% within group",
                      legend=dict(orientation="h", y=1.12))
    st.plotly_chart(fig, use_container_width=True)


def render_advanced_outcome_segmentation() -> None:
    df = st.session_state.get("active_df")
    if df is None or getattr(df, "empty", True):
        st.info("Run an analysis first — the advanced view needs loaded data.")
        return
    import pandas as pd
    cols = list(df.columns)

    st.markdown("#### Advanced segmentation builder")
    st.caption(
        "Define a Top group using up to three performance measures at once, then "
        "see how that group's answers differ from everyone else. This is a general "
        "tool — the 'Top group' can be your best-growing firms, your most satisfied "
        "customers, whatever your measures define."
    )

    mode = st.radio("Mode", ["Manual", "AI (Round 3)"], horizontal=True, key="adv_seg_mode")
    if mode.startswith("AI"):
        st.info("AI mode (Round 3): GPT recommends the measure set, weights, cutoff "
                "and balance, with written reasoning; Python computes the actual "
                "group. Requires the Portkey key in Replit Secrets.")
        return

    # ---- Sector / Region: FILTER first ----
    st.markdown("**Sector / Region**")
    st.caption("Use these to focus on part of the sample, and to split the charts below.")
    default_dims = _guess(cols, _DIM_HINTS, n=2)
    balance_dims = st.multiselect(
        "Sector / region dimensions", cols, default=default_dims, key="adv_seg_dims",
        help="Chosen dimensions are available as a filter (below) and as a chart breakdown.",
    )
    work_df = df
    with st.expander("Filter the pool (optional)", expanded=False):
        for dim in balance_dims:
            vals = ["(all)"] + sorted(str(v) for v in df[dim].dropna().unique())
            pick = st.selectbox(f"{dim} =", vals, key=f"adv_filter_{dim}")
            if pick != "(all)":
                work_df = work_df[work_df[dim].astype(str) == pick]
        if len(work_df) != len(df):
            st.caption(f"Filtered pool: {len(work_df):,} of {len(df):,} respondents.")

    st.divider()

    # ---- Three named metric slots ----
    st.markdown("**Performance measures** — define what 'top' means")
    rev_default = (_guess([c for c in cols], _REVENUE_HINTS, n=1) or ["(not used)"])[0]
    mar_default = (_guess([c for c in cols], _MARGIN_HINTS, n=1) or ["(not used)"])[0]
    specs = []
    with st.container():
        s = _metric_slot("Revenue", cols, work_df, rev_default, "adv_rev")
        if s: specs.append(s)
    with st.container():
        s = _metric_slot("Gross Margin", cols, work_df, mar_default, "adv_mar")
        if s: specs.append(s)
    with st.container():
        s = _metric_slot("Custom", cols, work_df, "(not used)", "adv_cus")
        if s: specs.append(s)

    if not specs:
        st.warning("Pick at least one performance measure (Revenue, Gross Margin, or Custom).")
        return

    st.divider()

    # ---- Combination mode (intersection default) ----
    combo_label = st.radio(
        "How should the measures combine?",
        ["All measures at once (strict)", "Weighted blend"],
        index=0, key="adv_seg_combo",
        help="Strict = a respondent must be in the top X% on EVERY measure. "
             "Blend = top X% of a weighted average across measures.",
    )
    combination_mode = "intersection" if combo_label.startswith("All") else "blend"

    cutoff = st.slider("Top group = top X%   ·   Rest comparison = bottom X%",
                       5, 50, 25, 5, key="adv_seg_cutoff")
    pct = cutoff / 100.0

    # ---- Optional balanced pick (only meaningful in blend mode) ----
    stratify_dim = None
    if combination_mode == "blend" and balance_dims:
        if st.checkbox("Balanced pick — draw the Top group proportionally within each "
                       "category (avoids an all-one-sector group)", key="adv_seg_stratify"):
            stratify_dim = st.selectbox("Balance by", balance_dims, key="adv_seg_stratdim")

    config = WinnerScoringConfig(
        metrics=tuple(specs), winner_pct=pct, laggard_pct=pct,
        balance_strategy=(BalanceStrategy.STRATIFIED if stratify_dim else BalanceStrategy.NONE),
        balance_dimensions=tuple(balance_dims), stratify_dimension=stratify_dim,
        min_metrics_present=len(specs) if combination_mode == "intersection" else 1,
        combination_mode=combination_mode,
    )

    try:
        result = compute_winner_scoring(work_df, config)
    except Exception as exc:  # noqa: BLE001
        st.error(f"Scoring failed: {type(exc).__name__}: {exc}")
        return

    # ---- Counts ----
    st.markdown("### Result")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric(_TOP, result.winner_count)
    m2.metric(_REST, result.laggard_count)
    m3.metric("Middle", result.middle_count)
    m4.metric("Excluded (no data)", result.excluded_count)
    if combination_mode == "intersection":
        st.caption(f"Strict mode: respondents in the top {cutoff}% on **all "
                   f"{len(specs)} measures** simultaneously form the Top group.")
    if result.balance_warning:
        st.warning(result.balance_warning)

    mask_top = winner_mask(work_df, result)
    mask_rest = laggard_mask(work_df, result)

    # ---- Per-measure option charts (Top group red vs Rest) ----
    st.markdown("#### What the Top group chose vs the Rest")
    breakdown_dim = None
    if balance_dims:
        breakdown_dim = st.selectbox(
            "Break charts down by (optional)", ["(no breakdown)"] + balance_dims,
            key="adv_seg_breakdown",
        )
        if breakdown_dim == "(no breakdown)":
            breakdown_dim = None

    for spec in specs:
        st.markdown(f"**{spec.column}**")
        if breakdown_dim is None:
            _option_choice_chart(work_df, mask_top, mask_rest, spec.column,
                                 f"{spec.column} — {_TOP} vs {_REST}")
        else:
            cats = sorted(str(v) for v in work_df[breakdown_dim].dropna().unique())
            for cat in cats:
                cat_mask = work_df[breakdown_dim].astype(str) == cat
                _option_choice_chart(
                    work_df, mask_top & cat_mask, mask_rest & cat_mask, spec.column,
                    f"{spec.column} · {breakdown_dim}={cat}",
                )

    # ---- Composition vs population (kept, relabelled) ----
    if result.composition:
        st.markdown("#### Top group composition vs whole pool")
        comp_df = pd.DataFrame([{
            "Dimension": c.dimension_id, "Category": c.category,
            f"{_TOP} %": round(c.winner_share * 100, 1),
            "Pool %": round(c.population_share * 100, 1),
            "Index": (round(c.index_ratio, 2) if c.index_ratio is not None else None),
            "Over-indexed": "\u26a0" if c.over_indexed else "",
        } for c in result.composition])
        st.dataframe(comp_df, use_container_width=True, hide_index=True)

    # ---- Lock ----
    if st.button("Lock this Top group", type="primary", key="adv_seg_lock"):
        st.session_state["advanced_winner_result"] = result
        st.session_state["advanced_winner_ids"] = result.winner_ids
        st.session_state["advanced_laggard_ids"] = result.laggard_ids
        st.success(f"Locked: {result.winner_count} in the Top group / "
                   f"{result.laggard_count} in the Rest. The differentiator and "
                   "AI-insight pipeline can now run on this split.")
