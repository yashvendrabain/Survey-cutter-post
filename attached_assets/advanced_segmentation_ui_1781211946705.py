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


def _is_binary_flag(series) -> bool:
    """True if the column is a 0/1 (or yes/no) flag — i.e. an exploded
    multi-select option column, which is NOT a useful single metric."""
    import pandas as pd
    vals = set(str(v).strip().lower() for v in series.dropna().unique())
    return vals.issubset({"0", "1", "0.0", "1.0", "yes", "no", "true", "false",
                          "selected", "not selected", ""}) and len(vals) <= 3


def _metric_eligible_cols(df, cols):
    """Columns that make sense as a ranking metric: numeric or banded scalars.

    Excludes exploded multi-select OPTION columns (0/1 flags) — those were
    showing up as repeated 'Q10: CRM data / Q10: Web analytics / ...' entries
    and aren't meaningful single metrics. Dedupes while preserving order.
    """
    import pandas as pd
    out = []
    seen = set()
    for c in cols:
        if c in seen:
            continue
        series = df[c]
        if _is_binary_flag(series):
            continue  # skip multi-select option flags
        numeric_share = pd.to_numeric(series, errors="coerce").notna().mean()
        if numeric_share > 0.5 or _is_band_column(series):
            out.append(c)
            seen.add(c)
    return out


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
    direction = st.radio(
        "Direction", ["Higher is better", "Lower is better"],
        key=f"{key_prefix}_dir", horizontal=True,
    )
    # Weight only matters in blend mode; tuck it away so intersection demo stays clean.
    with st.expander("Advanced: weight (blend mode only)", expanded=False):
        weight = st.slider("Weight", 0.0, 3.0, 1.0, 0.5, key=f"{key_prefix}_w",
                           help="Ignored in 'All measures at once' (intersection) mode.")

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


def _option_choice_chart(work_df, mask_win, mask_lag, mask_oth, question_col, title, key):
    """Grouped bar: % of Top group / Comparison group / Others choosing each option.

    Top group red, Comparison group black, Others grey. Percentage is within each
    group (denominator = that group's count of people who answered this question).
    """
    try:
        import plotly.graph_objects as go
        import pandas as pd
    except Exception:
        return
    col = work_df[question_col]
    answered = col.notna()
    options = [str(v) for v in col.dropna().unique()]
    if not options or len(options) > 25:
        st.caption(f"({question_col}: too many distinct values to chart.)")
        return
    win_n = max(1, int((mask_win & answered).sum()))
    lag_n = max(1, int((mask_lag & answered).sum()))
    oth_n = max(1, int((mask_oth & answered).sum()))
    win_pct, lag_pct, oth_pct = [], [], []
    for opt in options:
        is_opt = col.astype(str) == opt
        win_pct.append(round(float((is_opt & mask_win).sum()) / win_n * 100, 1))
        lag_pct.append(round(float((is_opt & mask_lag).sum()) / lag_n * 100, 1))
        oth_pct.append(round(float((is_opt & mask_oth).sum()) / oth_n * 100, 1))
    fig = go.Figure()
    fig.add_bar(name="Top group", x=options, y=win_pct, marker_color="#CC0000",
                text=[f"{v:.0f}%" for v in win_pct], textposition="outside")
    fig.add_bar(name="Comparison group", x=options, y=lag_pct, marker_color="#1A1A1A",
                text=[f"{v:.0f}%" for v in lag_pct], textposition="outside")
    fig.add_bar(name="Others", x=options, y=oth_pct, marker_color="#BBBBBB",
                text=[f"{v:.0f}%" for v in oth_pct], textposition="outside")
    # Title on top, legend BELOW the title (not overlapping). Extra top margin
    # gives both room; yaxis headroom so the outside % labels don't clip.
    fig.update_layout(
        barmode="group", height=380, font_family="Arial",
        title=dict(text=title, x=0, y=0.97, yanchor="top",
                   font=dict(size=14, color="#1A1A1A")),
        margin=dict(t=90, b=40, l=10, r=10),
        yaxis_title="% within group",
        yaxis=dict(range=[0, max(105, max(win_pct + lag_pct + oth_pct + [0]) * 1.18)]),
        legend=dict(orientation="h", yanchor="bottom", y=1.0, x=0,
                    bgcolor="rgba(0,0,0,0)"),
    )
    st.plotly_chart(fig, use_container_width=True, key=key)


def render_advanced_outcome_segmentation() -> None:
    df = st.session_state.get("active_df")
    if df is None or getattr(df, "empty", True):
        st.info("Run an analysis first — the advanced view needs loaded data.")
        return
    import pandas as pd
    cols = list(df.columns)
    work_df = df  # may be narrowed by the optional pool filter below

    st.markdown("#### Advanced segmentation builder")
    st.caption(
        "Define a Top group using up to three performance measures at once, then "
        "see how that group's answers differ from everyone else. This is a general "
        "tool — the 'Top group' can be your best-growing firms, your most satisfied "
        "customers, whatever your measures define."
    )

    # ---- Four vertical config boxes: Revenue | Gross Margin | Sector | Custom ----
    st.markdown("**Define your Top group** \u2014 set up to three measures, plus the sector/region split")
    # Clean, deduped, metric-eligible columns for the measure dropdowns (excludes
    # exploded multi-select option flags that were showing as repeated entries).
    metric_cols = _metric_eligible_cols(df, cols)
    rev_default = (_guess([c for c in metric_cols], _REVENUE_HINTS, n=1) or ["(not used)"])[0]
    mar_default = (_guess([c for c in metric_cols], _MARGIN_HINTS, n=1) or ["(not used)"])[0]
    default_dims = _guess(cols, _DIM_HINTS, n=2)

    box_rev, box_mar, box_sec, box_cus = st.columns(4)
    specs = []
    with box_rev:
        with st.container(border=True):
            s = _metric_slot("Revenue", metric_cols, work_df, rev_default, "adv_rev")
            if s: specs.append(s)
    with box_mar:
        with st.container(border=True):
            s = _metric_slot("Gross Margin", metric_cols, work_df, mar_default, "adv_mar")
            if s: specs.append(s)
    with box_sec:
        with st.container(border=True):
            st.markdown("**Sector / Region**")
            st.caption("Filter the pool and split the charts.")
            balance_dims = st.multiselect(
                "Dimensions", cols, default=default_dims, key="adv_seg_dims",
                help="Available as a filter and as a chart breakdown.",
            )
    with box_cus:
        with st.container(border=True):
            s = _metric_slot("Custom", metric_cols, work_df, "(not used)", "adv_cus")
            if s: specs.append(s)

    # Optional pool filter (kept compact, below the boxes)
    if balance_dims:
        with st.expander("Filter the pool (optional)", expanded=False):
            for dim in balance_dims:
                vals = ["(all)"] + sorted(str(v) for v in df[dim].dropna().unique())
                pick = st.selectbox(f"{dim} =", vals, key=f"adv_filter_{dim}")
                if pick != "(all)":
                    work_df = work_df[work_df[dim].astype(str) == pick]
            if len(work_df) != len(df):
                st.caption(f"Filtered pool: {len(work_df):,} of {len(df):,} respondents.")

    if not specs:
        st.warning("Pick at least one measure (Revenue, Gross Margin, or Custom).")
        return

    st.divider()

    # ---- Combination mode (intersection default) ----
    combo_label = st.selectbox(
        "How should the measures combine?",
        ["Intersection — top X% on EVERY measure (recommended)",
         "Weighted blend — top X% of a weighted average"],
        index=0, key="adv_seg_combo",
    )
    combination_mode = "intersection" if combo_label.startswith("Intersection") else "blend"

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

    mask_win = winner_mask(work_df, result)
    mask_lag = laggard_mask(work_df, result)
    # Others = everyone in the pool who is NOT a winner and NOT a laggard.
    mask_oth = ~mask_win & ~mask_lag

    # ---- Top group / Comparison group ID lists (counts + downloads) ----
    st.markdown("#### Top group & comparison group respondents")
    ccol1, ccol2, ccol3 = st.columns(3)
    ccol1.metric("Top group", int(mask_win.sum()))
    ccol2.metric("Comparison group", int(mask_lag.sum()))
    ccol3.metric("Others", int(mask_oth.sum()))
    win_ids = [str(i) for i in work_df.index[mask_win]]
    lag_ids = [str(i) for i in work_df.index[mask_lag]]
    dl1, dl2 = st.columns(2)
    dl1.download_button("Download top group IDs", "\n".join(win_ids),
                        file_name="top_group_ids.csv", key="adv_dl_win",
                        use_container_width=True, disabled=not win_ids)
    dl2.download_button("Download comparison group IDs", "\n".join(lag_ids),
                        file_name="comparison_group_ids.csv", key="adv_dl_lag",
                        use_container_width=True, disabled=not lag_ids)
    with st.expander("Show group ID lists", expanded=False):
        import pandas as pd
        maxlen = max(len(win_ids), len(lag_ids))
        st.dataframe(pd.DataFrame({
            "Top group IDs": win_ids + [""] * (maxlen - len(win_ids)),
            "Comparison group IDs": lag_ids + [""] * (maxlen - len(lag_ids)),
        }), use_container_width=True, hide_index=True)

    # ---- Per-measure charts: Top group / Comparison group / Others ----
    st.markdown("#### What each group chose, per measure")
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
            _option_choice_chart(work_df, mask_win, mask_lag, mask_oth, spec.column,
                                 f"{spec.column} \u2014 Top group vs Comparison group vs Others",
                                 key=f"adv_chart_{spec.column}")
        else:
            cats = sorted(str(v) for v in work_df[breakdown_dim].dropna().unique())
            for cat in cats:
                cat_mask = work_df[breakdown_dim].astype(str) == cat
                _option_choice_chart(
                    work_df, mask_win & cat_mask, mask_lag & cat_mask, mask_oth & cat_mask,
                    spec.column, f"{spec.column} · {breakdown_dim}={cat}",
                    key=f"adv_chart_{spec.column}_{cat}")

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
