"""Global DataFrame filter application for Survey Insight Engine analyses."""

from __future__ import annotations

import pandas as pd

from src.models import GlobalFilterState


def apply_global_filter(
    df: pd.DataFrame,
    global_state: GlobalFilterState,
) -> tuple[pd.DataFrame, dict[str, int | str]]:
    """Apply the global filter to df and return the filtered data plus stats."""

    rows_before = int(len(df))

    if not global_state.is_active():
        return df, {
            "rows_before": rows_before,
            "rows_after": rows_before,
            "rows_removed": 0,
            "filter_description": "(no global filter)",
        }

    mask = pd.Series(True, index=df.index)
    for filter_spec in global_state.filters:
        if filter_spec.filter_question_id not in df.columns:
            raise ValueError(
                f"global filter column {filter_spec.filter_question_id!r} not in data"
            )
        values = filter_spec.get_effective_values()
        if values is None:
            continue
        column = df[filter_spec.filter_question_id]
        if len(values) == 1:
            mask = mask & (column == values[0])
        else:
            mask = mask & column.isin(values)

    filtered_df = df[mask].copy()
    rows_after = int(len(filtered_df))

    return filtered_df, {
        "rows_before": rows_before,
        "rows_after": rows_after,
        "rows_removed": rows_before - rows_after,
        "filter_description": global_state.description(),
    }
