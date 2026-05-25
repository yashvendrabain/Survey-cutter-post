"""Global DataFrame filter application for Survey Insight Engine analyses."""

from __future__ import annotations

import pandas as pd

from src.models import GlobalFilterState


def _specific_filter_values(filter_spec) -> list | None:
    values = getattr(filter_spec, "filter_values", None)
    if values is None:
        return None
    values_list = list(values)
    return values_list or None


def _filter_description(global_state: GlobalFilterState) -> str:
    if not global_state.filters:
        return "(no global filter)"
    descriptions = []
    for filter_spec in global_state.filters:
        values = _specific_filter_values(filter_spec)
        if values is not None:
            descriptions.append(f"{filter_spec.filter_question_id} in {values}")
        else:
            descriptions.append(
                f"{filter_spec.filter_question_id} == {filter_spec.filter_value!r}"
            )
    return " AND ".join(descriptions)


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
        values = _specific_filter_values(filter_spec)
        if values is not None:
            mask = mask & df[filter_spec.filter_question_id].isin(values)
        else:
            mask = mask & (df[filter_spec.filter_question_id] == filter_spec.filter_value)

    filtered_df = df[mask].copy()
    rows_after = int(len(filtered_df))

    return filtered_df, {
        "rows_before": rows_before,
        "rows_after": rows_after,
        "rows_removed": rows_before - rows_after,
        "filter_description": _filter_description(global_state),
    }
