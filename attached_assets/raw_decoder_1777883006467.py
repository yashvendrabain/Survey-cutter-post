"""Raw survey data decoder for the Survey Insight Engine."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd

from src.datamap_parser import DataMap, ParsedQuestion
from src.models import DataQualityReport

try:
    from config import MISSING_VALUE_TOKENS, HIGH_MISSING_THRESHOLD
except ModuleNotFoundError as exc:
    if exc.name != "config":
        raise
    MISSING_VALUE_TOKENS = {"", "NA", "N/A", "NULL", "null", "None", "nan"}
    HIGH_MISSING_THRESHOLD = 0.5


RESPONDENT_ID_CANDIDATES = (
    "record",
    "uuid",
    "respondent_id",
    "id",
    "ID",
    "RespondentID",
)


def decode_raw_data(
    path: str,
    data_map: DataMap,
) -> tuple[pd.DataFrame, DataQualityReport]:
    """Load raw survey data and return a decoded DataFrame plus quality report."""

    dataframe = _load_raw_file(path)
    dataframe = _strip_string_values(dataframe)
    dataframe = _replace_missing_tokens(dataframe)

    original_columns = tuple(str(column) for column in dataframe.columns)
    respondent_id_column = _find_respondent_id_column(dataframe.columns.tolist())
    warnings: list[str] = []
    generated_respondent_id = False

    if respondent_id_column is None:
        respondent_id_column = "respondent_id"
        dataframe[respondent_id_column] = range(1, len(dataframe) + 1)
        generated_respondent_id = True
        warnings.append("no respondent ID column found; generated sequential IDs")

    expected_columns = _expected_columns_from_datamap(data_map)
    columns_in_datamap = len(set(original_columns).intersection(expected_columns))
    columns_not_in_datamap = tuple(
        column for column in original_columns if column not in expected_columns
    )
    value_ranges = _value_ranges_by_column(data_map)
    no_numeric_coercion_columns = _string_preserved_columns(data_map)

    coercion_log: list[dict] = []
    for column in dataframe.columns:
        if column == respondent_id_column:
            continue
        if column in no_numeric_coercion_columns:
            continue
        dataframe[column] = _coerce_column_to_numeric(
            dataframe[column], str(column), coercion_log
        )

    missing_pct = {
        str(column): float(dataframe[column].isna().mean())
        for column in dataframe.columns
    }
    out_of_range_pct = {
        str(column): _out_of_range_pct(dataframe[column], value_ranges.get(str(column)))
        for column in dataframe.columns
    }

    for column, pct in missing_pct.items():
        if pct > HIGH_MISSING_THRESHOLD:
            warnings.append(f"column {column} has {pct:.1%} missing values")
    for column, pct in out_of_range_pct.items():
        if pct > 0.05:
            warnings.append(f"column {column} has {pct:.1%} out-of-range values")

    report = DataQualityReport(
        total_rows=int(len(dataframe)),
        total_columns=int(len(dataframe.columns)),
        columns_in_datamap=int(columns_in_datamap),
        columns_not_in_datamap=columns_not_in_datamap,
        per_column_missing_pct=missing_pct,
        per_column_out_of_range_pct=out_of_range_pct,
        coercion_log=tuple(coercion_log),
        warnings=tuple(warnings),
    )

    if generated_respondent_id:
        dataframe[respondent_id_column] = dataframe[respondent_id_column].astype("Int64")

    return dataframe, report


def _load_raw_file(path: str) -> pd.DataFrame:
    suffix = Path(path).suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(path, dtype=str)
    if suffix == ".xlsx":
        return pd.read_excel(path, dtype=str, engine="openpyxl")
    raise ValueError(f"unsupported raw data file extension: {suffix or '<none>'}")


def _strip_string_values(dataframe: pd.DataFrame) -> pd.DataFrame:
    stripped = dataframe.copy()
    for column in stripped.columns:
        stripped[column] = stripped[column].map(
            lambda value: value.strip() if isinstance(value, str) else value
        )
    return stripped


def _replace_missing_tokens(dataframe: pd.DataFrame) -> pd.DataFrame:
    replaced = dataframe.copy()
    for column in replaced.columns:
        replaced[column] = replaced[column].map(_replace_missing_value)
    return replaced


def _replace_missing_value(value: Any) -> Any:
    if pd.isna(value):
        return pd.NA
    if isinstance(value, str) and value in MISSING_VALUE_TOKENS:
        return pd.NA
    return value


def _find_respondent_id_column(columns: list[str]) -> str | None:
    for candidate in RESPONDENT_ID_CANDIDATES:
        if candidate in columns:
            return candidate
    return None


def _expected_columns_from_datamap(data_map: DataMap) -> set[str]:
    expected: set[str] = set()
    for question in data_map["questions"]:
        sub_columns = question["sub_columns"]
        if sub_columns:
            expected.update(sub_column_id for sub_column_id, _ in sub_columns)
        else:
            expected.add(question["canonical_id"])
    return expected


def _value_ranges_by_column(data_map: DataMap) -> dict[str, tuple[int, int]]:
    value_ranges: dict[str, tuple[int, int]] = {}
    for question in data_map["questions"]:
        value_range = question["value_range"]
        if value_range is None:
            continue
        for column in _question_expected_columns(question):
            value_ranges[column] = value_range
    return value_ranges


def _string_preserved_columns(data_map: DataMap) -> set[str]:
    preserved: set[str] = set()
    for question in data_map["questions"]:
        if question["type_hint"] in {"open_text", "open_numeric"}:
            preserved.update(_question_expected_columns(question))
    return preserved


def _question_expected_columns(question: ParsedQuestion) -> tuple[str, ...]:
    if question["sub_columns"]:
        return tuple(sub_column_id for sub_column_id, _ in question["sub_columns"])
    return (question["canonical_id"],)


def _coerce_column_to_numeric(
    series: pd.Series, column_name: str, coercion_log: list[dict]
) -> pd.Series:
    original = series.copy()
    numeric_series = pd.to_numeric(original, errors="coerce")
    affected_mask = original.notna() & numeric_series.isna()
    affected_count = int(affected_mask.sum())

    if affected_count > 0:
        values_coerced = sorted(
            {str(value) for value in original[affected_mask].dropna().tolist()}
        )
        coercion_log.append(
            {
                "column": column_name,
                "from_type": "string",
                "to_type": "numeric",
                "values_coerced": values_coerced,
                "rows_affected": affected_count,
            }
        )

    if int(numeric_series.notna().sum()) == 0:
        return original
    return numeric_series


def _out_of_range_pct(
    series: pd.Series, value_range: tuple[int, int] | None
) -> float:
    if value_range is None or not pd.api.types.is_numeric_dtype(series):
        return 0.0

    valid_values = series.dropna()
    valid_n = int(len(valid_values))
    if valid_n == 0:
        return 0.0

    low, high = value_range
    out_of_range_count = int((~valid_values.between(low, high)).sum())
    return float(out_of_range_count / valid_n)
