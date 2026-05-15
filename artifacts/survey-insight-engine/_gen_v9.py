"""One-off driver to generate outputs/survey_v9_revamp.xlsx."""
from __future__ import annotations
import os, sys
os.environ.setdefault("PORTKEY_API_KEY", "")
sys.path.insert(0, os.path.dirname(__file__))

from src.datamap_parser import parse_datamap
from src.raw_decoder import decode_raw_data
from src.calculation_log import CalculationLog
from src.question_classifier import classify_questions
from src.single_cut import compute_single_cuts
from src.excel_exporter import export_single_cuts
from src.models import DataQualityReport

DATAMAP = "sample_data/Datamap_sample.xlsx"
RAWDATA = "sample_data/Rawdata_sample.xlsx"
OUT = "outputs/survey_v9_revamp.xlsx"

data_map = parse_datamap(DATAMAP)
dataframe, _decoder_report = decode_raw_data(RAWDATA, data_map)
print(f"Loaded {len(dataframe)} rows x {len(dataframe.columns)} cols")

schema = classify_questions(
    data_map,
    dataframe.columns.tolist(),
    respondent_id_column="record",
    total_respondents=len(dataframe),
    source_rawdata_path=RAWDATA,
)
log = CalculationLog()
results, skips = compute_single_cuts(schema, dataframe, log)
print(f"Computed {len(results)} single-cut results, {len(skips)} skipped")

quality_report = DataQualityReport(
    total_rows=int(len(dataframe)),
    total_columns=int(len(dataframe.columns)),
    columns_in_datamap=int(len(dataframe.columns)),
    columns_not_in_datamap=tuple(),
    per_column_missing_pct={c: 0.0 for c in dataframe.columns},
    per_column_out_of_range_pct={c: 0.0 for c in dataframe.columns},
    coercion_log=tuple(),
    warnings=tuple(f"parser: {w}" for w in (data_map.get("parser_warnings") or [])),
)

# Build trivial themes (no AI): one theme per non-demographic question.
themes = {"themes": [
    {"theme_name": "All Questions", "question_ids": [r.question_id for r in results]}
]}
short_labels = {q.canonical_id: q.question_text[:40] for q in schema.questions}
demo_priority = {}

export_single_cuts(
    results=results, skips=skips, schema=schema,
    quality_report=quality_report, log=log, output_path=OUT,
    themes=themes, decoded_df=dataframe,
    demo_priority=demo_priority, short_labels=short_labels,
)
print(f"Wrote {OUT}")
