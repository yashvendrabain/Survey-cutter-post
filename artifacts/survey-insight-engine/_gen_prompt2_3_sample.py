from __future__ import annotations
import os, sys, time
os.environ.setdefault("PORTKEY_API_KEY", "")
sys.path.insert(0, os.path.dirname(__file__))

from src.datamap_parser import parse_datamap
from src.raw_decoder import decode_raw_data
from src.calculation_log import CalculationLog
from src.question_classifier import classify_questions
from src.single_cut import compute_single_cuts
from src.excel_exporter import export_single_cuts
from src.models import DataQualityReport

t0 = time.time()
data_map = parse_datamap("sample_data/Datamap_sample.xlsx")
df, _ = decode_raw_data("sample_data/Rawdata_sample.xlsx", data_map)
print(f"[{time.time()-t0:5.1f}s] loaded rows={len(df)} cols={len(df.columns)}")

schema = classify_questions(data_map, df.columns.tolist(), respondent_id_column="record", total_respondents=len(df), source_rawdata_path="sample_data/Rawdata_sample.xlsx")
log = CalculationLog()
results, skips = compute_single_cuts(schema, df, log)
print(f"[{time.time()-t0:5.1f}s] results={len(results)} skips={len(skips)}")

qreport = DataQualityReport(total_rows=len(df), total_columns=len(df.columns), columns_in_datamap=len(df.columns), columns_not_in_datamap=tuple(), per_column_missing_pct={c:0.0 for c in df.columns}, per_column_out_of_range_pct={c:0.0 for c in df.columns}, coercion_log=tuple(), warnings=tuple())
themes = {"themes":[{"theme_name":"All Questions","question_ids":[r.question_id for r in results]}]}
short_labels = {q.canonical_id: q.question_text[:40] for q in schema.questions}

export_single_cuts(results=results, skips=skips, schema=schema, quality_report=qreport, log=log, output_path="outputs/survey_prompt2_3_sample.xlsx", themes=themes, decoded_df=df, short_labels=short_labels)
print(f"[{time.time()-t0:5.1f}s] sample done size={os.path.getsize('outputs/survey_prompt2_3_sample.xlsx')/1024/1024:.2f}MiB")
