from __future__ import annotations
import os, sys, time, pickle
os.environ.setdefault("PORTKEY_API_KEY", "")
sys.path.insert(0, os.path.dirname(__file__))

from src.calculation_log import CalculationLog
from src.question_classifier import classify_questions
from src.single_cut import compute_single_cuts
from src.excel_exporter import export_single_cuts
from src.models import DataQualityReport

t0 = time.time()
with open("outputs/_bcn_data_map.pkl","rb") as f: dmp = pickle.load(f)
with open("outputs/_bcn_raw_df.pkl","rb") as f: df = pickle.load(f)
dm = dmp["data_map"]; lr = dmp["load_report"]
print(f"[{time.time()-t0:5.1f}s] loaded pkls rows={len(df)} cols={len(df.columns)}")

schema = classify_questions(dm, df.columns.tolist(), respondent_id_column="record", total_respondents=len(df), source_rawdata_path=lr.raw_data_source)
print(f"[{time.time()-t0:5.1f}s] classified n_q={len(schema.questions)}")
for tid in ("Q26","Q30","Q36","Q38","Q41"):
    matches = [q for q in schema.questions if q.canonical_id == tid or q.canonical_id.startswith(tid+"r")]
    parent = [q for q in matches if q.canonical_id == tid]
    sibs = [q for q in matches if q.canonical_id != tid]
    print(f"  {tid}: parent={len(parent)==1} role={parent[0].possible_role if parent else None} rows={len(parent[0].grid_row_labels) if parent else 0} sibs_left={len(sibs)}")

log = CalculationLog()
results, skips = compute_single_cuts(schema, df, log)
print(f"[{time.time()-t0:5.1f}s] cuts results={len(results)} skips={len(skips)}")

qreport = DataQualityReport(total_rows=len(df), total_columns=len(df.columns), columns_in_datamap=len(df.columns), columns_not_in_datamap=tuple(), per_column_missing_pct={c:0.0 for c in df.columns}, per_column_out_of_range_pct={c:0.0 for c in df.columns}, coercion_log=tuple(), warnings=tuple())
themes = {"themes":[{"theme_name":"All Questions","question_ids":[r.question_id for r in results]}]}
short_labels = {q.canonical_id: q.question_text[:40] for q in schema.questions}

export_single_cuts(results=results, skips=skips, schema=schema, quality_report=qreport, log=log, output_path="outputs/survey_prompt2_4_bcn.xlsx", themes=themes, decoded_df=df, short_labels=short_labels)
print(f"[{time.time()-t0:5.1f}s] BCN done size={os.path.getsize('outputs/survey_prompt2_4_bcn.xlsx')/1024/1024:.2f}MiB")
