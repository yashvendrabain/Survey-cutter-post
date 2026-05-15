"""End-to-end BCN driver: mimics the Streamlit Run-Analysis path on the
combined BCN xlsx, with peak-memory + wall-time tracking. Reports Q38 counts
and Run_Summary notice from the produced workbook."""
from __future__ import annotations
import os, sys, time, resource, gc
os.environ.setdefault("PORTKEY_API_KEY", "")
sys.path.insert(0, os.path.dirname(__file__))


class _UF:
    """Mimic streamlit's UploadedFile interface (just .name + read())."""
    def __init__(self, path: str):
        self._path = path
        self.name = os.path.basename(path)
        self._buf: bytes | None = None

    def read(self) -> bytes:
        if self._buf is None:
            with open(self._path, "rb") as f:
                self._buf = f.read()
        return self._buf

    def seek(self, *_a, **_k):
        return 0

    def getvalue(self) -> bytes:
        return self.read()


def _peak_mb() -> float:
    # ru_maxrss is KB on Linux
    return resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 1024.0


def main(bcn_path: str, out_path: str) -> None:
    print(f"BCN driver — input: {bcn_path}")
    print(f"Pre-load peak RSS: {_peak_mb():.1f} MiB")

    t0 = time.time()
    from src.io import load_survey_inputs
    data_map, raw_df, load_report = load_survey_inputs([_UF(bcn_path)])
    t_load = time.time() - t0
    print(f"[{t_load:6.1f}s] load_survey_inputs OK  rows={len(raw_df)} cols={len(raw_df.columns)}  peak={_peak_mb():.1f} MiB")
    print(f"          scenario={load_report.scenario}  questions={load_report.questions_parsed}")

    from src.question_classifier import classify_questions
    from src.calculation_log import CalculationLog
    from src.single_cut import compute_single_cuts
    from src.excel_exporter import export_single_cuts
    from src.models import DataQualityReport

    schema = classify_questions(
        data_map,
        raw_df.columns.tolist(),
        respondent_id_column="record",
        total_respondents=len(raw_df),
        source_rawdata_path=bcn_path,
    )
    log = CalculationLog()
    t1 = time.time()
    results, skips = compute_single_cuts(schema, raw_df, log)
    t_cuts = time.time() - t1
    print(f"[{time.time()-t0:6.1f}s] compute_single_cuts OK  results={len(results)} skips={len(skips)}  peak={_peak_mb():.1f} MiB  ({t_cuts:.1f}s)")

    qr = DataQualityReport(
        total_rows=int(len(raw_df)),
        total_columns=int(len(raw_df.columns)),
        columns_in_datamap=int(len(raw_df.columns)),
        columns_not_in_datamap=tuple(),
        per_column_missing_pct={c: 0.0 for c in raw_df.columns},
        per_column_out_of_range_pct={c: 0.0 for c in raw_df.columns},
        coercion_log=tuple(),
        warnings=tuple(),
    )
    themes = {"themes": [
        {"theme_name": "All Questions", "question_ids": [r.question_id for r in results]}
    ]}
    short_labels = {q.canonical_id: q.question_text[:40] for q in schema.questions}

    t2 = time.time()
    export_single_cuts(
        results=results, skips=skips, schema=schema,
        quality_report=qr, log=log, output_path=out_path,
        themes=themes, decoded_df=raw_df,
        demo_priority={}, short_labels=short_labels,
    )
    t_exp = time.time() - t2
    print(f"[{time.time()-t0:6.1f}s] export_single_cuts OK  ({t_exp:.1f}s)  peak={_peak_mb():.1f} MiB")
    print(f"Wrote {out_path} ({os.path.getsize(out_path)/1024/1024:.2f} MiB)")
    print(f"TOTAL wall: {time.time()-t0:.1f}s  TOTAL peak RSS: {_peak_mb():.1f} MiB")


if __name__ == "__main__":
    bcn = sys.argv[1] if len(sys.argv) > 1 else "sample_data/BCN_combined.xlsx"
    out = sys.argv[2] if len(sys.argv) > 2 else "outputs/survey_v9_bcn.xlsx"
    main(bcn, out)
