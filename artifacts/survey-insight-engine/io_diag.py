"""
Diagnostic v4 — use the SAME io layer + reconcile + compute the app uses, so we
see the actual option_map and selections the app produces (not parse_datamap).

Run from project root:
    python3 io_diag.py <raw_data_file> <datamap_file>
or if your inputs are a single workbook the io layer accepts, pass it once.
Pass the SAME file(s) you upload in the app.
"""
from __future__ import annotations
import sys


def main() -> None:
    args = sys.argv[1:]
    if not args:
        print("usage: python3 io_diag.py <input file(s) exactly as uploaded to the app>")
        sys.exit(1)

    # Mirror the app: build file-like objects from the paths.
    class _F:
        def __init__(self, p):
            self.name = p
            self._p = p
        def read(self):
            with open(self._p, "rb") as fh:
                return fh.read()
        def getvalue(self):
            return self.read()
        def seek(self, *a, **k):
            return 0

    uploaded = [_F(p) for p in args]

    from src.io import load_survey_inputs  # type: ignore
    from src.question_classifier import classify_questions, reconcile_multiselect_value_subtypes  # type: ignore
    from src.calculation_log import CalculationLog  # type: ignore
    from src.single_cut import compute_single_cuts  # type: ignore

    print("[io] calling load_survey_inputs exactly like the app...")
    data_map, raw_df, load_report = load_survey_inputs(uploaded)
    print(f"[io] raw_df shape: {getattr(raw_df, 'shape', '?')}")

    # 1) data_map sub_columns for Q1 as the io layer parsed them
    questions = data_map["questions"] if isinstance(data_map, dict) else getattr(data_map, "questions", [])
    for qid in ("Q1", "QTerms"):
        q = next((x for x in questions if (x.get("canonical_id") if isinstance(x, dict) else getattr(x, "canonical_id", None)) == qid), None)
        print(f"\n===== {qid}: io-layer parsed =====")
        if q is None:
            print("  NOT FOUND")
            continue
        g = (lambda k: q.get(k)) if isinstance(q, dict) else (lambda k: getattr(q, k, None))
        subs = g("sub_columns") or []
        print(f"  value_range = {g('value_range')}")
        print(f"  options     = {g('options')}")
        print(f"  sub_columns ({len(subs)}) first 5 = {subs[:5]}")

    # 2) classify + reconcile EXACTLY like the app
    schema = classify_questions(
        data_map,
        raw_df.columns.tolist(),
        respondent_id_column="record",
        total_respondents=len(raw_df),
        source_rawdata_path=getattr(load_report, "raw_data_source", "?"),
    )
    schema = reconcile_multiselect_value_subtypes(schema, raw_df)
    for qid in ("Q1", "QTerms"):
        spec = next((s for s in schema.questions if getattr(s, "canonical_id", None) == qid), None)
        print(f"\n===== {qid}: FINAL spec (post-reconcile, what app computes) =====")
        if spec is None:
            print("  NOT FOUND")
            continue
        om = getattr(spec, "option_map", {}) or {}
        print(f"  type = {spec.question_type}")
        print(f"  option_map ({len(om)}) first 5 = {list(om.items())[:5]}")

    # 3) compute and dump selections (the object the exporter renders)
    log = CalculationLog()
    results, skips = compute_single_cuts(schema, raw_df, log)
    for qid in ("Q1", "QTerms"):
        res = next((r for r in results if getattr(r, "question_id", None) == qid), None)
        print(f"\n===== {qid}: LIVE selections (exporter input) =====")
        if res is None:
            print("  no result")
            continue
        sel = getattr(res, "selections", None) or {}
        for k in list(sel.keys())[:5]:
            print(f"  {k!r} -> label={sel[k].get('label')!r}  count={sel[k].get('count')}")


if __name__ == "__main__":
    main()
