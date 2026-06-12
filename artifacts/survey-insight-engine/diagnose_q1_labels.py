"""
Diagnostic: find exactly where Q1's multi-select option labels (United States...)
are lost. Run from the project root:

    cd ~/workspace/artifacts/survey-insight-engine
    python3 diagnose_q1_labels.py "<path to the .xlsx survey input you uploaded>"

It prints Q1's structure at each pipeline stage so we can see which step drops
the labels. Paste the full output back.
"""
from __future__ import annotations

import sys


def main() -> None:
    if len(sys.argv) < 2:
        print("usage: python3 diagnose_q1_labels.py <survey_input.xlsx>")
        sys.exit(1)
    path = sys.argv[1]

    from src.datamap_parser import parse_datamap  # type: ignore
    # 1) Raw parse of the datamap
    datamap = parse_datamap(path)
    questions = getattr(datamap, "questions", datamap)

    def find(qid: str):
        for q in questions:
            cid = q["canonical_id"] if isinstance(q, dict) else getattr(q, "canonical_id", None)
            if cid == qid:
                return q
        return None

    for qid in ("Q1", "QTerms"):
        q = find(qid)
        print(f"\n================= {qid} =================")
        if q is None:
            print("  NOT FOUND in parsed datamap")
            continue
        get = (lambda k: q.get(k)) if isinstance(q, dict) else (lambda k: getattr(q, k, None))
        print("  STAGE 1 — raw parsed ParsedQuestion:")
        print("    value_range:", get("value_range"))
        print("    options    :", get("options"))
        subs = get("sub_columns") or []
        print(f"    sub_columns ({len(subs)}): first 5 = {subs[:5]}")

    # 2) After label-pattern matching (the adapter that runs before classify)
    print("\n\n========== AFTER apply_label_pattern_matching ==========")
    try:
        from src.adapters.label_pattern_subcolumn import apply_label_pattern_matching  # type: ignore
        for qid in ("Q1", "QTerms"):
            q = find(qid)
            if q is None:
                continue
            raw_order = tuple(
                (q.get("sub_columns") if isinstance(q, dict) else getattr(q, "sub_columns", []))
            )
            ids = tuple(s[0] for s in (raw_order or []))
            transformed = apply_label_pattern_matching(q, ids)
            tget = (lambda k: transformed.get(k)) if isinstance(transformed, dict) else (lambda k: getattr(transformed, k, None))
            subs = tget("sub_columns") or []
            print(f"\n  {qid}: sub_columns ({len(subs)}) first 5 = {subs[:5]}")
            print(f"       options = {tget('options')}")
    except Exception as exc:  # noqa: BLE001
        print("  (could not run adapter directly:", type(exc).__name__, exc, ")")

    # 3) Final spec after classification — the option_map the calculators see
    print("\n\n========== FINAL QuestionSpec (what _multi_select reads) ==========")
    try:
        from src.question_classifier import classify_questions  # type: ignore
        schema = classify_questions(datamap)
        specs = getattr(schema, "questions", schema)
        for qid in ("Q1", "QTerms"):
            spec = next((s for s in specs if getattr(s, "canonical_id", None) == qid), None)
            if spec is None:
                print(f"\n  {qid}: NOT FOUND")
                continue
            om = getattr(spec, "option_map", {}) or {}
            items = list(om.items())[:5]
            print(f"\n  {qid}: type={getattr(spec,'question_type',None)}")
            print(f"       option_map ({len(om)}) first 5 = {items}")
            grl = getattr(spec, "grid_row_labels", None) or {}
            print(f"       grid_row_labels ({len(grl)}) first 5 = {list(grl.items())[:5]}")
    except Exception as exc:  # noqa: BLE001
        print("  (could not classify:", type(exc).__name__, exc, ")")


if __name__ == "__main__":
    main()
