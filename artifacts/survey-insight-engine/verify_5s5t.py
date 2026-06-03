import sys
import pandas as pd
from src.question_classifier import (
    classify_questions,
    reconcile_multiselect_value_subtypes,
    GRID_BINARY_SELECT,
)
from src.excel_exporter import (
    _grid_spec_subtype, _decode_option_value, _is_selected_value, _live_value_present,
)

DATA = "../../attached_assets/winvslag2024.xlsx"

# Production loader. If load_survey_inputs needs different args, adjust here and paste the error.
# load_survey_inputs expects a list of UploadedFile-like objects (name/read/seek),
# not a path string, so wrap the workbook bytes in a minimal adapter.
import io as _stdlib_io
import os as _os
from src.io import load_survey_inputs


class _PathUpload:
    def __init__(self, path):
        self.name = _os.path.basename(path)
        with open(path, "rb") as fh:
            self._buf = _stdlib_io.BytesIO(fh.read())

    def read(self):
        self._buf.seek(0)
        return self._buf.read()

    def seek(self, offset):
        return self._buf.seek(offset)


data_map, df, _report = load_survey_inputs([_PathUpload(DATA)])
df.columns = [str(c) for c in df.columns]

schema = classify_questions(data_map, list(df.columns), total_respondents=len(df))
pre = {q.canonical_id: q.question_type.value for q in schema.questions}
schema = reconcile_multiselect_value_subtypes(schema, df)
by = {q.canonical_id: q for q in schema.questions}

# (1) 5S: exactly the 10 expected reclassifications
changed = [(q, by[q].question_type.value) for q in by if pre[q] != by[q].question_type.value]
rank = {q for q, t in changed if t == "RANK_ORDER"}
alloc = {q for q, t in changed if t == "NUMERIC_ALLOCATION"}
assert rank == {"Q21","Q39","Q82","Q86","Q92","Q100","Q101"}, f"RANK mismatch: {sorted(rank)}"
assert alloc == {"Q66","Q71","Q94"}, f"ALLOC mismatch: {sorted(alloc)}"
print("PASS 5S: rank", sorted(rank), "alloc", sorted(alloc))

# (2) Issue 1: the 9 reported grids keep a real GRID_* role (no warning-text poisoning)
broken = ["Q27","Q28","Q30","Q62","Q63","Q64","Q91","Q97","Q103"]
bad = [q for q in broken
       if str(by[q].possible_role) not in ("GRID_RATED","GRID_CATEGORICAL","GRID_BINARY_SELECT")]
assert not bad, f"poisoned possible_role: {bad}"

def render(spec, v):
    if _grid_spec_subtype(spec) == GRID_BINARY_SELECT:
        return "Selected" if _is_selected_value(v) else None
    if getattr(spec, "label_to_numeric_value", None):
        return None if not _live_value_present(v) else v
    return _decode_option_value(v, spec.option_map)

still_selected = []
for q in broken:
    for c in by[q].raw_columns:
        if c in df.columns and "Selected" in set(df[c].map(lambda v: render(by[q], v)).dropna().unique()):
            still_selected.append((q, c)); break
assert not still_selected, f"still rendering 'Selected': {still_selected}"
print("PASS issue1: 9 grids decode to labels, none render 'Selected'")

# (3) Q97 must decode to its real scale, not a binary flag
q97 = by["Q97"]
dist = df[q97.raw_columns[0]].map(lambda v: render(q97, v)).value_counts(dropna=True).to_dict()
assert "Selected" not in dist and any(isinstance(k, str) for k in dist), f"Q97 not decoded: {dist}"
print("Q97 first-column distribution:", dist)
print("ALL REAL-FILE CHECKS PASSED")
