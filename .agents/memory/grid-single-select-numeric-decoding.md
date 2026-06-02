---
name: GRID_SINGLE_SELECT numeric-code decoding (survey-insight-engine)
description: Why grid single-select scale columns must skip option-label mapping, and the numeric-coercion-before-text-detection ordering trap.
---

# GRID_SINGLE_SELECT numeric preservation

A "grid single select" question in `survey-insight-engine` is identified by
`_is_grid_single_select_question`: `type_hint == "values_range"` AND has `sub_columns`
AND has `options` (e.g. winvslag2024 Q27, a 1-5 scale across several row-items).

## The rule
These questions must **not** go through the generic `option_map` decoding in
`_decode_option_columns` — that path overwrote the respondent's numeric scale code (1-5)
with the option *label* (or "Selected"), destroying the actual answer. Instead they route
to `_decode_grid_single_select_question`, which **leaves numeric columns untouched** and
only applies 1/0 selection semantics to columns that actually contain text.

The numeric-vs-text decision is `_grid_column_uses_text_selection(series)`: True iff the
dropna'd series has at least one string that is non-blank, non-template, and not
float-parseable.

## Ordering trap (durable gotcha)
`_coerce_column_to_numeric` runs **before** `_decode_option_columns` for any column not in
`numeric_metadata_by_column`. GRID_SINGLE_SELECT columns are NOT in that map (they have no
`label_to_numeric_value`), so by the time grid decoding runs, a mixed column like
`["3", "Selected"]` has already become `[3.0, NaN]`. Consequence:
- pure-numeric grid columns → correctly preserved as numbers, and
- **mixed numeric+text grid columns silently drop the text selections to NaN** — no warning.
This is acceptable for clean single-encoding data but is a real correctness risk if a
production file mixes encodings row-by-row.

## Diagnostic
If a grid scale column (e.g. Q27) STILL shows "Selected" in `_RawData` after this fix, the
cause is **upstream** (raw encoding / question classification), NOT the option-map overwrite
path — that path no longer touches GRID_SINGLE_SELECT. Look earlier in the pipeline.
