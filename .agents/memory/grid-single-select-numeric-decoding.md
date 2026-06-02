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

## Structural-detection decoder (round "5K.4") — current approach + known gap
After the classifier pre-pass was rolled back, the decoder was rewritten to detect grids
**structurally** from raw column headers (no classifier dependency): `_grid_patterns_for_decoding`
→ `_detect_grid_pattern_for_decoding` tries, in order, `_has_double_colon_grid_columns`
(`Qx: dim1 :: dim2`), `_has_categorical_grid_row_columns` (`Qx: <rowlabel>` headers whose labels
do NOT match the question's option labels), `_has_explicit_grid_row_sub_columns` (≥2 sub_columns).
A detected question is gated OUT of option-mapping so numeric grid values are preserved. This
**fixes** the Q4 SINGLE_SELECT regression (plain single-col questions detect no pattern → still
option-mapped) and preserves Q27/Q14 numeric grids.

**KNOWN SEVERE GAP (confirmed, not yet fixed):** the binary-multi-select exemption is asymmetric.
`_has_explicit_grid_row_sub_columns` calls `_grid_options_look_binary_selection` to bail out when
options look like Selected/Not-selected; **`_has_categorical_grid_row_columns` does NOT**. So a
binary multi-select with colon-style headers (`Q3: B2B`, `Q3: B2C`) AND generic options
(`Selected`/`Not selected`) is mis-detected as `grid_categorical_row` → binary decode is SKIPPED →
raw codes survive instead of "Selected". Confirmed by repro: options=segment-names → pattern None
(correct); options=Selected/Not-selected → grid_categorical_row (bug). This directly threatens the
winvslag Q3 multi-select verification (expects "Selected"). **Fix belongs in Codex:** add the
`_grid_options_look_binary_selection` guard to `_has_categorical_grid_row_columns`.
**Why:** the row-label-vs-option-label test passes only when option labels equal the segment
names; when options are generic Selected/Not-selected they never match the segment row labels, so
the "looks like a grid" heuristic fires on a plain multi-select.

## Type-gated classifier pre-pass — TRIED AND ROLLED BACK
A revision (round "5K.3") that ran the classifier pre-pass inside `decode_raw_data` was
**rolled back** after real-fixture (winvslag2024) testing showed it was net-negative:
1. **New regression:** plain `SINGLE_SELECT` questions (single column + integer codes, e.g. Q4
   "2,000 to 4,999 employees") stopped being option-mapped — `_RawData` kept raw ints (5,6,4…),
   so downstream `COUNTIFS` against label strings all returned 0. Q1/Q2 (same structure) still
   worked, i.e. the pre-pass mis-gated some SINGLE_SELECT as grid/open and skipped mapping.
2. **Original bug NOT fixed:** Q27/Q97/Q14 STILL showed "Selected" — the grid collapse is
   **upstream** (raw encoding / classification), not in the option-map overwrite path.
**Lesson:** the option-map overwrite path was never the true root cause for the "Selected"
grids; gating it via a classifier pre-pass only broke working SINGLE_SELECT decoding. Do NOT
reinstate the pre-pass. Any real fix must be validated against the **real** winvslag2024
fixture (Q4 must decode to labels; Q1/Q2 must stay working), not synthetic tests.
The detail below describes the rolled-back design for reference only.

## (rolled-back design detail) Type-gated classifier pre-pass + fail-open risk
A revision runs the full `question_classifier` EARLY inside `decode_raw_data`
(`_question_types_for_decoding`) to produce `{canonical_id: QuestionType}`, then
`_decode_option_columns` gates on it: `GRID_RATED` → skip option mapping entirely;
`GRID_SINGLE_SELECT` (by classified type OR the `_is_grid_single_select_question` heuristic)
→ grid decode. Ordering is safe: the classifier keys off data-map metadata + raw column
names, not post-coercion values, and uses copied dicts so it does not mutate `data_map`.

**FAIL-OPEN HAZARD (durable):** `_question_types_for_decoding` does `except Exception: return {}`
with no logging. On classifier crash the type map is empty, and **GRID_RATED protection is
lost** — GRID_RATED is gated ONLY on classified type (no heuristic), so a rated grid WITH
non-empty options would be re-corrupted by generic option mapping. GRID_SINGLE_SELECT still
survives via its heuristic; rated grids with empty options also survive (empty option_map →
skipped). This is fail-open on the critical path: the safe behavior would be to skip option
mapping for all grid-like sub-column questions in fallback mode.
**Why:** the whole point of the pre-pass is to prevent numeric-grid corruption; swallowing the
exception silently re-enables the exact bug it was added to fix.
