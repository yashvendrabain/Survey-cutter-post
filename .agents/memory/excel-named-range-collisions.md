---
name: Excel named-range generation collisions (survey-insight-engine exporter)
description: How #NAME? errors arise from header-derived named ranges, the rstrip fix, and the remaining silent-overwrite collision gap
---

# Excel named-range generation in src/excel_exporter.py

Defined names for live-filter formulas are derived from column headers
(`_column_data_name(header)` → `f"{header.rstrip('_')}_data"`, plus `_resolved` /
`_wrapped` families). Headers are truncated to a 31-char stem for the named range.

## The #NAME? bug that was fixed
Truncating a header to 31 chars could reintroduce a **trailing underscore**, and
the old `data_name=f"{header}_data"` then produced a **double** underscore
(`..._margin__data`) that didn't match the single-underscore name defined
elsewhere (`..._margin_data`). Formulas referenced the double-underscore name →
undefined → `#NAME?` in Excel. On the real 107q winvslag2024 workbook this broke
29 of ~2706 names. Fix: `rstrip('_')` before appending `_data` at all derivation
sites, plus a `_validate_no_undefined_names(wb, strict=)` guard wired into both
export paths before `workbook.save()`.

## Remaining latent gap (NOT yet fixed — report, don't improvise)
`rstrip('_')` + 31-char truncation can collapse **two distinct headers** to the
same `*_data` name (e.g. `ABC_` and `ABC`). `_add_named_range` deletes-and-replaces
on duplicate key, so the second definition **silently overwrites** the first →
wrong data binding with **no `#NAME?`**. `_unique_live_header` enforces uniqueness
on headers, not on the post-normalized data_name, so the collision is invisible.
`_validate_no_undefined_names` only catches *undefined* references, never *wrong*
bindings, so it will NOT catch this.

**Why it matters:** the validator gives false confidence — "0 undefined names" does
not prove formulas point at the right ranges. **How to apply:** any future change to
named-range derivation should add a pre-save uniqueness check on the *normalized*
defined names (fail fast on collision), and consider exporting demos with
`strict_formula_name_validation=True` rather than the non-strict default (which only
logs + writes a Warnings sheet and can still ship a subtly-wrong workbook).
