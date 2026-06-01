---
name: Adapter data-map sheet lookup & wizard-fallback gating (survey-insight-engine)
description: Why files route to the format wizard — the real cause was a hardcoded "Sheet1" lookup, not the confidence threshold. Also covers the needs_wizard gate.
---

# Why a parseable workbook routes to the format wizard

In `artifacts/survey-insight-engine`, two separate things govern wizard fallback. Keep them
straight — one was a red herring.

## 1. The REAL root cause: hardcoded data-map sheet name
Both adapters (`compact_two_column`, `bcn_multicolumn`) originally located the data-map sheet
by the hardcoded constant `DATAMAP_SHEET_NAME = "Sheet1"`. A workbook whose data map lived on a
differently-named sheet (e.g. `"Data map"`, `"Datamap"`, `"Codebook"`) returned confidence 0.0
from *every* adapter → `needs_wizard` True → wizard. This is what actually broke `winvslag2024`.

**Fix:** a `_find_datamap_sheet(workbook)` helper in each adapter that matches sheet names by
normalized substring against a `DATAMAP_KEYWORDS` list (datamap, data map, codebook, schema,
dictionary, questions, questionnaire, variables, metadata), falling back to `"Sheet1"`.

**How to apply:** if a real workbook routes to the wizard, FIRST check whether its data map is on
a sheet named something other than `"Sheet1"` and confirm `_find_datamap_sheet` recognizes it.
`compact.detect()` is binary (0.9 / 0.0); `bcn.detect()` is 0.7 (format) + raw-column boosts.
Neither produces a 0.3–0.5 score for a compact file.

## 2. Red herring: the confidence threshold
An earlier hotfix changed `needs_wizard` to gate at `CONFIDENCE_REJECT_THRESHOLD` (0.3) instead
of `CONFIDENCE_USE_THRESHOLD` (0.5). **This could NOT have fixed winvslag** — because the adapter
scores are binary/banded, no compact file ever lands in 0.3–0.5. The "0.3–0.5 dead zone" theory
was wrong. The gate-at-0.3 change is still correct in principle (it should match the
`pick_adapter` parse-acceptance floor, registry.py:43), but it was not the winvslag fix.

**Lesson:** before theorizing about thresholds, confirm the adapter actually produces a score in
the contested band. Here it never did — the score was a hard 0.0 from a sheet-name miss.
