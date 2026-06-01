---
name: Wizard fallback must gate at the adapter parse floor, not the "use" threshold
description: survey-insight-engine adapter router — needs_wizard must use CONFIDENCE_REJECT_THRESHOLD, not CONFIDENCE_USE_THRESHOLD, or it intercepts files the adapters can actually parse.
---

# Wizard fallback gate must equal the adapter parse-acceptance floor

In `artifacts/survey-insight-engine/src/adapters/registry.py`, the router has two distinct
confidence gates:
- `CONFIDENCE_REJECT_THRESHOLD` (0.3) — `pick_adapter` only refuses to parse *below* this.
- `CONFIDENCE_USE_THRESHOLD` (0.5) — was originally a dead constant.

`needs_wizard` decides whether to divert a file to the manual setup wizard. It MUST gate at
`CONFIDENCE_REJECT_THRESHOLD` (`best_score < CONFIDENCE_REJECT_THRESHOLD`).

**Why:** when `needs_wizard` gated at `CONFIDENCE_USE_THRESHOLD` (0.5), any file scoring in the
0.3–0.5 band — which `pick_adapter` can fully parse — was wrongly diverted to the wizard. This
regressed `winvslag2024.xlsx` (worked through Round 5K.1, broke after the 5L wizard was added).
The wizard path was brand new; the bug was that its gate was stricter than the adapters' real
capability, creating a 0.3–0.5 "dead zone".

**How to apply:** if a previously-working file suddenly routes to the wizard, check that
`needs_wizard` gates at the same floor `pick_adapter` rejects at (0.3), not at USE_THRESHOLD.
Lowering/raising USE_THRESHOLD is a red herring — it's unused by the parse path.
