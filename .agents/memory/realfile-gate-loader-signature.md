---
name: Real-file gate loader signature
description: survey-insight-engine verify scripts must wrap the workbook path, not pass it as a string, to load_survey_inputs.
---

# Real-file gate loader signature

Codex deploy-runbook verify scripts (e.g. `verify_5s5t.py`) call `load_survey_inputs(DATA)` with `DATA` as a path string. That always fails with `AttributeError: 'str' object has no attribute 'name'`.

**Rule:** `load_survey_inputs` takes a *list* of Streamlit-`UploadedFile`-like objects, each exposing `name: str`, `read() -> bytes`, `seek(offset)`. Wrap the path in a tiny adapter and pass a list: `load_survey_inputs([_PathUpload(DATA)])`.

**Why:** the runbooks keep handing a bare path; the production loader detects scenario from `file.name`, so a string blows up immediately. The runbook itself authorizes "adjust the loader call here and paste the error" — so fixing the *verify harness* (not any installed source file) is in-bounds.

**How to apply:** when a 5x deploy runbook includes a real-file gate, expect to add the `_PathUpload` wrapper (name from `os.path.basename`, BytesIO buffer, read() seeks to 0). Place verify_*.py in the artifact dir so `../../attached_assets/winvslag2024.xlsx` resolves.
