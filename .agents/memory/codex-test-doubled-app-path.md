---
name: Codex-uploaded test doubled APP_PATH bug
description: Recurring bug in Codex-authored test files for survey-insight-engine — APP_PATH resolves to a non-existent doubled path.
---

# Codex test files: doubled APP_PATH

Codex-authored `tests/test_*.py` files for `artifacts/survey-insight-engine` sometimes
set `APP_PATH = Path(__file__).resolve().parents[1] / "artifacts" / "survey-insight-engine" / "app.py"`.

Because the test file already lives **inside** `artifacts/survey-insight-engine/tests/`,
`parents[1]` is already `artifacts/survey-insight-engine/`, so the extra
`/ "artifacts" / "survey-insight-engine"` produces a non-existent doubled path
(`.../survey-insight-engine/artifacts/survey-insight-engine/app.py`) → `FileNotFoundError`
in `setUpClass`, erroring every test in the file.

**Correct form** (matches all existing passing tests): `parents[1] / "app.py"`.

**Why:** verbatim-install workflow — files are cp'd exactly as uploaded and only
explicitly-authorized fixes may be applied. The user's STEP 1.5 sed typically globs only
`test_wizard_*.py`, so any other file with the same bug (e.g. `test_ui_wizard_outcome_laggard.py`)
is left broken and must be reported verbatim, awaiting authorization.

**How to apply:** during install, grep `grep "APP_PATH = Path" tests/*.py` for the doubled
segment across ALL new/modified test files, not just the ones in the authorized sed glob.
Flag any out-of-glob matches to the user before running the full suite.
