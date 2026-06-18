---
name: Python deps provisioning (uv) in this pnpm monorepo
description: How to declare/lock Python deps for the Streamlit artifact without breaking the working .pythonlibs; why installLanguagePackages fails here.
---

# Python dependency provisioning (uv) in this monorepo

The Python toolchain here is `uv` with the runtime venv at `.pythonlibs/` (NOT a
`.venv`). The root `pyproject.toml` + `uv.lock` are the declarative source of truth;
`artifacts/survey-insight-engine/requirements.txt` is secondary documentation that
must mirror the actually-installed versions.

**`installLanguagePackages({language:"python", ...})` FAILS here.** Its `uv add`
resolves the interpreter to the read-only Nix store python
(`/nix/store/.../python3.11/site-packages`) and dies with `Permission denied
(os error 13)` mid-install (e.g. on typing_extensions). On failure `uv add` rolls
back its `pyproject.toml` edit, so no harm — but you cannot use this tool to manage
python deps in this repo.

**Use instead:** edit `pyproject.toml` by hand, then run
`UV_PYTHON=.pythonlibs/bin/python uv lock`. `uv lock` only resolves + writes the
lockfile; it never copies files into site-packages, so it sidesteps the permission
error and does NOT touch the working `.pythonlibs`. Verify with `uv lock --check`.
Avoid `uv sync` — it would rewrite `.pythonlibs` (clobber risk) and target the wrong
python.

**Determinism caveat:** `uv lock` resolves transitives to *latest-compatible*, so the
lock can list newer patch/minor versions than the ad-hoc `.pythonlibs` (which was
pip-populated, not uv-synced). Direct/top-level pins match; ~18 transitives drifted
newer. Publish ships `.pythonlibs` as-is and does NOT run a python install, so the
lock is a reproducible recipe, not what actually ships — the drift is latent, not
breaking.

**Gotchas found in `.pythonlibs`:**
- `streamlit-sortables` is installed but imported nowhere (unused leftover) — don't
  add it to deps.
- `openai` is split-brain: `openai.__version__` == 2.41.0 (the running code) but the
  dist-info metadata == 2.36.0. Pin to 2.41.0 (what actually runs). Don't force a
  reinstall on a working app just to fix the metadata.
