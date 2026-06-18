---
name: Streamlit/Python publish EACCES into nix store (non-venv .pythonlibs)
description: Why a Replit publish build fails at "Installing packages" with Permission denied into the nix store, and the .pythonlibs + run-command fix.
---

# Python publish fails: deploy `uv sync` writes to read-only nix store

**Symptom:** Publish build fails at the "Installing packages" step — `uv sync`
downloads fine, then `Failed to install <pkg> ... Permission denied (os error 13)`
copying into `/nix/store/<hash>-python3-.../lib/python3.11/site-packages/`.

**Root cause:** In a Replit repl, `.pythonlibs` (created by the python module +
`pip install --user`) is NOT a real venv — no `pyvenv.cfg`, and its `bin/python`
symlinks into the read-only nix store. If `.pythonlibs` is shipped into the
deploy image (i.e. NOT excluded by `.replitignore`), the deploy's `uv sync`
(UV_PROJECT_ENVIRONMENT=/home/runner/workspace/.pythonlibs is active in the
build) resolves that interpreter and installs into the nix store → EACCES. Dev
never hits this because dev installs via `pip install --user`, never `uv sync`.

**Fix (3 parts):**
1. Add `.pythonlibs` to `.replitignore`. With no pre-existing dir, the deploy's
   `uv sync` creates a FRESH proper venv there from uv.lock (writable) → passes.
   Bonus: trims the image (here ~719MB).
2. Run production via the venv's own interpreter, not bare PATH: in artifact.toml
   `[services.production].run` use `/home/runner/workspace/.pythonlibs/bin/python
   -m streamlit run app.py ...`. Absolute path is correct — the build env uses
   that exact path. Removes PATH dependence.
3. Optional fail-fast gate: artifact.toml `[services.production].build` =
   `<venv python> -c 'import streamlit, pandas, ...'` runs AFTER the repo-level
   uv sync; fails the build early with a clear ImportError instead of an opaque
   promote failure.

**Why / gotchas:**
- uv 0.9.x does NOT accept `project-environment` in `uv.toml` (unknown field) —
  the project env is env-var-only. But the deploy already sets
  UV_PROJECT_ENVIRONMENT=.pythonlibs, so excluding the dir is enough.
- `.replit` cannot be edited directly (blocked); deployment run/build live in
  `artifact.toml` for pnpm monorepos, edited via `verifyAndReplaceArtifactToml`.
- Local sanity check: `UV_PROJECT_ENVIRONMENT=/tmp/fresh uv sync --frozen
  --no-dev` should create a venv with a working `bin/streamlit`.
