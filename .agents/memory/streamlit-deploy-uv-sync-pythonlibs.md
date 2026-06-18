---
name: Python publish EACCES — deploy uv sync writes to read-only nix store
description: Why a Replit publish build fails at "Installing packages" with Permission denied into the nix store, and the fix that actually reaches the deploy build (remove the ROOT uv-project trigger; run prod from the shipped .pythonlibs).
---

# Python publish fails: deploy `uv sync` writes to read-only nix store

**Symptom:** Publish build fails at "Installing packages" — `uv sync` prepares
wheels, then `Failed to install <pkg> ... Permission denied (os error 13)` copying
into `/nix/store/<hash>-python3-.../lib/python3.11/site-packages/`. Errors on an
early leaf pkg (e.g. numpy/iniconfig) but is trying to install the whole set there.

**Root cause / trigger:** This deploy is **Node-first**. The "Installing packages"
build phase runs NO Python install step **except** an auto `uv lock` + `uv sync`
that Replit fires **only when a ROOT `pyproject.toml` + `uv.lock` exist**. That
auto `uv sync` targets `UV_PROJECT_ENVIRONMENT`, which the python module defaults to
`.pythonlibs` — a NON-venv whose `bin/python` symlinks into the read-only nix store,
so uv resolves the nix-store prefix as the target and dies with EACCES.

**Proof of the trigger (build-history bisect):** builds that ran BEFORE a root
`uv.lock` existed show "Installing packages" doing only the Node esbuild step (no
Python install at all) and PASS the build phase; every build AFTER a root
`pyproject.toml`+`uv.lock` were added fails at the auto `uv sync`. The root uv
project is exactly what introduces the failure.

**DISPROVEN fix (do NOT retry): redirecting UV_PROJECT_ENVIRONMENT via `setEnvVars`.**
`setEnvVars({environment:"shared"})` writes `[userenv.shared]` into `.replit`. That
store reaches the **workspace shell** and the **runtime**, but **NOT the deploy
BUILD container** — 5 builds after setting `UV_PROJECT_ENVIRONMENT=.venv` (shared)
failed with the IDENTICAL nix-store EACCES. A prior note claimed "shared scope
reaches the build (confirmed in a fresh shell)"; the fresh-shell test only proves
the *workspace* sees it, which is a different environment from the build container.
uv 0.9.x has no pyproject/uv.toml field for the project env (env-var only), so there
is no committed lever that the build is known to read.

**The fix that WORKS: remove the ROOT uv-project trigger; serve prod from `.pythonlibs`.**
1. Move `pyproject.toml` + `uv.lock` OUT of the repo root into the artifact dir
   (`artifacts/survey-insight-engine/`). No root uv project ⇒ deploy runs no auto
   `uv sync` ⇒ build phase passes (matches the pre-uv-lock builds).
2. Point production at the shipped env: artifact.toml `[services.production].run`
   and `.build` use `/home/runner/workspace/.pythonlibs/bin/python ...` (identical
   to the working `[services.development]` command). The `.build` is an `import`
   smoke-test — keep it so any missing dep fails LOUDLY in the build logs (which ARE
   retrievable for gce) instead of silently hanging at runtime.
3. Do NOT exclude `.pythonlibs` from the image. With no install step, the app's
   Python deps reach prod ONLY via this shipped `.pythonlibs`.

**Why this delivers deps without an install step:** the deploy image is a snapshot
of the workspace filesystem (INCLUDING gitignored files), trimmed only by
`.replitignore`. Evidence: `.replitignore` bothers to exclude gitignored dirs
(`dist`, `.cache`, `.local`) — pointless unless gitignored content ships by default.
So the gitignored, lock-aligned `.pythonlibs` (~700MB) ships and provides the deps.

**The verify/realign scripts need no path change after the move:** their
`repo_root()` walks up from the script's own dir to the first `uv.lock`, so a lock
in the artifact dir auto-resolves; `uv export` runs with `cwd=that dir`.

**Verify before publishing (all local, no publish needed):**
- root has NO `pyproject.toml`/`uv.lock`;
- `.pythonlibs/bin/python artifacts/survey-insight-engine/scripts/verify_python_lock.py` exits 0;
- run the exact prod command on a spare port ⇒ `/_stcore/health` == 200 (~8s here).

**Still-open risk after the build passes:** a separate PROMOTE/startup failure (port
21049 never binds) hit the early pre-uv-lock builds. Mitigations now in place that
those builds lacked: `--server.fileWatcherType none` and a `.pythonlibs` realigned
to the lock. The "formalize deps fixed promote" claim in
`streamlit-autoscale-port-timeout.md` was NEVER actually validated (every build
carrying it died at uv-sync before reaching promote). If promote still fails, the
next levers are bumping the Reserved VM size or temporarily switching to Autoscale
to get runtime logs (gce runtime logs are not retrievable).

**`.replit` editing:** set env vars via `setEnvVars`/`deleteEnvVars` (writes
`[userenv.*]`), and deployment run/build via `verifyAndReplaceArtifactToml`.
