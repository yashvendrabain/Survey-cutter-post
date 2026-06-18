---
name: Python publish EACCES — deploy uv sync writes to read-only nix store
description: Why a Replit publish build fails at "Installing packages" with Permission denied into the nix store, and the real fix (redirect UV_PROJECT_ENVIRONMENT to a writable .venv).
---

# Python publish fails: deploy `uv sync` writes to read-only nix store

**Symptom:** Publish build fails at "Installing packages" — `uv sync` prepares
wheels fine, then `Failed to install <pkg> ... Permission denied (os error 13)`
copying into `/nix/store/<hash>-python3-.../lib/python3.11/site-packages/`. It
errors on an early leaf package (e.g. iniconfig) but is actually trying to
install the WHOLE selected set there.

**DEBUG FIRST — is the failing build even using the fix?** A Replit publish builds
from the *committed git tree* at publish time. Before re-diagnosing, compare each
failed build's timestamp (`listDeploymentBuilds`) with the fix commit's timestamp
(`git show -s --format=%cI <sha>`). Builds that predate the fix commit fail with the
OLD behavior and are NOT evidence the fix is broken — the cure is simply to re-publish
so the build picks up the committed fix. (Real case: 6 builds all failed with the
nix-store EACCES; every one ran 20+ min BEFORE the fix was committed.)

**The fix is validated by a REAL `uv sync` (not just `--dry-run`):** with
`UV_PROJECT_ENVIRONMENT=/home/runner/workspace/.venv` active, `uv sync` exits 0,
creates `.venv` (`pyvenv.cfg home = .../.pythonlibs/bin`), and
`.venv/bin/python -c 'import streamlit,pandas,...'` (the artifact.toml prod build
check) succeeds. `.venv` and `.pythonlibs` are gitignored via `/etc/.gitignore`, so a
local simulation sync cannot be accidentally committed/bloat the image.

**Root cause:** When a root `pyproject.toml` + `uv.lock` exist, the Replit deploy
auto-runs a bare `uv lock` + `uv sync`. The python module injects
`UV_PROJECT_ENVIRONMENT=/home/runner/workspace/.pythonlibs`. `.pythonlibs` is NOT
a real venv (no `pyvenv.cfg`; its `bin/python` symlinks into the read-only nix
store). uv therefore resolves the *nix store prefix* as the target env and tries
to install there → EACCES. The nix python does NOT have the project deps "baked
in" — `uv sync --dry-run` plans to install the full set (here 62 prod / 65 with
dev) regardless. Dev never hits this because dev installs via `pip install
--user` into `.pythonlibs/lib`, never `uv sync`.

**The fix that WORKS: redirect the project env to a writable venv.**
1. `setEnvVars({ values: { UV_PROJECT_ENVIRONMENT: "/home/runner/workspace/.venv" }, environment: "shared" })`.
   This writes `[userenv.shared]` into `.replit` and OVERRIDES the module's
   default (verify: a fresh `bash -lc 'echo $UV_PROJECT_ENVIRONMENT'` prints the
   new value). The deploy build inherits the same env store, so its auto
   `uv sync` now creates/uses the writable `.venv`.
2. Point production at that venv: artifact.toml `[services.production].run` and
   `.build` use `/home/runner/workspace/.venv/bin/python ...`. Keep
   `[services.development]` on `.pythonlibs/bin/python` (dev is unchanged).
3. Do NOT exclude `.pythonlibs` from the image. The `.venv` uv creates from
   `.pythonlibs/bin/python3` records `home = .../.pythonlibs/bin` and its
   `bin/python` symlinks to `.pythonlibs/bin/python3`; stripping `.pythonlibs`
   breaks `.venv/bin/python` at runtime. (`sys.base_prefix` = the always-present
   nix-store python, but the intermediate `.pythonlibs` symlink must survive.)

**Verify before publishing:** with the env active, `uv sync --dry-run --frozen`
must say `Would create project environment at: .venv` and show NO `/nix/store`
target.

**Disproven dead-ends (do not retry):**
- Adding `.pythonlibs` to `.replitignore` does NOTHING for this. The python
  module re-provisions `.pythonlibs` as a non-venv in the build (it is gitignored,
  so not shipped from git anyway); `.replitignore` only trims the final image.
- `[tool.uv] default-groups = []` does NOT fix it, and (empirically, uv 0.9.5) does
  not even drop the dev group: a fresh `uv sync` — even `uv sync --no-default-groups` —
  still installs pytest/iniconfig/pluggy. So it is purely cosmetic here, NOT a leaner
  prod venv. Harmless (the install succeeds into the writable `.venv` regardless), but
  do not rely on it to exclude dev deps. Dev verify/realign scripts still pass
  `uv export --all-groups` so `.pythonlibs` tracks the full lock (incl pytest).

**Why / gotchas:**
- uv 0.9.x does NOT accept `project-environment` in `uv.toml`/pyproject (unknown
  field) — the project env is env-var-only, so `UV_PROJECT_ENVIRONMENT` is the
  only lever.
- `.replit` cannot be hand-edited (blocked); set env vars via `setEnvVars` (which
  writes `[userenv.*]`) and deployment run/build via `verifyAndReplaceArtifactToml`.
- A user-set env-store value WINS over the module-injected default (confirmed in a
  fresh shell), which is why `shared` scope reaches the build.
