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

**Writing into `.pythonlibs` (this is the key trick):** the env is a `--user` site —
`PYTHONUSERBASE=.pythonlibs`. The Nix `pip` refuses installs with
`error: externally-managed-environment` (PEP 668). The working invocation is
`.pythonlibs/bin/python -m pip install --user --break-system-packages ...` — `--user`
targets the writable `.pythonlibs/lib/python3.11/site-packages` (never the read-only
Nix store), and `--break-system-packages` overrides the PEP 668 guard *without*
touching the store. This is how to install/upgrade/force-reinstall deps here.

**The drift was layered stale `dist-info`, not just "uv resolved newer".** Ad-hoc pip
runs overwrote code but left OLD `*.dist-info` dirs behind, so one package had multiple
dist-info versions (e.g. openai had 2.36.0 + 2.40.0 + 2.41.0). `importlib.metadata` /
`uv` then report whichever it scans first → split-brain (openai code 2.41 vs metadata
2.36) and phantom "drift". To make `.pythonlibs` match the lock exactly:
1. `uv export --no-emit-project --format requirements-txt` (omit `--no-dev` to KEEP
   pytest et al. — the deployed env should mirror the *tested* lock), strip hashes →
   `name==ver` list.
2. Remove stale/duplicate `*.dist-info` dirs and any package not in the lock
   (e.g. streamlit-sortables) — for extras, delete files via their `RECORD` first.
3. `python -m pip install --user --break-system-packages --no-deps --force-reinstall
   -r <list>` to rewrite every package's code+metadata to one coherent lock version.
   Run it BACKGROUNDED (`nohup ... &`) — a 66-pkg force-reinstall exceeds the bash
   tool's hard timeout and gets SIGKILLed mid-run, leaving a half-cleaned env.

**Markers matter — do NOT strip them when building the install list.** `uv export`
emits env markers, e.g. `colorama==0.4.6 ; sys_platform == 'win32'` and
`watchdog==5.0.3 ; sys_platform != 'darwin'`. A naive `sed 's/ .*//'` drops the
marker and will install win32-only `colorama` on Linux — a real fidelity break. Keep
the marker and evaluate it (`packaging.markers.Marker(...).evaluate()`) so
platform-gated packages are only installed/expected where they apply. On Linux the
lock has 66 entries but only 65 install (colorama excluded; watchdog included).

**There is a committed, reproducible check:**
`artifacts/survey-insight-engine/scripts/verify_python_lock.py` — run with the project
interpreter (`.pythonlibs/bin/python …`). It uses `uv export` + marker evaluation to
assert installed==lock for every package, flags duplicate dist-info, unexpected
extras, and openai code-vs-metadata split-brain; exits non-zero on any drift. This is
the auditable proof since `.pythonlibs` is gitignored. `requirements.txt` in that
artifact is a marker-preserving snapshot of the same export (mirror, not source of
truth). Verify success also via `uv lock --check` and `/_stcore/health` == 200.

The old "don't force a reinstall on a working app" caution is superseded for a
deliberate lock-realignment like this — `pip --user --break-system-packages
--force-reinstall` is safe here; `uv sync` is still the thing to avoid (wrong python +
clobber).

**There is now an auto-fixer wired into post-merge:**
`artifacts/survey-insight-engine/scripts/realign_python_lock.py` is the *fixer*
counterpart to the verifier — `scripts/post-merge.sh` runs it (with the project
interpreter) after every merge so `.pythonlibs` is realigned to `uv.lock`
automatically. It is idempotent and incremental: it only reinstalls packages that
drift (wrong/missing version or carrying duplicate dist-info), wiping all stale
dist-info layers (files via RECORD) first so one coherent version remains; a clean
env is a sub-5s no-op. **Operational gotcha:** the script clears stale dist-info
BEFORE the force-reinstall, so if the reinstall is interrupted mid-run (bash-tool
hard timeout or a backgrounded process killed when the shell session ends) the
cleared packages end up MISSING — a half-cleaned env. It is idempotent, so just run
it to completion (use `setsid`/post-merge, not a foregrounded bash call) and it
reinstalls the missing ones. The post-merge timeout was bumped to 180s to fit a
worst-case full reinstall.
