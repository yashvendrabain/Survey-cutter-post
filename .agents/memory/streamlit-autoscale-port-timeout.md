---
name: Streamlit prod port-never-binds (file watcher), VM logs not retrievable
description: Why a Streamlit artifact fails the publish promote/health-check (port never opens) on BOTH Autoscale and Reserved VM, the real fix, and the diagnosis limits.
---

# Streamlit publish fails: required port never opens (both Autoscale AND Reserved VM)

A Streamlit artifact in this monorepo fails to publish at the **promote / startup-probe**
phase. Symptoms (multi-artifact deploy waits for ports `[8080 21049]`):

- Node `api-server` (8080) binds fast and serves `/api/healthz` 200.
- Streamlit (21049) **never binds** → `not all artifact ports opened within timeout
  detected=1`, `a port configuration was specified but the required port was never opened`,
  repeated `healthcheck / returned status 500` → SIGTERM.
- The streamlit process starts (gets a pid) but emits **no stdout/stderr** (prod stdout is
  block-buffered, so any banner/error is lost).

**CPU-throttle theory is DISPROVEN.** It fails on **both** Autoscale (~60s window) **and** a
full-CPU **Reserved VM** (e2-small; ~8 min "waiting for ready" then fail). 8 minutes on
always-on CPU is far more than enough for a working start (local boots <30s), so it is not
slowness/throttling — Streamlit hangs or crashes indefinitely before binding.

**The exact prod run command works locally.** Running the artifact.toml production `run`
verbatim (on a spare port) binds the port and returns `/_stcore/health` 200 with the normal
banner. So the command + app code + `.pythonlibs` are fine; the failure is
**production-environment-specific**.

**Most likely cause + fix (applied):** Streamlit's default `fileWatcherType = "auto"` sets up
an inotify watcher recursively over a large `src/` tree + a ~9000-line `app.py` during server
bootstrap, *before* Tornado finishes binding. In a constrained prod container this can hang or
hit inotify watch/instance limits. **Fix: add `--server.fileWatcherType none` to the
production (and dev) run command** via `verifyAndReplaceArtifactToml`. Zero behavior risk — a
deployed app needs no hot-reload. Validated locally: health still 200 with the flag.

**If that is insufficient, in order:**
1. Bump the Reserved VM size (e2-small 0.5 vCPU/2GB is underpowered for a pandas/numpy/scipy
   bootstrap). User-only change in the Deployments pane.
2. Formalize Python deps for prod instead of relying on the shipped 723MB `.pythonlibs`
   (root `pyproject.toml` has `dependencies=[]`, and the Node-first build runs no `pip install`,
   so deps reach prod ONLY via the image's `.pythonlibs`). Do NOT bolt on a production
   `pip install -r requirements.txt` blindly — in a Node-first deploy it may install off the
   runtime python path, no-op as "already satisfied," or clobber a working `.pythonlibs`.

**Diagnosis limit (important):** `fetchDeploymentLogs` returns **runtime logs only for
Autoscale (cloud_run)**, NOT for Reserved VM (gce) — a failed VM build surfaces build logs but
no app stdout/stderr. To actually see why a Streamlit container fails, diagnose on Autoscale,
or make startup robust enough to pass blindly.

**Why deployment type can't be auto-fixed:** it is user-only (Deployments pane), not settable
via `artifact.toml`.

## Resolution (confirmed fix — option #2 was the real one)

The `--server.fileWatcherType none` flag alone was **NOT** sufficient. The promote/"failed to
start" failure was the Python deps reaching prod only via the gitignored `.pythonlibs`. The fix
that worked was formalizing deps, NOT a prod `pip install` and NOT a VM size bump:

- Declare the **exact** runtime deps in root `pyproject.toml` pinned to the *actually-installed*
  versions, and regenerate `uv.lock` (`UV_PYTHON=.pythonlibs/bin/python uv lock`).
- Keep `.pythonlibs` realigned to the lock via a **post-merge script** (`realign_python_lock.py`)
  that wipes stale `*.dist-info` and force-reinstalls drifted packages to the locked versions
  (`pip install --user --break-system-packages --no-deps --force-reinstall`). A verify script
  (`verify_python_lock.py`) asserts installed==lock and is registered as a `python-lock`
  validation. See `python-deps-provisioning.md`.
- **Gotcha:** `requirements.txt` pins were *wrong/aspirational* — claimed streamlit 1.58.0 but
  1.39.0 actually runs; openai had 3 layered dist-info dirs (split-brain). Trust what's installed
  (and the lock), never the hand-written requirements.txt.

**Concrete gce diagnostic marker:** `getDeploymentBuild(<failed-id>)` DOES return build logs for
a Reserved VM (gce) build, but a promote failure shows only ~25 lines ending at
`info: Waiting for deployment to be ready` — no app stdout/stderr. That last line == startup
probe never got 200; the real traceback is unreachable on gce (confirms the diagnosis limit).
