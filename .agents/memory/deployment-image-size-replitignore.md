---
name: Deployment image size & .replitignore
description: Why a publish can fail with "image size is over the limit of 8 GiB" and how .replitignore (not .gitignore) is the fix.
---

# Deployment image over 8 GiB

A Replit publish can fail at the very end of the build (after compilation/push succeeds) with:
`error: image size is over the limit of 8 GiB`. This is a packaging-size failure, NOT a code or health-check failure — the build logs show everything succeeding up to the final image upload.

**Why:** The deployment image is built from the repl **filesystem**, and `.gitignore` does NOT exclude files from that image. Only **`.replitignore`** (dockerignore format) trims the deployed image. So large gitignored dirs (`dist/`, `.cache/`, `node_modules`, build archives) are still baked into the image unless listed in `.replitignore`.

**How to apply:** When a publish fails on image size, fetch the failed build logs (`listDeploymentBuilds` → `getDeploymentBuild`) to confirm the 8 GiB message, then measure on-disk footprint and add non-runtime bloat to `.replitignore`. Keep runtime deps (`.pythonlibs`, app `src`, `sample_data`, needed `node_modules`); exclude generated archives, caches, `.git`, test fixtures, scratch `outputs/`, uploaded `attached_assets/`, and editor backups. `.replitignore` only affects the image — it does not delete files from disk and does not affect dev workflows.

**Project-specific (survey-insight-engine):** the dominant culprit was a multi-GB `dist/survey-insight-engine-project.zip` (a self-referential project download bundle) plus a ~675 MB `tests/` fixture dir and ~319 MB `outputs/` scratch dir. `du -sh .` at repo root hangs/times out — measure per-directory with a `timeout` wrapper instead.
