---
name: Streamlit must deploy as Reserved VM, not Autoscale
description: Why a Streamlit artifact fails the publish promote/health-check on Autoscale and must use Reserved VM (vm) deployment type.
---

# Streamlit deploys must use Reserved VM (vm), not Autoscale

A Streamlit artifact in this monorepo fails to publish on **Autoscale (cloud_run)** at the
**promote / startup-probe** phase, even after the image-size problem is solved. Symptoms in
the build + runtime logs:

- Build phase succeeds (image pushed), reaches `Creating Autoscale service`, then fails ~5 min later.
- Runtime: `not all artifact ports opened within timeout expected=[8080 21049] detected=1`,
  `a port configuration was specified but the required port was never opened`, repeated
  `healthcheck / returned status 500`. The Node `api-server` (8080) opens its port and serves
  `/api/healthz` 200; the Streamlit service (21049) never binds within the ~60s window.
- The Streamlit process starts (gets a pid) and is still **alive** at the timeout (receives SIGTERM,
  did not crash), and prints **no output** (prod stdout is block-buffered, so messages are lost).

**Root cause (two reinforcing reasons):**
1. Autoscale/Cloud Run **throttles CPU during container startup** (before the port opens). Python/
   Streamlit's heavier bootstrap doesn't finish binding the port in time under that throttle, while
   the lighter Node server does. → startup probe never gets a 200 → promote fails.
2. Streamlit is **fundamentally stateful**: persistent WebSocket per session, in-memory
   `st.session_state`, and local-filesystem file uploads — all of which the deployment skill lists as
   things that do NOT work on Autoscale.

**Fix:** the user must switch the deployment to **Reserved VM (`vm`)** in the Publishing/Deployments
pane. On a VM the CPU is always allocated (no startup throttle, so the port binds fast like in dev)
and state/websockets persist. Reserved VM is a fixed monthly cost vs Autoscale's pay-per-use.

**Why this can't be auto-fixed:** deployment type is NOT changeable programmatically (not via
`artifact.toml`); only the user can change it in the Deployments pane. Streamlit binds its HTTP port
during *server bootstrap*, before `app.py` runs per-session — so "port never opened" means the server
bootstrap (not the app code) didn't finish in time, which points at the platform/CPU-throttle, not a
code bug.

**How to apply:** for any Streamlit (or other stateful/websocket) artifact, recommend Reserved VM
from the start; do not waste cycles trying to make it pass Autoscale's startup probe.
