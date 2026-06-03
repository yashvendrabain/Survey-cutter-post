---
name: Streamlit websocket idle timeout on Replit proxy
description: Why long synchronous Streamlit runs "silently fail" after ~30s on Replit, and the heartbeat fix pattern
---

# Streamlit long-run "silent failure" on Replit (~30s)

A Streamlit run that does long (>~30s) synchronous CPU/IO work on the script
thread will appear to "silently fail" in the browser: no error box, terminal
shows only the entry log line, headless repro of the same pipeline works fine.

**Root cause:** During the blocking work no websocket frames flow, so the Replit
reverse proxy hits its idle timeout (~30s) and closes the socket. The browser
fires `WebSocket onerror`, auto-reconnects, and the server starts a *fresh*
session (`Session with id ... is already connected! Connecting to a new
session.`) with empty `session_state` (`run_complete=False`). Nothing renders;
any error from the original run is delivered to a dead socket.

**Why it's not a code crash:** the run's `try/except` never fires (no Python
exception). Confirm by checking there are no `os._exit`/`sys.exit`/`signal`
calls that could bypass it — if none, and no error box appears, suspect the
proxy/websocket drop, not the pipeline.

**Fix pattern (keepalive heartbeat):** run each blocking phase in a
`concurrent.futures.ThreadPoolExecutor(max_workers=1)` worker; on the main
thread poll `future.result(timeout=5)` and call `status.update(...)` on each
`TimeoutError` so websocket frames keep flowing. Key correctness points:
- Worker callables must be pure (no `st.session_state` / Streamlit access);
  only the main thread touches Streamlit.
- `future.result()` **auto-re-raises** worker exceptions, so the existing
  `try/except` around the run still catches real errors — no explicit re-raise,
  no infinite heartbeat loop.

**Why:** large combined survey files make `load_survey_inputs` alone ~57s
(e.g. 1187-column BCN file), well past the idle timeout. Small files finish
before ~30s, which is why the bug only bites large inputs.

**How to apply:** any Streamlit artifact on Replit with a multi-second blocking
action behind a button should emit periodic `status.update()` (or other
websocket traffic) during the work, not run it as one uninterrupted blocking call.

## Heartbeat alone is not enough if a single phase still exceeds the timeout

The real winvslag2024 (107q) silent-kill was NOT load/decoder/memory — it was the
Stage 4 AI enrichment running its three Portkey calls *sequentially*
(themes ~17s + demo ~2s + short_labels ~35.5s ≈ 54s) as one blocking phase. Even
heartbeat-wrapped, a single phase that long risks the kill; the durable fix was to
**parallelize the three calls** (ThreadPoolExecutor max_workers=3, one cache dict
per call — never the shared module cache from worker threads) so the phase drops to
~max(call) and *then* wrap that whole parallel block as the `work=` arg of the
heartbeat. Pattern: shrink the longest phase first, then keepalive-wrap it.

**Latent trap:** the same three AI calls also run sequentially, no heartbeat, with
the shared `_INSIGHT_CACHE`, inside the post-run full-workbook refresh path. Any
interaction that triggers a full refresh can reintroduce the timeout outside the
initial Run-Analysis path. **Why:** only the initial run path was hardened.

**Timeout-vs-latency tension:** bulk-call timeouts were lowered to 25s, but
short_labels was measured at ~35.5s — it can legitimately time out and fall back to
deterministic labels (pipeline still completes, AI quality silently degrades). If
AI label quality matters, raise that timeout (heartbeat already prevents the idle
kill, so a longer per-call timeout is safe).
