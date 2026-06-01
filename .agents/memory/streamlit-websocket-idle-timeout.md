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
