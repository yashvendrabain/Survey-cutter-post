---
name: Streamlit orphan process blocks workflow restart
description: restart_workflow fails TASK_FAILED when a stale streamlit holds the port; kill the orphan first
---

When the survey-insight-engine workflow ("artifacts/survey-insight-engine: web") shows FAILED and `restart_workflow` fails with `TASK_FAILED ... had failing tasks / Consider optimizing resource-intensive startup`, the usual cause is NOT a code fault. A stale orphan `streamlit run app.py` process from an earlier session (check `ps -eo pid,etime` — etime of several hours) is still bound to the server port (21049) and serving HTTP 200. The supervisor's fresh managed process cannot bind the port, so it marks its own task failed.

**Why:** Streamlit prints its "You can now view your app" banner immediately at launch (before the script runs), so the supervisor log shows a clean banner with no traceback even though the new process never bound. The orphan keeps answering curl on the port, masking the problem.

**How to apply (diagnose → repair):**
1. `ps -eo pid,etime,cmd | grep '[s]treamlit run app.py'` — a long etime (hours) = orphan from a prior session.
2. A manual `streamlit run` will say `Port 21049 is already in use`, confirming the orphan.
3. `kill <pid>` (SIGKILL if needed), confirm port free (`curl --max-time 4` returns 000), then `restart_workflow` — it now binds cleanly (fresh process has a young etime).

**Extra risk:** an hours-old orphan re-execs app.py on each rerun (picks up new app.py) but keeps OLD `src/*` modules cached in sys.modules — so src-layer changes silently never load. Always kill+clean-restart after installing src changes, never rely on the orphan. Do NOT use pkill/nohup as the serving mechanism (creates the orphan in the first place); use restart_workflow.
