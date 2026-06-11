#!/usr/bin/env bash
# =====================================================================
# Survey Analysis Engine — deploy: WS2/WS3 renames + WS4 nav fix
#   + product tour + chatbot wiring.
# The app.py here is ALREADY EDITED and AST-verified. We swap it in
# wholesale (not anchor-patch), so there is no blind-replace risk.
# Run from: ~/workspace/artifacts/survey-insight-engine
# Revert:   cp app.py.bak_revamp app.py   (and restart)
# =====================================================================
set -u
cd ~/workspace/artifacts/survey-insight-engine || { echo "WRONG DIR"; exit 1; }

# ---- 0. Upload checklist (do this in the Replit file pane FIRST) -----
#   Upload to these exact paths:
#     src/assistant_bot.py
#     src/product_tour.py
#     src/chat_panel.py
#     assistant_faq.json          (project root)
#     app.py                      (the EDITED one — overwrites current)
#   (handover_doc.md -> docs/, optional, no wiring)

# ---- 1. Pre-flight: required modules present -------------------------
missing=0
for f in src/assistant_bot.py src/product_tour.py src/chat_panel.py assistant_faq.json app.py; do
  if [ ! -f "$f" ]; then echo "MISSING: $f"; missing=1; fi
done
[ $missing -eq 1 ] && { echo "Upload the missing files, then re-run."; exit 1; }

# ---- 2. Backup current app.py (the one Replit had before this swap) --
# NOTE: only meaningful if you DIDN'T already overwrite via upload.
if [ -f app.py.bak_revamp ]; then echo "backup already exists, keeping it";
else cp app.py app.py.bak_revamp; echo "Backup -> app.py.bak_revamp"; fi

# ---- 3. AST-check everything -----------------------------------------
for f in app.py src/assistant_bot.py src/product_tour.py src/chat_panel.py; do
  python3 -c "import ast; ast.parse(open('$f').read()); print('OK  $f')" || \
    { echo "AST FAIL $f — STOP"; exit 2; }
done
python3 -c "import json; json.load(open('assistant_faq.json')); print('OK  assistant_faq.json')" || exit 2

# ---- 4. §5 protective markers (must read exactly as shown) -----------
echo "--- markers (expected in parens) ---"
printf "%-40s %s\n" "_run_with_status_heartbeat (6):" "$(grep -c '_run_with_status_heartbeat' app.py)"
printf "%-40s %s\n" "ThreadPoolExecutor(max_workers=3) (1):" "$(grep -c 'ThreadPoolExecutor(max_workers=3)' app.py)"
printf "%-40s %s\n" "navjump_ (4):" "$(grep -c 'navjump_' app.py)"
printf "%-40s %s\n" "data-view (4):" "$(grep -c 'data-view' app.py)"
printf "%-40s %s\n" "_nav_onclick (3):" "$(grep -c '_nav_onclick' app.py)"
printf "%-40s %s\n" "maybe_render_tour (2):" "$(grep -c 'maybe_render_tour' app.py)"
printf "%-40s %s\n" "render_chat_panel (2):" "$(grep -c 'render_chat_panel' app.py)"

# ---- 5. CONFIRM the bot's LLM function name --------------------------
echo "--- LLM helper in ai_insights (bot expects _portkey_chat_json) ---"
grep -n "def _portkey_chat_json\|def _raw_chat\|def .*chat.*json\|portkey\|Portkey" src/ai_insights.py | head
echo ">>> If NONE of the above is a JSON-returning chat helper, edit"
echo ">>> src/assistant_bot.py _llm_json() to call the real function."
echo ">>> Until then the bot runs deterministically (routing, survey Qs,"
echo ">>> and hypothesis cross-cut tables all still work)."

# ---- 6. Clean restart ------------------------------------------------
pkill -9 -f streamlit; sleep 3
ps aux | grep "[s]treamlit" | wc -l   # want 0
nohup streamlit run app.py --server.port 21049 --server.address 0.0.0.0 \
  --server.headless true > /tmp/streamlit.log 2>&1 & disown
sleep 10
ps aux | grep "[s]treamlit" | wc -l   # want 1
curl -s -o /dev/null -w "HTTP %{http_code}\n" localhost:21049/

echo
echo "MANUAL CHECKS:"
echo "  1. Click a top nav tab (Upload / Outcome / Download) — view MUST change."
echo "     (this is the WS4 dead-button fix; if it still doesn't move, tell me)"
echo "  2. On a single-cut, toggle 'Single pane view' (was 'Normal view')."
echo "  3. Step 2 wizard shows 'Local Filters' with the explainer caption."
echo "  4. Tour auto-opens once; Skip works; doesn't return on rerun."
echo "  5. Chat bubble bottom-right: 'is there an NPS question?' -> grounded answer;"
echo "     a hypothesis -> returns a TABLE + hedge."
