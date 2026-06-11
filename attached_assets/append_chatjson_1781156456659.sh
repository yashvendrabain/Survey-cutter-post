#!/usr/bin/env bash
# Append the shared JSON chat helper to ai_insights.py and verify.
# Run from project root. Safe to inspect first; idempotent guard included.
set -u
cd ~/workspace/artifacts/survey-insight-engine || { echo "WRONG DIR"; exit 1; }

TARGET="src/ai_insights.py"
ADD="ai_insights_chatjson_addition.py"   # upload this file to project root first

[ -f "$TARGET" ] || { echo "MISSING $TARGET — are you in the right dir?"; exit 1; }
[ -f "$ADD" ]    || { echo "Upload $ADD to project root first."; exit 1; }

# Idempotency: don't append twice.
if grep -q "def _portkey_chat_json" "$TARGET"; then
  echo "ALREADY PRESENT: _portkey_chat_json exists in $TARGET — skipping append."
else
  cp "$TARGET" "${TARGET}.bak_chatjson"
  cat "$ADD" >> "$TARGET"
  echo "Appended helper to $TARGET (backup: ${TARGET}.bak_chatjson)"
fi

echo "--- verify ---"
grep -c "def _portkey_chat_json" "$TARGET"   # want 1
python3 -c "import ast; ast.parse(open('$TARGET').read()); print('ai_insights.py AST OK')" \
  || { echo 'AST FAIL — revert: cp '"${TARGET}.bak_chatjson"' '"$TARGET"; exit 2; }

echo "--- restart ---"
pkill -9 -f "[s]treamlit"; sleep 3
nohup streamlit run app.py --server.port 21049 --server.address 0.0.0.0 \
  --server.headless true > /tmp/streamlit.log 2>&1 & disown
sleep 10
curl -s -o /dev/null -w "HTTP %{http_code}\n" localhost:21049/
echo "Now the chatbot's LLM phrasing + hypothesis question-mapping are live."
echo "Test: open chat, ask a free-text hypothesis -> should map to 2 questions,"
echo "show a table, and give a hedged verdict."
