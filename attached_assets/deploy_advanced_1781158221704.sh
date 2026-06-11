#!/usr/bin/env bash
# Deploy: advanced segmentation v2 (generic Top group/Rest, 3 named metrics,
# intersection mode) + the winner_scoring.py intersection engine support.
# Run from project root.  Revert: restore the two .bak files + restart.
set -u
cd ~/workspace/artifacts/survey-insight-engine || { echo "WRONG DIR"; exit 1; }

# Upload these two to project root first, then this script moves them in:
#   winner_scoring.py            -> src/winner_scoring.py
#   advanced_segmentation_ui.py  -> src/advanced_segmentation_ui.py
for f in winner_scoring.py advanced_segmentation_ui.py; do
  [ -f "$f" ] || { echo "Upload $f to project root first."; exit 1; }
done

cp src/winner_scoring.py           src/winner_scoring.py.bak_intersection 2>/dev/null
cp src/advanced_segmentation_ui.py src/advanced_segmentation_ui.py.bak_v2 2>/dev/null
cp winner_scoring.py            src/winner_scoring.py
cp advanced_segmentation_ui.py  src/advanced_segmentation_ui.py
echo "Placed. Backups: *.bak_intersection / *.bak_v2"

echo "--- AST ---"
for f in src/winner_scoring.py src/advanced_segmentation_ui.py; do
  python3 -c "import ast; ast.parse(open('$f').read()); print('OK', '$f')" || { echo "AST FAIL $f"; exit 2; }
done

echo "--- engine sanity: default mode unchanged, intersection present ---"
grep -c 'combination_mode' src/winner_scoring.py            # want >=3
grep -c 'def _intersection_cohorts' src/winner_scoring.py   # want 1
grep -c 'Top group' src/advanced_segmentation_ui.py         # want >0 (generic relabel)
grep -c 'winners\|laggards\|Winners\|Laggards' src/advanced_segmentation_ui.py  # want 0 in UI strings (ids/masks ok)

echo "--- restart ---"
pkill -9 -f "[s]treamlit"; sleep 3
nohup streamlit run app.py --server.port 21049 --server.address 0.0.0.0 \
  --server.headless true > /tmp/streamlit.log 2>&1 & disown
sleep 10
curl -s -o /dev/null -w "HTTP %{http_code}\n" localhost:21049/
tail -15 /tmp/streamlit.log

echo
echo "MANUAL VALIDATION (synthetic green != trusted; verify on REAL survey):"
echo "  1. Outcome Segmentation -> Open Advanced. UI says 'Top group'/'Rest', no winners/laggards."
echo "  2. Three slots: Revenue, Gross Margin, Custom — each picks a question."
echo "  3. A '40%+' band shows 45% by default and is editable."
echo "  4. Combination mode defaults to 'All measures at once (strict)'."
echo "  5. Charts: Top group red vs Rest grey, per measure; breakdown by Sector/Region works."
echo "  6. Sanity-check the Top-group COUNT by hand against the data for one config."
