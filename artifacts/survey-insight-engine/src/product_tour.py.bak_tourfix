"""Survey Analysis Engine — optional interactive product tour.

Self-contained so app.py only needs:  from src.product_tour import maybe_render_tour
then a single call `maybe_render_tour()` near the top of main() (after CSS inject).

The tour is an overlay walkthrough (spotlight + tooltip card) driven entirely
client-side, so stepping next/back does NOT trigger Streamlit reruns. It is
optional: a "Skip tour" control dismisses it, and a session flag stops it from
reappearing. A "Restart tour" helper is exposed for a Help-menu button.

Steps are data-driven (TOUR_STEPS). Each targets a CSS selector already present
in the app (section anchors, nav tabs, the sidebar). If a target isn't on the
current view the step auto-centres so the tour never breaks.
"""

from __future__ import annotations

import json
from typing import Any

_TOUR_SEEN_KEY = "product_tour_seen"
_TOUR_FORCE_KEY = "product_tour_force"

# Each step: title, body (plain, new-user friendly), and a CSS selector to spotlight.
TOUR_STEPS: list[dict[str, str]] = [
    {
        "title": "Welcome to the Survey Analysis Engine",
        "body": "This quick tour shows the four steps to turn a survey into a "
                "consultant-ready workbook. It takes about a minute. You can skip "
                "anytime.",
        "selector": "",
    },
    {
        "title": "Step 1 — Upload your survey",
        "body": "Drop your raw data and data map (or a combined Excel, or raw data "
                "plus a Word questionnaire). The tool figures out the format for you.",
        "selector": "#section-upload",
    },
    {
        "title": "Step 2 — Decide your output structure",
        "body": "Choose whether questions are grouped into themed sheets or kept on "
                "one, and set up Local Filters — slicers that let a reader filter a "
                "sheet (e.g. show only APAC) right inside Excel.",
        "selector": "#section-filter",
    },
    {
        "title": "Step 3 — Single cuts",
        "body": "Every question gets a 'single cut' — the breakdown of how people "
                "answered, with the standout values highlighted. Pick a question in "
                "the sidebar to explore it.",
        "selector": "#section-singlecuts",
    },
    {
        "title": "Cross cuts",
        "body": "Relate two questions — for example, average revenue growth by "
                "region. Build one beside any question, or from the Cross cuts screen.",
        "selector": "#section-crosscuts",
    },
    {
        "title": "Outcome Segmentation — Winners vs Laggards",
        "body": "Pick what 'success' means (e.g. revenue growth), split respondents "
                "into Winners and Laggards, and the tool finds the questions where "
                "the two groups differ most. Every number is computed, not guessed.",
        "selector": "#section-outcome",
    },
    {
        "title": "Step 4 — Download",
        "body": "Get your workbook anytime after running the analysis. Open it in "
                "Excel desktop so the live filters work.",
        "selector": "#section-downloads",
    },
    {
        "title": "Ask the assistant",
        "body": "The chat button (bottom-right) answers how-to questions, tells you "
                "what's in your survey, and can test a hypothesis against your data — "
                "always showing the table behind its answer.",
        "selector": "",
    },
]


def _tour_html(steps: list[dict[str, str]]) -> str:
    steps_json = json.dumps(steps)
    return """
<div id="sae-tour-root"></div>
<style>
#sae-tour-overlay { position:fixed; inset:0; z-index:2000000; background:rgba(10,10,10,0.55); }
#sae-tour-card { position:fixed; max-width:340px; background:#fff; border-radius:12px;
  box-shadow:0 18px 50px rgba(0,0,0,0.3); padding:18px 20px; font-family:Arial,sans-serif;
  border-top:4px solid #CC0000; z-index:2000001; }
#sae-tour-card h4 { margin:0 0 8px; font-size:15px; color:#0A0A0A; }
#sae-tour-card p { margin:0 0 14px; font-size:13px; line-height:1.55; color:#333; }
#sae-tour-bar { display:flex; align-items:center; justify-content:space-between; }
#sae-tour-dots { font-size:11px; color:#888; }
.sae-tour-btn { border:none; border-radius:7px; font-size:12px; font-weight:700;
  padding:8px 14px; cursor:pointer; font-family:Arial,sans-serif; }
.sae-tour-primary { background:#CC0000; color:#fff; }
.sae-tour-ghost { background:#F0F0F0; color:#333; margin-right:8px; }
.sae-tour-skip { background:none; color:#888; text-decoration:underline; font-size:11px;
  border:none; cursor:pointer; }
.sae-tour-spot { position:fixed; z-index:2000000; border:3px solid #CC0000; border-radius:8px;
  box-shadow:0 0 0 9999px rgba(10,10,10,0.55); pointer-events:none; transition:all 0.2s ease; }
</style>
<script>
(function(){
  var steps = __STEPS__;
  var doc = window.parent.document;
  var i = 0;
  function clear(){ ["sae-tour-overlay","sae-tour-card","sae-tour-spotbox"].forEach(function(id){
    var e=doc.getElementById(id); if(e) e.remove(); }); }
  function done(){ clear();
    try { window.parent.sessionStorage.setItem("sae_tour_seen","1"); } catch(e){}
  }
  function render(){
    clear();
    var step = steps[i];
    var card = doc.createElement("div"); card.id="sae-tour-card";
    var spot=null, tgt = step.selector ? doc.querySelector(step.selector) : null;
    if(tgt){
      var r = tgt.getBoundingClientRect();
      spot = doc.createElement("div"); spot.id="sae-tour-spotbox"; spot.className="sae-tour-spot";
      spot.style.top=(r.top-6)+"px"; spot.style.left=(r.left-6)+"px";
      spot.style.width=(r.width+12)+"px"; spot.style.height=(r.height+12)+"px";
      doc.body.appendChild(spot);
      card.style.top = Math.min(r.bottom+14, window.parent.innerHeight-220)+"px";
      card.style.left = Math.min(r.left, window.parent.innerWidth-360)+"px";
    } else {
      var ov = doc.createElement("div"); ov.id="sae-tour-overlay"; doc.body.appendChild(ov);
      card.style.top="50%"; card.style.left="50%"; card.style.transform="translate(-50%,-50%)";
    }
    card.innerHTML =
      '<h4>'+step.title+'</h4><p>'+step.body+'</p>'+
      '<div id="sae-tour-bar"><span id="sae-tour-dots">'+(i+1)+' / '+steps.length+'</span>'+
      '<span><button class="sae-tour-btn sae-tour-ghost" id="sae-tour-back">Back</button>'+
      '<button class="sae-tour-btn sae-tour-primary" id="sae-tour-next">'+
      (i===steps.length-1?"Done":"Next")+'</button></span></div>'+
      '<div style="margin-top:10px;"><button class="sae-tour-skip" id="sae-tour-skip">Skip tour</button></div>';
    doc.body.appendChild(card);
    doc.getElementById("sae-tour-back").onclick=function(){ if(i>0){i--;render();} };
    doc.getElementById("sae-tour-next").onclick=function(){ if(i<steps.length-1){i++;render();} else {done();} };
    doc.getElementById("sae-tour-skip").onclick=function(){ done(); };
  }
  // Only auto-start if not seen this session.
  var seen=null; try { seen=window.parent.sessionStorage.getItem("sae_tour_seen"); } catch(e){}
  if(!seen){ setTimeout(render, 400); }
})();
</script>
""".replace("__STEPS__", steps_json)


def maybe_render_tour(force: bool = False) -> None:
    """Render the tour unless the user has already seen/skipped it this session.

    Pass force=True (e.g. from a 'Restart tour' button) to show it again.
    """
    try:
        import streamlit as st
        from streamlit.components.v1 import html as components_html
    except Exception:
        return

    if st.session_state.get(_TOUR_SEEN_KEY) and not force and not st.session_state.get(_TOUR_FORCE_KEY):
        return
    st.session_state[_TOUR_SEEN_KEY] = True
    st.session_state[_TOUR_FORCE_KEY] = False
    components_html(_tour_html(TOUR_STEPS), height=0)


def restart_tour() -> None:
    """Call from a Help-menu button to replay the tour."""
    import streamlit as st
    st.session_state[_TOUR_FORCE_KEY] = True
    # Clear the client-side seen flag so the script auto-starts again.
    try:
        from streamlit.components.v1 import html as components_html
        components_html(
            "<script>try{window.parent.sessionStorage.removeItem('sae_tour_seen');}catch(e){}</script>",
            height=0,
        )
    except Exception:
        pass
