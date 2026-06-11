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
        "body": "Upload your survey file to begin. Go ahead and add it now — "
                "I\u2019ll move on automatically once it\u2019s in.",
        "selector": "[data-testid=\"stFileUploader\"]",
        "gate": "upload",
    },
    {
        "title": "Step 2 — Decide the structure of your Excel output",
        "body": "Just below the upload, you decide how the workbook is organised: "
                "questions grouped into themed sheets (Pricing, Brand, …) or all on "
                "one sheet. Scroll down to see the two options — pick whichever fits.",
        "selector": "",
        "scroll_text": "how should your Excel output",
    },
    {
        "title": "Step 3 — Categories, filter count & cross-cuts",
        "body": "A short 3-step setup follows: Categories (group/rename questions), "
                "then the number of Local Filters to show atop each sheet, then how "
                "cross-cut sheets are laid out. Walk through it, then scroll back up to Run.",
        "selector": "",
        "scroll_text": "Categories",
    },
    {
        "title": "Step 4 — Run the analysis",
        "body": "When you're set, click Run analysis (near the top of the upload "
                "section). Every single cut is computed in seconds. Click it to continue.",
        "selector": "",
        "gate": "run",
        "find_text": "Run analysis",
    },
    {
        "title": "Step 5 — Single cuts",
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
        "title": "Step 6 — Download",
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


def _tour_html(steps: list[dict[str, str]], force: bool = False) -> str:
    steps_json = json.dumps(steps)
    return """
<script>
(function(){
  var pdoc = window.parent.document;
  if (pdoc && pdoc.head && !pdoc.getElementById("sae-tour-style")) {
    var sty = pdoc.createElement("style"); sty.id="sae-tour-style";
    sty.textContent =
      "#sae-tour-card{position:fixed;width:440px;max-width:92vw;background:#fff;border-radius:14px;"+
      "box-shadow:0 18px 50px rgba(0,0,0,0.28);font-family:Arial,sans-serif;border-top:4px solid #CC0000;"+
      "z-index:2000001;overflow:hidden;}"+
      "#sae-tour-handle{display:flex;align-items:center;gap:8px;padding:9px 16px 5px;cursor:move;"+
      "user-select:none;color:#BBB;font-size:11px;}"+
      "#sae-tour-handle .grip{font-size:13px;color:#CC0000;}"+
      "#sae-tour-body{padding:4px 22px 18px;}"+
      "#sae-tour-card h4{margin:0 0 8px;font-size:16px;color:#0A0A0A;}"+
      "#sae-tour-card p{margin:0 0 14px;font-size:13px;line-height:1.55;color:#333;}"+
      "#sae-tour-bar{display:flex;align-items:center;justify-content:space-between;}"+
      "#sae-tour-dots{font-size:11px;color:#888;}"+
      ".sae-tour-btn{border:none;border-radius:7px;font-size:12px;font-weight:700;padding:8px 16px;"+
      "cursor:pointer;font-family:Arial,sans-serif;}"+
      ".sae-tour-primary{background:#CC0000;color:#fff;}"+
      ".sae-tour-ghost{background:#F0F0F0;color:#333;margin-right:8px;}"+
      ".sae-tour-skip{background:none;color:#888;text-decoration:underline;font-size:11px;border:none;cursor:pointer;}";
    pdoc.head.appendChild(sty);
  }
})();
</script>
<script>
(function(){
  var steps = __STEPS__;
  var doc = window.parent.document;
  // Persist across Streamlit reruns (which re-inject this script).
  if (typeof window.__saeStep !== "number") window.__saeStep = 0;
  if (__FORCE__){ window.__saeDone=false; window.__saeStep=0; window.__saeStarted=false; window.__saePos=null; }

  function remove(){ var c=doc.getElementById("sae-tour-card"); if(c) c.remove(); }

  function show(){
    if (window.__saeDone) { remove(); return; }
    var i = window.__saeStep;
    if (i < 0) i = 0;
    if (i >= steps.length){ window.__saeDone=true; remove(); return; }
    var step = steps[i];
    remove();

    var last = (i === steps.length-1);
    var card = doc.createElement("div"); card.id="sae-tour-card";
    card.innerHTML =
      '<div id="sae-tour-handle"><span class="grip">&#9776;</span><span>Guided tour &mdash; drag to move</span></div>'+
      '<div id="sae-tour-body">'+
      '<h4>'+step.title+'</h4><p>'+step.body+'</p>'+
      '<div id="sae-tour-bar"><span id="sae-tour-dots">'+(i+1)+' / '+steps.length+'</span>'+
      '<span><button class="sae-tour-btn sae-tour-ghost" id="sae-b">Back</button>'+
      '<button class="sae-tour-btn sae-tour-primary" id="sae-n">'+(last?"Done":"Next")+'</button></span></div>'+
      '<div style="margin-top:10px;"><button class="sae-tour-skip" id="sae-s">Skip tour</button></div>'+
      '</div>';
    doc.body.appendChild(card);

    // position: centre until first Next, then bottom-centre; honour dragged pos.
    var w = card.offsetWidth||440, h = card.offsetHeight||200;
    var vw = window.parent.innerWidth, vh = window.parent.innerHeight;
    if (window.__saePos){ card.style.left=window.__saePos.x+"px"; card.style.top=window.__saePos.y+"px"; }
    else if (!window.__saeStarted){ card.style.left=Math.max(8,(vw-w)/2)+"px"; card.style.top=Math.max(8,(vh-h)/2)+"px"; }
    else { card.style.left=Math.max(8,(vw-w)/2)+"px"; card.style.top=Math.max(8,vh-h-28)+"px"; }

    // drag
    var hd=doc.getElementById("sae-tour-handle"), drag=false, ox=0, oy=0;
    hd.addEventListener("mousedown", function(e){ drag=true; var r=card.getBoundingClientRect(); ox=e.clientX-r.left; oy=e.clientY-r.top; e.preventDefault(); });
    doc.addEventListener("mousemove", function(e){ if(!drag) return;
      var x=Math.max(4,Math.min(e.clientX-ox, window.parent.innerWidth-card.offsetWidth-4));
      var y=Math.max(4,Math.min(e.clientY-oy, window.parent.innerHeight-card.offsetHeight-4));
      card.style.left=x+"px"; card.style.top=y+"px"; window.__saePos={x:x,y:y}; });
    doc.addEventListener("mouseup", function(){ drag=false; });

    doc.getElementById("sae-b").onclick=function(){ if(window.__saeStep>0){ window.__saeStep--; show(); } };
    doc.getElementById("sae-n").onclick=function(){
      window.__saeStarted=true;
      if(window.__saeStep < steps.length-1){ window.__saeStep++; show(); }
      else { window.__saeDone=true; remove(); try{window.parent.sessionStorage.setItem("sae_tour_seen","1");}catch(e){} }
    };
    doc.getElementById("sae-s").onclick=function(){
      window.__saeDone=true; remove(); try{window.parent.sessionStorage.setItem("sae_tour_seen","1");}catch(e){}
    };
  }

  if (__FORCE__){ setTimeout(show,120); setTimeout(show,450); return; }
  if (window.__saeDone) return;
  if (doc.getElementById("sae-tour-card")) return;
  var seen=null; try{ seen=window.parent.sessionStorage.getItem("sae_tour_seen"); }catch(e){}
  if(!seen || window.__saeStep>0){ setTimeout(show,300); }
})();
</script>
""".replace("__STEPS__", steps_json).replace("__FORCE__", "true" if force else "false")

def maybe_render_tour(force: bool = False) -> None:
    """Render the tour unless the user has already seen/skipped it this session.

    Pass force=True (e.g. from a 'Restart tour' button) to show it again.
    """
    try:
        import streamlit as st
        from streamlit.components.v1 import html as components_html
    except Exception:
        return

    forced = bool(force) or bool(st.session_state.get(_TOUR_FORCE_KEY))
    if st.session_state.get(_TOUR_SEEN_KEY) and not forced:
        return
    st.session_state[_TOUR_SEEN_KEY] = True
    st.session_state[_TOUR_FORCE_KEY] = False
    # Forced tours need more height so the injected iframe definitely executes JS.
    components_html(_tour_html(TOUR_STEPS, force=forced), height=1)


def restart_tour() -> None:
    """Mark the tour to force-open on the next render (handled by maybe_render_tour)."""
    import streamlit as st
    st.session_state[_TOUR_FORCE_KEY] = True
