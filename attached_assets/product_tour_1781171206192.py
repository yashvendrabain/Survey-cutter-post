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
    var sty = pdoc.createElement("style"); sty.id = "sae-tour-style";
    sty.textContent =
      "#sae-tour-card{position:fixed;width:460px;max-width:92vw;background:#fff;border-radius:14px;"+
      "box-shadow:0 18px 50px rgba(0,0,0,0.28);font-family:Arial,sans-serif;"+
      "border-top:4px solid #CC0000;z-index:2000001;overflow:hidden;}"+
      "#sae-tour-handle{display:flex;align-items:center;gap:8px;padding:9px 16px 5px;"+
      "cursor:move;user-select:none;color:#BBB;font-size:11px;}"+
      "#sae-tour-handle .grip{letter-spacing:2px;font-size:13px;color:#CC0000;}"+
      "#sae-tour-body{padding:4px 22px 18px;}"+
      "#sae-tour-card h4{margin:0 0 8px;font-size:16px;color:#0A0A0A;}"+
      "#sae-tour-card p{margin:0 0 14px;font-size:13px;line-height:1.55;color:#333;}"+
      "#sae-tour-bar{display:flex;align-items:center;justify-content:space-between;}"+
      "#sae-tour-dots{font-size:11px;color:#888;}"+
      ".sae-tour-btn{border:none;border-radius:7px;font-size:12px;font-weight:700;"+
      "padding:8px 16px;cursor:pointer;font-family:Arial,sans-serif;}"+
      ".sae-tour-primary{background:#CC0000;color:#fff;}"+
      ".sae-tour-ghost{background:#F0F0F0;color:#333;margin-right:8px;}"+
      ".sae-tour-skip{background:none;color:#888;text-decoration:underline;font-size:11px;"+
      "border:none;cursor:pointer;}"+
      "#sae-tour-arrow{position:fixed;z-index:2000000;pointer-events:none;font-size:52px;line-height:1;"+
      "color:#CC0000;filter:drop-shadow(0 2px 4px rgba(0,0,0,0.3));"+
      "animation:saeArrowBounce 1s ease-in-out infinite;}"+
      "@keyframes saeArrowBounce{0%,100%{opacity:0.5;}50%{opacity:1;}}";
    pdoc.head.appendChild(sty);
  }
})();
</script>
<script>
(function(){
  var steps = __STEPS__;
  var doc = window.parent.document;
  // CRITICAL: Streamlit re-injects this whole script on every rerun (uploads,
  // the download bar, the chat panel all trigger reruns). A local `var i = 0`
  // therefore RESETS the tour to step 0 on every rerun and spawns duplicate
  // handlers — which is why it appeared "stuck after step 2". Persist step +
  // done-state on window so a remount RESUMES the current step.
  if (typeof window.__saeTourStep !== "number") window.__saeTourStep = 0;
  function getI(){ return window.__saeTourStep; }
  function setI(v){ window.__saeTourStep = v; }
  function clear(){
    ["sae-tour-card","sae-tour-arrow"].forEach(function(id){
      var e=doc.getElementById(id); if(e) e.remove();
    });
    if (window.__saeTourPoll){ clearInterval(window.__saeTourPoll); window.__saeTourPoll=null; }
  }
  function done(){ clear();
    window.__saeTourDone = true;
    try { window.parent.sessionStorage.setItem("sae_tour_seen","1"); } catch(e){}
  }
  function findTarget(step){
    var tgt = step.selector ? doc.querySelector(step.selector) : null;
    if (!tgt && step.find_text){
      var bb = doc.querySelectorAll("button");
      for (var bi=0; bi<bb.length; bi++){
        if ((bb[bi].textContent||"").trim().toLowerCase().indexOf(step.find_text.toLowerCase())!==-1){ tgt=bb[bi]; break; }
      }
    }
    if (!tgt && step.scroll_text){
      var all = doc.querySelectorAll("h1,h2,h3,p,div,span,label");
      for (var si=0; si<all.length; si++){
        if ((all[si].textContent||"").toLowerCase().indexOf(step.scroll_text.toLowerCase())!==-1){ tgt=all[si]; break; }
      }
    }
    return tgt;
  }
  function drawArrow(tgt){
    var old = doc.getElementById("sae-tour-arrow"); if(old) old.remove();
    var vw = window.parent.innerWidth, vh = window.parent.innerHeight;
    var hasTarget = tgt && (function(){ var rr=tgt.getBoundingClientRect();
      return !(rr.width===0 && rr.height===0); })();
    if (!hasTarget) return;   // info step: no target -> no arrow (card text is enough)
    var arrow = doc.createElement("div"); arrow.id="sae-tour-arrow";
    var r = tgt.getBoundingClientRect();
    var cx = r.left + r.width/2;
    var cy = r.top + r.height/2;
    if (cy < vh*0.6){
      arrow.textContent = "\u2B07";  // points DOWN onto a target in the upper area
      arrow.style.left = Math.min(Math.max(8, cx-26), vw-60)+"px";
      arrow.style.top  = Math.max(8, r.top-58)+"px";
    } else {
      arrow.textContent = "\u2B06";  // points UP onto a target lower on screen
      arrow.style.left = Math.min(Math.max(8, cx-26), vw-60)+"px";
      arrow.style.top  = Math.min(vh-66, r.bottom+10)+"px";
    }
    doc.body.appendChild(arrow);
  }
  function render(){
    clear();
    var i = getI();
    if (i >= steps.length){ done(); return; }
    var step = steps[i];
    var tgt = findTarget(step);
    if (tgt){ try { tgt.scrollIntoView({behavior:"smooth", block:"center"}); } catch(e){} }
    setTimeout(function(){ drawArrow(findTarget(step)); }, 380);
    setTimeout(function(){ drawArrow(findTarget(step)); }, 900);

    var gate = step.gate || "";
    var gateOk = gateMet(gate);
    var nextLabel = (i===steps.length-1 ? "Done" : "Next");
    var hint = (gate && !gateOk)
      ? '<div style="margin-top:8px;font-size:11px;color:#CC0000;">'
        + (gate==="upload" ? "Upload a file to continue \u2014 I will advance automatically\u2026"
                           : "Click Run analysis to continue \u2014 I will advance automatically\u2026")
        + '</div>'
      : '';
    var card = doc.createElement("div"); card.id="sae-tour-card";
    card.innerHTML =
      '<div id="sae-tour-handle"><span class="grip">\u2630</span>'+
      '<span>Guided tour \u2014 drag to move</span></div>'+
      '<div id="sae-tour-body">'+
      '<h4>'+step.title+'</h4><p>'+step.body+'</p>'+ hint +
      '<div id="sae-tour-bar"><span id="sae-tour-dots">'+(i+1)+' / '+steps.length+'</span>'+
      '<span><button class="sae-tour-btn sae-tour-ghost" id="sae-tour-back">Back</button>'+
      '<button class="sae-tour-btn sae-tour-primary" id="sae-tour-next">'+nextLabel+'</button></span></div>'+
      '<div style="margin-top:10px;"><button class="sae-tour-skip" id="sae-tour-skip">Skip tour</button></div>'+
      '</div>';
    doc.body.appendChild(card);

    // --- positioning ---
    // First appearance: center of screen. After the first Next (or once docked),
    // sit bottom-center. If the user has dragged it, keep their position.
    function placeCard(){
      var w = card.offsetWidth || 460, h = card.offsetHeight || 200;
      var vw = window.parent.innerWidth, vh = window.parent.innerHeight;
      if (window.__saeTourPos){
        card.style.left = window.__saeTourPos.x + "px";
        card.style.top  = window.__saeTourPos.y + "px";
      } else if (!window.__saeTourStarted){
        // before the first Next -> dead center
        card.style.left = Math.max(8,(vw - w)/2) + "px";
        card.style.top  = Math.max(8,(vh - h)/2) + "px";
      } else {
        // docked bottom-center
        card.style.left = Math.max(8,(vw - w)/2) + "px";
        card.style.top  = Math.max(8, vh - h - 28) + "px";
      }
    }
    placeCard();
    setTimeout(placeCard, 30);  // re-place once real height is known

    // --- drag ---
    (function(){
      var handle = doc.getElementById("sae-tour-handle");
      var dragging=false, ox=0, oy=0;
      handle.addEventListener("mousedown", function(e){
        dragging=true;
        var rect=card.getBoundingClientRect();
        ox = e.clientX - rect.left; oy = e.clientY - rect.top;
        e.preventDefault();
      });
      doc.addEventListener("mousemove", function(e){
        if(!dragging) return;
        var x = e.clientX - ox, y = e.clientY - oy;
        x = Math.max(4, Math.min(x, window.parent.innerWidth - card.offsetWidth - 4));
        y = Math.max(4, Math.min(y, window.parent.innerHeight - card.offsetHeight - 4));
        card.style.left = x+"px"; card.style.top = y+"px";
        window.__saeTourPos = {x:x,y:y}; window.__saeTourMoved = true;
      });
      doc.addEventListener("mouseup", function(){ dragging=false; });
    })();

    doc.getElementById("sae-tour-back").onclick=function(){ if(getI()>0){ setI(getI()-1); render(); } };
    doc.getElementById("sae-tour-next").onclick=function(){
      // Always advance — never gets stuck regardless of targets/polls.
      if (window.__saeTourPoll){ clearInterval(window.__saeTourPoll); window.__saeTourPoll=null; }
      window.__saeTourStarted = true;  // dock to bottom-center from now on
      var ci = getI();
      if(ci < steps.length-1){ setI(ci+1); render(); } else { done(); }
    };
    doc.getElementById("sae-tour-skip").onclick=function(){ done(); };

    if (gate && !gateOk){
      if (window.__saeTourPoll){ clearInterval(window.__saeTourPoll); }
      var advanced = false;
      window.__saeTourPoll = setInterval(function(){
        if (advanced) return;
        if (gateMet(gate)){
          advanced = true;
          clearInterval(window.__saeTourPoll); window.__saeTourPoll=null;
          var ci = getI();
          if(ci<steps.length-1){ setI(ci+1); render(); } else { done(); }
        }
      }, 600);
    }
  }
  function gateMet(gate){
    try {
      if (!gate) return true;
      if (gate === "upload"){
        if (doc.querySelector('[data-testid="stFileUploaderFile"], [data-testid="stFileUploaderFileName"]')) return true;
        return (doc.body.innerText||"").indexOf("Uploaded:") !== -1;
      }
      if (gate === "run"){ return !!doc.querySelector(".stat-num"); }
      return true;
    } catch(e){ return true; }
  }
  if (__FORCE__){
    window.__saeTourDone = false;
    window.__saeTourStep = 0;
    window.__saeTourStarted = false;
    window.__saeTourPos = null;
    try { window.parent.sessionStorage.removeItem("sae_tour_seen"); } catch(e){}
    setTimeout(render, 150);
    setTimeout(render, 500);
    return;
  }
  // Non-forced remount: if finished, stay closed; else resume the saved step.
  if (window.__saeTourDone) { return; }
  if (doc.getElementById("sae-tour-card")) { return; }  // already showing
  var seen=null; try { seen=window.parent.sessionStorage.getItem("sae_tour_seen"); } catch(e){}
  if(!seen || typeof window.__saeTourStep === "number" && window.__saeTourStep > 0){
    setTimeout(render, 300);
  }
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
    components_html(_tour_html(TOUR_STEPS, force=forced), height=(1 if forced else 0))


def restart_tour() -> None:
    """Mark the tour to force-open on the next render (handled by maybe_render_tour)."""
    import streamlit as st
    st.session_state[_TOUR_FORCE_KEY] = True
