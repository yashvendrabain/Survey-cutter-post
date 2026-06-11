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
      ".sae-tour-skip{background:none;color:#888;text-decoration:underline;font-size:11px;border:none;cursor:pointer;}"+"#sae-tour-arrow{position:fixed;z-index:2000000;pointer-events:none;font-size:52px;line-height:1;color:#CC0000;filter:drop-shadow(0 2px 4px rgba(0,0,0,0.3));animation:saeArrowB 1s ease-in-out infinite;}"+"@keyframes saeArrowB{0%,100%{opacity:0.5;}50%{opacity:1;}}";
    pdoc.head.appendChild(sty);
  }
})();
</script>
<script>
(function(){
  var doc = window.parent.document;
  var win = window.parent;
  win.__saeSteps = __STEPS__;
  if (typeof win.__saeStep !== "number") win.__saeStep = 0;
  if (__FORCE__){ win.__saeDone=false; win.__saeStep=0; win.__saeStarted=false; win.__saePos=null; }

  win.__saeRemove = function(){
    var c=doc.getElementById("sae-tour-card"); if(c && c.parentNode) c.parentNode.removeChild(c);
    var a=doc.getElementById("sae-tour-arrow"); if(a && a.parentNode) a.parentNode.removeChild(a);
  };

  function findTarget(step){
    if(!step) return null;
    var t = step.selector ? doc.querySelector(step.selector) : null;
    if(!t && step.find_text){
      var bb=doc.querySelectorAll("button");
      for(var i=0;i<bb.length;i++){ if((bb[i].textContent||"").trim().toLowerCase().indexOf(step.find_text.toLowerCase())!==-1){ t=bb[i]; break; } }
    }
    if(!t && step.scroll_text){
      var all=doc.querySelectorAll("h1,h2,h3,h4,p,div,span,label");
      for(var j=0;j<all.length;j++){ if((all[j].textContent||"").toLowerCase().indexOf(step.scroll_text.toLowerCase())!==-1){ t=all[j]; break; } }
    }
    return t;
  }
  function drawArrow(step){
    var a=doc.getElementById("sae-tour-arrow"); if(a&&a.parentNode) a.parentNode.removeChild(a);
    var tgt=findTarget(step); if(!tgt) return;
    var r=tgt.getBoundingClientRect(); if(r.width===0&&r.height===0) return;
    try{ tgt.scrollIntoView({behavior:"smooth",block:"center"}); }catch(e){}
    setTimeout(function(){
      var rr=tgt.getBoundingClientRect();
      var ar=doc.createElement("div"); ar.id="sae-tour-arrow";
      var cx=rr.left+rr.width/2, vh=win.innerHeight, vw=win.innerWidth;
      if(rr.top < vh*0.55){ ar.textContent="\u2B07"; ar.style.left=Math.min(Math.max(8,cx-26),vw-60)+"px"; ar.style.top=Math.max(8,rr.top-58)+"px"; }
      else { ar.textContent="\u2B06"; ar.style.left=Math.min(Math.max(8,cx-26),vw-60)+"px"; ar.style.top=Math.min(vh-66,rr.bottom+10)+"px"; }
      doc.body.appendChild(ar);
    }, 360);
  }

  win.__saeBuild = function(){
    var steps = win.__saeSteps;
    if (win.__saeDone){ win.__saeRemove(); return; }
    var i = win.__saeStep; if(i<0) i=0;
    if (i >= steps.length){ win.__saeDone=true; win.__saeRemove(); return; }
    var step = steps[i];
    win.__saeRemove();
    if (!doc.body) return;
    var last = (i === steps.length-1);
    var card = doc.createElement("div"); card.id="sae-tour-card";
    card.innerHTML =
      '<div id="sae-tour-handle"><span class="grip">&#9776;</span><span>Guided tour &mdash; drag to move</span></div>'+
      '<div id="sae-tour-body">'+
      '<h4>'+step.title+'</h4><p>'+step.body+'</p>'+
      '<div id="sae-tour-bar"><span id="sae-tour-dots">'+(i+1)+' / '+steps.length+'</span>'+
      '<span><button class="sae-tour-btn sae-tour-ghost" data-sae="back">Back</button>'+
      '<button class="sae-tour-btn sae-tour-primary" data-sae="next">'+(last?"Done":"Next")+'</button></span></div>'+
      '<div style="margin-top:10px;"><button class="sae-tour-skip" data-sae="skip">Skip tour</button></div>'+
      '</div>';
    doc.body.appendChild(card);
    var w=card.offsetWidth||440, h=card.offsetHeight||200;
    var vw=win.innerWidth, vh=win.innerHeight;
    if (win.__saePos){ card.style.left=win.__saePos.x+"px"; card.style.top=win.__saePos.y+"px"; }
    else if (!win.__saeStarted){ card.style.left=Math.max(8,(vw-w)/2)+"px"; card.style.top=Math.max(8,(vh-h)/2)+"px"; }
    else { card.style.left=Math.max(8,(vw-w)/2)+"px"; card.style.top=Math.max(8,vh-h-28)+"px"; }
    drawArrow(step);
  };

  // ONE persistent delegated handler. It ALWAYS calls win.__saeBuild (the latest
  // instance), never a stale local closure -> fixes "dead after reopen".
  if (!win.__saeWired){
    win.__saeWired = true;
    doc.addEventListener("click", function(e){
      var t=e.target;
      while (t && t!==doc.body && !(t.getAttribute && t.getAttribute("data-sae"))) t=t.parentNode;
      if (!t || !t.getAttribute) return;
      var act=t.getAttribute("data-sae"); if(!act) return;
      var steps = win.__saeSteps || [];
      if (act==="back"){ if(win.__saeStep>0){ win.__saeStep--; win.__saeBuild(); } }
      else if (act==="next"){
        win.__saeStarted=true;
        if (win.__saeStep < steps.length-1){ win.__saeStep++; win.__saeBuild(); }
        else { win.__saeDone=true; win.__saeRemove(); try{win.sessionStorage.setItem("sae_tour_seen","1");}catch(err){} }
      } else if (act==="skip"){
        win.__saeDone=true; win.__saeRemove(); try{win.sessionStorage.setItem("sae_tour_seen","1");}catch(err){}
      }
    }, true);
    var drag=false, ox=0, oy=0;
    doc.addEventListener("mousedown", function(e){
      var t=e.target; while(t && t!==doc.body && t.id!=="sae-tour-handle") t=t.parentNode;
      if(!t || t.id!=="sae-tour-handle") return;
      var card=doc.getElementById("sae-tour-card"); if(!card) return;
      drag=true; var r=card.getBoundingClientRect(); ox=e.clientX-r.left; oy=e.clientY-r.top; e.preventDefault();
    }, true);
    doc.addEventListener("mousemove", function(e){
      if(!drag) return; var card=doc.getElementById("sae-tour-card"); if(!card) return;
      var x=Math.max(4,Math.min(e.clientX-ox, win.innerWidth-card.offsetWidth-4));
      var y=Math.max(4,Math.min(e.clientY-oy, win.innerHeight-card.offsetHeight-4));
      card.style.left=x+"px"; card.style.top=y+"px"; win.__saePos={x:x,y:y};
    }, true);
    doc.addEventListener("mouseup", function(){ drag=false; }, true);
  }

  if (__FORCE__){ setTimeout(win.__saeBuild,120); setTimeout(win.__saeBuild,450); return; }
  if (win.__saeDone) return;
  if (doc.getElementById("sae-tour-card")) return;
  var seen=null; try{ seen=win.sessionStorage.getItem("sae_tour_seen"); }catch(e){}
  if(!seen || win.__saeStep>0){ setTimeout(win.__saeBuild,300); }
})();
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
