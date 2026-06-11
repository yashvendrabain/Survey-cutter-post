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


_TOUR_STEP_KEY = "product_tour_step"


def _tour_steps():
    return TOUR_STEPS


def maybe_render_tour(force: bool = False) -> None:
    """Native Streamlit tour — NO injected JavaScript.

    Renders an in-page guided card at the top of the app using real Streamlit
    widgets, so it cannot fail the way the previous injected-JS tour did. Shows
    automatically on first visit (until skipped/finished), and on demand via
    restart_tour().
    """
    try:
        import streamlit as st
    except Exception:
        return

    forced = bool(force) or bool(st.session_state.get(_TOUR_FORCE_KEY))
    if forced:
        st.session_state[_TOUR_SEEN_KEY] = False
        st.session_state[_TOUR_STEP_KEY] = 0
        st.session_state[_TOUR_FORCE_KEY] = False

    # Don't show if already seen/skipped this session and not forced.
    if st.session_state.get(_TOUR_SEEN_KEY):
        return

    steps = _tour_steps()
    i = int(st.session_state.get(_TOUR_STEP_KEY, 0) or 0)
    if i < 0:
        i = 0
    if i >= len(steps):
        st.session_state[_TOUR_SEEN_KEY] = True
        return
    step = steps[i]

    # Render the tour card as a bordered container at the top.
    with st.container(border=True):
        st.markdown(
            f"<div style='display:flex;align-items:center;justify-content:space-between;'>"
            f"<div style='font-size:11px;font-weight:700;letter-spacing:0.08em;"
            f"text-transform:uppercase;color:#CC0000;'>Guided tour \u00b7 step {i+1} of {len(steps)}</div>"
            f"</div>"
            f"<div style='font-size:17px;font-weight:700;color:#0A0A0A;margin-top:4px;'>"
            f"{step['title']}</div>"
            f"<div style='font-size:13px;color:#444;line-height:1.55;margin:8px 0 4px;'>"
            f"{step['body']}</div>",
            unsafe_allow_html=True,
        )
        c_back, c_next, c_spacer, c_skip = st.columns([1, 1, 4, 1])
        with c_back:
            if st.button("\u2039 Back", key=f"tour_back_{i}", disabled=(i == 0),
                         use_container_width=True):
                st.session_state[_TOUR_STEP_KEY] = max(0, i - 1)
                st.rerun()
        with c_next:
            last = (i == len(steps) - 1)
            if st.button("Done" if last else "Next \u203a", key=f"tour_next_{i}",
                         type="primary", use_container_width=True):
                if last:
                    st.session_state[_TOUR_SEEN_KEY] = True
                    st.session_state[_TOUR_STEP_KEY] = 0
                else:
                    st.session_state[_TOUR_STEP_KEY] = i + 1
                st.rerun()
        with c_skip:
            if st.button("Skip", key=f"tour_skip_{i}", use_container_width=True):
                st.session_state[_TOUR_SEEN_KEY] = True
                st.session_state[_TOUR_STEP_KEY] = 0
                st.rerun()


def restart_tour() -> None:
    """Replay the tour from the start (used by the sidebar button)."""
    import streamlit as st
    st.session_state[_TOUR_FORCE_KEY] = True
    st.session_state[_TOUR_SEEN_KEY] = False
    st.session_state[_TOUR_STEP_KEY] = 0
