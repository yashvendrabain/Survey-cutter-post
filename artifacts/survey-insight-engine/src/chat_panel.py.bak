"""Survey Analysis Engine — chat panel UI (frontend for assistant_bot).

Self-contained so app.py only needs:  from src.chat_panel import render_chat_panel
then a single call `render_chat_panel()` near the END of main() (so it overlays
all views).

This module owns ONLY presentation + session wiring. All answers come from
assistant_bot.handle_message, which is calculation-first and validated. This
file never computes a number.

UI design:
  - A sticky launcher button (bottom-right) toggles an in-page chat panel.
    The launcher is drawn with components.v1 so it can float over Streamlit.
  - The actual conversation uses native Streamlit chat widgets (st.chat_input,
    st.chat_message) rendered inside an expander-like container, because native
    widgets are the only reliable way to round-trip user text into Python.
  - Hypothesis answers render the returned table with st.dataframe and show the
    caveats + an "AI-phrased / not grounded" badge when was_grounded is False.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

_HISTORY_KEY = "chat_history"          # list[dict]: {role, text, table, caption, caveats, grounded}
_OPEN_KEY = "chat_panel_open"
_FAQ_PATH_DEFAULT = "assistant_faq.json"


def _load_faq_ground_truth(faq_path: str = _FAQ_PATH_DEFAULT) -> str:
    """Flatten the FAQ JSON into a compact reference string for the bot."""
    p = Path(faq_path)
    if not p.exists():
        # Try a couple of common locations.
        for alt in ("src/assistant_faq.json", "/tmp/assistant_faq.json"):
            if Path(alt).exists():
                p = Path(alt)
                break
        else:
            return ""
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        return "\n".join(f"Q: {f['q']}\nA: {f['a']}" for f in data.get("faqs", []))
    except Exception:
        return ""


def render_chat_panel(faq_path: str = _FAQ_PATH_DEFAULT) -> None:
    """Native Streamlit assistant panel — NO injected JavaScript.

    A toggle opens an in-page panel with three tabs (Tool questions, Survey,
    Hypothesis validator) and the chat input lives inside the panel itself.
    """
    try:
        import streamlit as st
    except Exception:
        return

    st.session_state.setdefault(_HISTORY_KEY, [])
    st.session_state.setdefault(_OPEN_KEY, False)
    st.session_state.setdefault("chat_mode", "tool")

    # Native open/close toggle: red, with an SVG chat icon (no emoji). Scoped so
    # it works wherever the panel is rendered (we put it in the sidebar).
    _chat_icon_svg = (
        "data:image/svg+xml;utf8,"
        "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='16' viewBox='0 0 24 24' "
        "fill='none' stroke='white' stroke-width='2.2' stroke-linecap='round' stroke-linejoin='round'>"
        "<path d='M21 11.5a8.38 8.38 0 0 1-8.5 8.5 8.5 8.5 0 0 1-3.6-.8L3 21l1.9-5.4A8.38 8.38 0 0 1 4 "
        "11.5 8.5 8.5 0 0 1 12.5 3 8.38 8.38 0 0 1 21 11.5z'/></svg>"
    )
    st.markdown(
        "<style>#sae-chat-toggle-anchor + div .stButton button{"
        "background:#CC0000 !important;color:#fff !important;border:none !important;"
        "font-weight:700 !important;border-radius:10px !important;}"
        "#sae-chat-toggle-anchor + div .stButton button:hover{background:#B30000 !important;}"
        "#sae-chat-toggle-anchor + div .stButton button::before{content:'';display:inline-block;"
        f"width:16px;height:16px;margin-right:7px;vertical-align:-2px;background:url(\"{_chat_icon_svg}\") no-repeat center/contain;}}"
        "</style>"
        "<div id='sae-chat-toggle-anchor'></div>",
        unsafe_allow_html=True,
    )
    label = "Close assistant" if st.session_state[_OPEN_KEY] else "Have a question? Ask away"
    if st.button(label, key="chat_open_toggle", use_container_width=True):
        st.session_state[_OPEN_KEY] = not st.session_state[_OPEN_KEY]
        st.rerun()

    if not st.session_state[_OPEN_KEY]:
        return

    # Chat-card styling: make the native container read as a real chat widget —
    # rounded, shadowed, fixed-ish width, gradient header, bubble-style messages.
    st.markdown(
        """
        <style>
        #sae-chat-card-anchor + div[data-testid="stVerticalBlockBorderWrapper"]{
          max-width:440px; border:1px solid #E3E3E8 !important; border-radius:18px !important;
          box-shadow:0 18px 50px rgba(0,0,0,0.18) !important; overflow:hidden !important;
          background:#fff !important; padding:0 !important;
        }
        #sae-chat-card-anchor + div[data-testid="stVerticalBlockBorderWrapper"] > div{
          padding:0 16px 12px !important;
        }
        .sae-chat-head{
          margin:0 -16px 8px; padding:16px 18px 14px;
          background:linear-gradient(135deg,#CC0000,#8E0000); color:#fff;
          display:flex; align-items:center; gap:12px;
        }
        .sae-chat-head .av{
          width:40px;height:40px;border-radius:50%;background:#fff;color:#CC0000;
          display:flex;align-items:center;justify-content:center;font-weight:800;font-size:20px;
        }
        .sae-chat-head .t1{font-size:17px;font-weight:700;line-height:1.1;}
        .sae-chat-head .t2{font-size:12px;opacity:0.9;margin-top:2px;}
        /* tab pills */
        #sae-chat-card-anchor + div [data-testid="stTabs"] [data-baseweb="tab-list"]{
          gap:6px; background:#F6F6F8; padding:5px; border-radius:10px;
        }
        #sae-chat-card-anchor + div [data-testid="stTabs"] [data-baseweb="tab"]{
          border-radius:8px; font-size:12px; font-weight:600; padding:6px 10px;
        }
        #sae-chat-card-anchor + div [data-testid="stTabs"] [aria-selected="true"]{
          background:#CC0000; color:#fff;
        }
        </style>
        <div id="sae-chat-card-anchor"></div>
        """,
        unsafe_allow_html=True,
    )

    with st.container(border=True):
        st.markdown(
            "<div class='sae-chat-head'>"
            "<div class='av'>a</div>"
            "<div><div class='t1'>Survey Assistant</div>"
            "<div class='t2'>How can we help you today?</div></div></div>",
            unsafe_allow_html=True,
        )

        tab_tool, tab_survey, tab_hyp = st.tabs(
            ["Tool questions", "Survey questions", "Hypothesis validator"]
        )
        mode_prompts = {
            "tool": "Ask how to use the tool…",
            "survey": "Ask what's in your survey…",
            "hyp": "State a hypothesis to test against the data…",
        }
        mode_intros = {
            "tool": "Ask about features, exports, filters, or how any part of the tool works.",
            "survey": "Ask what questions or segments exist in your loaded survey.",
            "hyp": "Propose a relationship (e.g. 'larger firms grow faster') and it'll be tested against your data, with the table shown.",
        }

        def _render_mode(mode_key: str, intro: str):
            st.caption(intro)
            # history for this mode
            for turn in st.session_state[_HISTORY_KEY]:
                if turn.get("mode") != mode_key:
                    continue
                with st.chat_message(turn["role"]):
                    st.write(turn["text"])
                    if turn.get("caveats"):
                        for c in turn["caveats"]:
                            st.warning(c)
                    if turn.get("table") is not None:
                        if turn.get("caption"):
                            st.caption(turn["caption"])
                        st.dataframe(turn["table"], use_container_width=True, hide_index=True)
                    if turn["role"] == "assistant" and not turn.get("grounded", True):
                        st.caption("\u26a0\ufe0f AI-phrased; confirm against the table/guide.")
            prompt = st.chat_input(mode_prompts[mode_key], key=f"chat_input_{mode_key}")
            if prompt:
                st.session_state[_HISTORY_KEY].append(
                    {"role": "user", "text": prompt, "mode": mode_key}
                )
                reply = _dispatch(prompt, faq_path)
                st.session_state[_HISTORY_KEY].append({
                    "role": "assistant", "text": reply.text, "table": reply.table,
                    "caption": reply.table_caption, "caveats": reply.caveats,
                    "grounded": reply.was_grounded, "mode": mode_key,
                })
                st.rerun()

        with tab_tool:
            _render_mode("tool", mode_intros["tool"])
        with tab_survey:
            _render_mode("survey", mode_intros["survey"])
        with tab_hyp:
            _render_mode("hyp", mode_intros["hyp"])


def _dispatch(prompt: str, faq_path: str):
    """Bridge UI -> assistant_bot, pulling live state from session."""
    import streamlit as st
    try:
        from src.assistant_bot import handle_message
    except Exception:
        try:
            from assistant_bot import handle_message  # flat layout fallback
        except Exception:
            from types import SimpleNamespace
            return SimpleNamespace(
                text="The assistant module isn't installed yet.",
                table=None, table_caption="", caveats=[], was_grounded=False,
            )
    return handle_message(
        prompt,
        schema=st.session_state.get("schema"),
        active_df=st.session_state.get("active_df"),
        log=st.session_state.get("log"),
        faq_ground_truth=_load_faq_ground_truth(faq_path),
    )
