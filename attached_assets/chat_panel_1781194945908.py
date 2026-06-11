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


def _launcher_html(is_open: bool) -> str:
    # Professional inline-SVG icons (no emoji). Chat bubble when closed, X when open.
    chat_svg = (
        "<svg width='24' height='24' viewBox='0 0 24 24' fill='none' "
        "xmlns='http://www.w3.org/2000/svg'><path d='M21 11.5a8.38 8.38 0 0 1-8.5 8.5 "
        "8.5 8.5 0 0 1-3.6-.8L3 21l1.9-5.4A8.38 8.38 0 0 1 4 11.5 8.5 8.5 0 0 1 12.5 3 "
        "8.38 8.38 0 0 1 21 11.5z' stroke='white' stroke-width='2' stroke-linecap='round' "
        "stroke-linejoin='round'/></svg>"
    )
    close_svg = (
        "<svg width='22' height='22' viewBox='0 0 24 24' fill='none' "
        "xmlns='http://www.w3.org/2000/svg'><path d='M18 6 6 18M6 6l12 12' stroke='white' "
        "stroke-width='2.2' stroke-linecap='round'/></svg>"
    )
    icon = close_svg if is_open else chat_svg
    show_prompt = "false" if is_open else "true"
    return """
<script>
(function(){
  var doc = window.parent.document;
  function ensure(){
    if(!doc.body){ return; }
    var STYLE_ID = "sae-chat-style";
    if(!doc.getElementById(STYLE_ID)){
      var s = doc.createElement("style"); s.id = STYLE_ID;
      s.textContent =
        "#sae-chat-launch{position:fixed;right:22px;bottom:22px;width:58px;height:58px;"+
        "border-radius:50%;background:#CC0000;border:none;"+
        "cursor:pointer;box-shadow:0 8px 22px rgba(204,0,0,0.4);z-index:1500000;"+
        "display:flex;align-items:center;justify-content:center;transition:background .15s;}"+
        "#sae-chat-launch:hover{background:#B30000;}"+
        "#sae-chat-prompt{position:fixed;right:92px;bottom:34px;z-index:1500000;"+
        "background:#0A0A0A;color:#fff;font-family:Arial,sans-serif;font-size:12.5px;"+
        "font-weight:600;padding:9px 13px;border-radius:10px;box-shadow:0 6px 18px rgba(0,0,0,0.25);"+
        "max-width:190px;line-height:1.4;}"+
        "#sae-chat-prompt:after{content:'';position:absolute;right:-6px;bottom:14px;"+
        "border:6px solid transparent;border-left-color:#0A0A0A;}";
      doc.head.appendChild(s);
    }
    var btn = doc.getElementById("sae-chat-launch");
    if(!btn){
      btn = doc.createElement("button");
      btn.id = "sae-chat-launch";
      btn.title = "Assistant";
      doc.body.appendChild(btn);
    }
    btn.innerHTML = "__ICON__";
    // Prompt bubble ("Have a question? Ask it") — only when closed.
    var prompt = doc.getElementById("sae-chat-prompt");
    if (__SHOW_PROMPT__){
      if(!prompt){
        prompt = doc.createElement("div"); prompt.id="sae-chat-prompt";
        prompt.textContent = "Have a question? Ask it \u2014 I can help.";
        doc.body.appendChild(prompt);
      }
    } else if (prompt){ prompt.remove(); }
    if(!btn.dataset.wired){
      btn.dataset.wired = "1";
      btn.addEventListener("click", function(){
        var all = doc.querySelectorAll('button');
        for(var i=0;i<all.length;i++){
          if((all[i].textContent||'').trim().indexOf('chat_toggle_signal')===0){ all[i].click(); return; }
        }
      });
    }
  }
  ensure();
  setTimeout(ensure, 200);
  setTimeout(ensure, 800);
  setInterval(ensure, 1500);
})();
</script>
""".replace("__ICON__", icon).replace("__SHOW_PROMPT__", show_prompt)


def render_chat_panel(faq_path: str = _FAQ_PATH_DEFAULT) -> None:
    """Render the sticky launcher + (when open) the chat conversation panel."""
    try:
        import streamlit as st
        from streamlit.components.v1 import html as components_html
    except Exception:
        return

    st.session_state.setdefault(_HISTORY_KEY, [])
    st.session_state.setdefault(_OPEN_KEY, False)

    # Hidden toggle the launcher's JS clicks. Off-screen (not display:none) so it
    # stays clickable. The nav-bar JS sweep also hides any button labelled
    # 'chat_toggle_signal'; this is a belt-and-suspenders inline hide.
    if st.button("chat_toggle_signal", key="chat_toggle_signal"):
        st.session_state[_OPEN_KEY] = not st.session_state[_OPEN_KEY]
        st.rerun()

    # height>=1 so the injected <script> reliably executes in the component iframe.
    components_html(_launcher_html(st.session_state[_OPEN_KEY]), height=1)

    if not st.session_state[_OPEN_KEY]:
        return

    with st.container():
        st.markdown(
            "<div style='border:1px solid #E0E0E0;border-top:4px solid #CC0000;"
            "border-radius:12px;padding:14px 16px;margin:8px 0;background:#fff;'>"
            "<b style='font-family:Arial;'>Assistant</b>"
            "<span style='font-size:11px;color:#888;font-family:Arial;'> · how-to · "
            "your survey · hypothesis testing</span></div>",
            unsafe_allow_html=True,
        )

        # Replay history.
        for turn in st.session_state[_HISTORY_KEY]:
            with st.chat_message(turn["role"]):
                st.write(turn["text"])
                if turn.get("caveats"):
                    for c in turn["caveats"]:
                        st.warning(c)
                if turn.get("table"):
                    if turn.get("caption"):
                        st.caption(turn["caption"])
                    st.dataframe(turn["table"], use_container_width=True, hide_index=True)
                if turn["role"] == "assistant" and not turn.get("grounded", True):
                    st.caption("\u26a0\ufe0f AI-phrased; not fully grounded — read the table/guide to confirm.")

        prompt = st.chat_input("Ask about the tool, your survey, or test a hypothesis…")
        if prompt:
            st.session_state[_HISTORY_KEY].append({"role": "user", "text": prompt})
            reply = _dispatch(prompt, faq_path)
            st.session_state[_HISTORY_KEY].append({
                "role": "assistant",
                "text": reply.text,
                "table": reply.table,
                "caption": reply.table_caption,
                "caveats": reply.caveats,
                "grounded": reply.was_grounded,
            })
            st.rerun()


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
