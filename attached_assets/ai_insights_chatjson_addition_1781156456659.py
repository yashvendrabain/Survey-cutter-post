

# =====================================================================
# APPEND-ONLY BLOCK — shared JSON chat helper for the in-app assistant.
# Reuses this module's existing imports/constants: os, json, OpenAI,
# PORTKEY_BASE_URL, PORTKEY_DEFAULT_MODEL. Keeps ai_insights.py as the
# single place that talks to Portkey. Never raises — returns None on any
# failure so callers (assistant_bot._llm_json) degrade to deterministic mode.
# =====================================================================

def _portkey_chat_json(system: str, user: str, *, temperature: float = 0.0):
    """Call Portkey/OpenAI and return a parsed JSON object, or None.

    Mirrors the existing call sites in this module (model=PORTKEY_DEFAULT_MODEL,
    OpenAI(api_key, base_url=PORTKEY_BASE_URL)). Strips ```json fences before
    parsing. Returns a dict on success; None if the key is missing, the SDK is
    unavailable, the call fails, or the response isn't a JSON object.
    """
    import json as _json
    import os as _os
    import re as _re

    api_key = _os.environ.get("PORTKEY_API_KEY", "").strip()
    if not api_key or OpenAI is None:
        return None
    try:
        client = OpenAI(api_key=api_key, base_url=PORTKEY_BASE_URL)
        response = client.chat.completions.create(
            model=PORTKEY_DEFAULT_MODEL,
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
            temperature=float(temperature),
            max_tokens=1200,
            timeout=25,
        )
        raw = response.choices[0].message.content or ""
    except Exception:
        return None

    cleaned = _re.sub(r"^```(?:json)?|```$", "", raw.strip(), flags=_re.MULTILINE).strip()
    try:
        obj = _json.loads(cleaned)
        return obj if isinstance(obj, dict) else None
    except Exception:
        return None
