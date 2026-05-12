"""AI insight generation through Portkey/OpenAI.

This module is the only place in the project that knows how to call Portkey.
All callers pass already-computed tables and receive plain InsightResult data.
"""

from __future__ import annotations

import hashlib
import json
import os
from typing import Any, Callable

from config import (
    AI_INSIGHT_MAX_TOKENS,
    AI_INSIGHT_TEMPERATURE,
    AI_INSIGHT_TIMEOUT_SECONDS,
    PORTKEY_BASE_URL,
    PORTKEY_DEFAULT_MODEL,
    PORTKEY_PREMIUM_MODEL,
)
from src.models import InsightResult

try:
    from openai import OpenAI
except ModuleNotFoundError:
    OpenAI = None  # type: ignore[assignment]


SYSTEM_PROMPT = """You are a senior strategy consultant writing slide
headlines for a C-suite presentation based on survey data.

Your job: write ONE headline that could appear at the top of a
PowerPoint slide. The headline must be a direct, confident takeaway -
not a description of the data.

HEADLINE RULES (enforce strictly):
1. Exactly ONE sentence. 10-25 words. No more.
2. Must contain at least one specific number from the table
   (%, x multiple, or count). Never write a headline without a number.
3. Start with the subject (Winners / Laggards / Companies /
   Respondents / [Industry name] - whatever is most specific).
4. Use a strong active verb: deliver, drive, invest, outpace,
   prioritize, focus, reveal, show, achieve, report, demonstrate.
5. If comparing two groups, name both:
   "Winners X% vs Laggards Y%" or "Winners deliver Nx vs average".
6. No hedging words: never use "suggest", "appear", "may", "could",
   "tend to", "seem". State it as fact.
7. No meta-commentary: never say "the data shows", "survey reveals",
   "respondents indicate". Just state the finding.
8. Non-technical language: a CFO with no stats background must
   understand it instantly. No "Cramer's V", no "statistically
   significant", no "lift ratio".
9. If the winner rate and laggard rate are both provided, prefer a
   comparative headline over a single-group headline.
10. Capitalise only the first word and proper nouns.

GOOD EXAMPLES (study these patterns):
- "Winners invest 2.3x more in GTM technology than Laggards"
- "Despite headwinds, Winners delivered 2x sector-average revenue
  growth in 2023"
- "67% of Winners prioritize inside sales vs. 34% of Laggards"
- "Winners expect 3.7x higher revenue growth in 2024"
- "Machinery sector outpaces all other industries with 2.6x winner
  growth multiple"
- "Winners are driving better returns from GenAI while Laggards
  struggle to meet expectations"

BAD EXAMPLES (never produce these):
- "The data suggests that winners may tend to invest more"
  [hedging + no number]
- "There is a statistically significant difference between groups"
  [technical jargon]
- "Survey respondents indicate varying levels of satisfaction"
  [meta-commentary + no number]
- "Winners and Laggards differ on this dimension"
  [no number + vague]

ABSOLUTE RULES - NEVER VIOLATE:
- NEVER invent a number. Every number must come from the JSON table.
- NEVER change a number. 67.3% stays 67.3% (or 67% if rounding).
- NEVER imply causality unless the table explicitly shows it.
- NEVER write more than one sentence.
- If no clear pattern exists in the data, write:
  "No strong differentiation visible on this dimension"

INPUT FORMAT:
You will receive a JSON object with the computed table. Use only the
numbers in that JSON. Do not use any other numbers.
"""

AI_SYSTEM_PROMPT = SYSTEM_PROMPT

USER_PROMPT_TEMPLATE = """Here is the computed survey data table:

{table_json}

Write one PPT-ready slide headline following the system prompt rules.
Return ONLY the headline sentence. No preamble, no explanation,
no punctuation at the end except if the sentence naturally requires it.
"""

TEMPLATE_INSIGHT = "No AI insight available — review the data table above."


def generate_insight(
    table_payload: dict,
    table_kind: str = "generic",
    title_hint: str = "",
    cache: dict | None = None,
    use_premium: bool = False,
) -> InsightResult:
    """Generate one PPT-ready headline for a computed table."""

    if cache is None:
        return _call_api(
            table_payload,
            table_kind=table_kind,
            title_hint=title_hint,
            use_premium=use_premium,
        )

    key = _payload_hash(table_kind, table_payload)
    if key in cache:
        return cache[key]

    result = _call_api(
        table_payload,
        table_kind=table_kind,
        title_hint=title_hint,
        use_premium=use_premium,
    )
    cache[key] = result
    return result


def _call_api(
    table_payload: dict,
    *,
    table_kind: str,
    title_hint: str,
    use_premium: bool,
) -> InsightResult:
    api_key = os.environ.get("PORTKEY_API_KEY", "").strip()
    if not api_key:
        return _template_fallback(
            table_payload,
            table_kind,
            title_hint,
            error_message="PORTKEY_API_KEY not set",
        )

    if OpenAI is None:
        return _template_fallback(
            table_payload,
            table_kind,
            title_hint,
            error_message="openai package not installed",
        )

    model = PORTKEY_PREMIUM_MODEL if use_premium else PORTKEY_DEFAULT_MODEL
    user_prompt = USER_PROMPT_TEMPLATE.format(
        table_json=_format_payload(table_kind, table_payload)
    )

    try:
        client = OpenAI(api_key=api_key, base_url=PORTKEY_BASE_URL)
        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_prompt},
            ],
            temperature=AI_INSIGHT_TEMPERATURE,
            max_tokens=AI_INSIGHT_MAX_TOKENS,
            timeout=AI_INSIGHT_TIMEOUT_SECONDS,
        )
    except Exception as exc:
        return _template_fallback(
            table_payload,
            table_kind,
            title_hint,
            error_message=f"API call failed: {type(exc).__name__}: {exc}",
        )

    raw = response.choices[0].message.content
    try:
        headline = _extract_headline(raw)
        if not headline:
            raise ValueError("headline empty in response")
    except (TypeError, json.JSONDecodeError, ValueError) as exc:
        return _template_fallback(
            table_payload,
            table_kind,
            title_hint,
            error_message=f"Response parse failed: {exc}",
        )

    return InsightResult(
        title="",
        insight=headline,
        was_template=False,
        model_used=model,
        tokens_used=int(getattr(response.usage, "total_tokens", 0) or 0),
        error_message="",
    )


def _extract_headline(raw: Any) -> str:
    if raw is None:
        return ""

    text = str(raw).strip()
    if not text:
        return ""

    if text.startswith("{"):
        parsed = json.loads(text)
        headline = parsed.get("headline") or parsed.get("insight") or parsed.get("title")
        return str(headline or "").strip()

    return text.strip().strip('"')


def _format_differentiator_payload(payload: dict) -> str:
    return json.dumps(
        {
            "analysis_type": "winner_vs_laggard_differentiator",
            "question": payload.get("question_text", ""),
            "top_differentiating_option": payload.get("top_option", ""),
            "winner_selection_rate": f"{payload.get('winner_rate', 0):.1%}",
            "laggard_selection_rate": f"{payload.get('loser_rate', 0):.1%}",
            "lift_ratio": f"{payload.get('lift', 1):.2f}x",
            "association_strength": f"Cramers V = {payload.get('cramers_v', 0):.3f}",
            "winner_n": payload.get("winner_n", 0),
            "laggard_n": payload.get("loser_n", 0),
        },
        indent=2,
    )


def _format_winner_profile_payload(payload: dict) -> str:
    traits = payload.get("traits", [])
    return json.dumps(
        {
            "analysis_type": "winner_profile_summary",
            "winner_label": payload.get("winner_label", "Winner"),
            "laggard_label": payload.get("loser_label", "Laggard"),
            "winner_n": payload.get("winner_n", 0),
            "laggard_n": payload.get("loser_n", 0),
            "defining_traits": [
                {
                    "question": t.get("question_id", ""),
                    "winning_option": t.get("option_label", ""),
                    "winner_rate": f"{t.get('winner_rate', 0):.1%}",
                    "laggard_rate": f"{t.get('loser_rate', 0):.1%}",
                    "lift": f"{t.get('lift', 1):.2f}x",
                    "rate_gap": f"+{t.get('rate_gap', 0):.1%}",
                }
                for t in traits[:5]
            ],
        },
        indent=2,
    )


def _format_single_cut_payload(payload: dict) -> str:
    return json.dumps(
        {
            "analysis_type": "single_question_distribution",
            "question": payload.get("question_text", ""),
            "question_type": payload.get("question_type", ""),
            "total_n": payload.get("valid_n", 0),
            "distribution": payload.get("distribution", {}),
            "top_option": payload.get("top_option", ""),
            "top_option_pct": payload.get("top_option_pct", ""),
        },
        indent=2,
    )


def _format_filtered_single_cut_payload(payload: dict) -> str:
    return json.dumps(
        {
            "analysis_type": "filtered_single_question_distribution",
            "question": payload.get("question_text", ""),
            "filter_applied": payload.get("filter_description", ""),
            "filtered_n": payload.get("filtered_n", 0),
            "total_n": payload.get("total_n", 0),
            "distribution": payload.get("distribution", {}),
            "top_option": payload.get("top_option", ""),
            "top_option_pct": payload.get("top_option_pct", ""),
        },
        indent=2,
    )


def _format_cross_cut_payload(payload: dict) -> str:
    return json.dumps(
        {
            "analysis_type": "cross_cut_comparison",
            "title": payload.get("title", ""),
            "source_questions": payload.get("source_questions", []),
            "analysis_type_detail": payload.get("analysis_type_detail", ""),
            "table": payload.get("table", {}),
            "key_finding": payload.get("key_finding", ""),
        },
        indent=2,
    )


FORMATTERS: dict[str, Callable[[dict], str]] = {
    "differentiator": _format_differentiator_payload,
    "winner_profile": _format_winner_profile_payload,
    "single_cut": _format_single_cut_payload,
    "filtered_single_cut": _format_filtered_single_cut_payload,
    "cross_cut": _format_cross_cut_payload,
}


def _format_payload(table_kind: str, payload: dict) -> str:
    formatter = FORMATTERS.get(table_kind, lambda p: json.dumps(p, indent=2, default=str))
    return formatter(payload)


def _payload_hash(table_kind: str, payload: dict) -> str:
    raw = json.dumps({"kind": table_kind, "payload": payload}, sort_keys=True, default=str)
    return hashlib.md5(raw.encode()).hexdigest()


def _template_fallback(
    payload: dict,
    table_kind: str,
    title_hint: str,
    error_message: str,
) -> InsightResult:
    return InsightResult(
        title="",
        insight=TEMPLATE_INSIGHT,
        was_template=True,
        model_used="(template fallback)",
        tokens_used=0,
        error_message=error_message,
    )
