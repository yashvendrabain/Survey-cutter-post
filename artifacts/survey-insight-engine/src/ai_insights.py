"""AI insight generation through Portkey/OpenAI.

This module is the only place in the project that knows how to call Portkey.
All callers pass already-computed tables and receive plain InsightResult data.
"""

from __future__ import annotations

import json
import os
from typing import Any

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


AI_SYSTEM_PROMPT = """You are a survey-data analyst. You read computed tables and write concise observations.
Strict rules:

1. NEVER invent numbers. Every number in your output must appear in the input table or be a direct comparison of two numbers in the table (e.g. "twice as likely" when one rate is 2x another).
2. NEVER speculate about respondents, motivations, or causes. Only describe what the data shows.
3. Use the question text and filter context to write natural prose, not bullet points.
4. If the data shows no clear pattern (rates within 5 percentage points of each other across categories), say so explicitly rather than fabricating a trend.
5. Output STRICTLY as JSON:
   {"title": "<5-10 word title>", "insight": "<2-3 sentence observation>"}
   No surrounding prose, no markdown, no explanations outside the JSON.
"""


def generate_insight(
    table_payload: dict,
    *,
    table_kind: str,
    title_hint: str = "",
    use_premium: bool = False,
) -> InsightResult:
    """Generate a title and insight for a computed table."""

    rows = table_payload.get("rows", [])
    if not rows:
        return _template_fallback(
            table_payload,
            table_kind,
            title_hint,
            error_message="No rows supplied",
        )

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
    user_prompt = _build_user_prompt(table_payload, table_kind)

    try:
        client = OpenAI(api_key=api_key, base_url=PORTKEY_BASE_URL)
        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": AI_SYSTEM_PROMPT},
                {"role": "user", "content": user_prompt},
            ],
            temperature=AI_INSIGHT_TEMPERATURE,
            max_tokens=AI_INSIGHT_MAX_TOKENS,
            timeout=AI_INSIGHT_TIMEOUT_SECONDS,
            response_format={"type": "json_object"},
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
        parsed = json.loads(raw)
        title = str(parsed.get("title", "")).strip()
        insight = str(parsed.get("insight", "")).strip()
        if not title or not insight:
            raise ValueError("title or insight empty in response")
    except (TypeError, json.JSONDecodeError, ValueError) as exc:
        return _template_fallback(
            table_payload,
            table_kind,
            title_hint,
            error_message=f"Response parse failed: {exc}",
        )

    return InsightResult(
        title=title,
        insight=insight,
        was_template=False,
        model_used=model,
        tokens_used=int(getattr(response.usage, "total_tokens", 0) or 0),
        error_message="",
    )


def _build_user_prompt(payload: dict, table_kind: str) -> str:
    question_text = payload.get("question_text", "")
    valid_n = payload.get("valid_n", 0)
    missing_n = payload.get("missing_n", 0)
    filters = payload.get("filters_applied", [])
    rows = payload.get("rows", [])
    summary = payload.get("summary", {})

    filters_str = ", ".join(str(filter_item) for filter_item in filters)
    if not filters_str:
        filters_str = "(none)"

    rows_json = json.dumps(rows, indent=2, sort_keys=True, default=str)
    summary_text = (
        json.dumps(summary, indent=2, sort_keys=True, default=str)
        if summary
        else "(none)"
    )

    return f"""Survey question: {question_text}

Table type: {table_kind}
Valid responses: N = {valid_n}
Missing responses: {missing_n}
Filters applied: {filters_str}

Computed table:
{rows_json}

Summary statistics: {summary_text}

Generate a JSON object with a "title" (5-10 words, sentence case, describes what the table shows) and an "insight" (2-3 sentences, observation grounded in the numbers above, no speculation)."""


def _template_fallback(
    payload: dict,
    table_kind: str,
    title_hint: str,
    error_message: str,
) -> InsightResult:
    qtext = str(payload.get("question_text", ""))
    valid_n = int(payload.get("valid_n", 0) or 0)
    rows = payload.get("rows", [])
    filters = payload.get("filters_applied", [])

    source = title_hint or qtext or "Survey result"
    words = source.split()
    title = " ".join(words[:8])
    if len(words) > 8:
        title += "..."
    title = title[0].upper() + title[1:] if title else "Result"

    if not rows:
        insight = f"No respondents matched the criteria (N = {valid_n})."
    elif table_kind in ("single_cut", "filtered_single_cut"):
        top = max(rows, key=lambda row: row.get("count", 0))
        label = top.get("label", "the top option")
        count = int(top.get("count", 0) or 0)
        rate = float(top.get("rate", 0) or 0)
        insight = f"Of {valid_n} respondents, {count} ({rate:.0%}) chose {label}."
        if filters:
            insight += f" Filters applied: {', '.join(str(f) for f in filters)}."
    else:
        insight = (
            f"Table contains {len(rows)} rows across {valid_n} valid respondents. "
            "AI insight unavailable; please review the data manually."
        )

    return InsightResult(
        title=title,
        insight=insight,
        was_template=True,
        model_used="(template fallback)",
        tokens_used=0,
        error_message=error_message,
    )
