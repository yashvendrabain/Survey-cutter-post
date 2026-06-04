"""AI insight generation through Portkey/OpenAI.

This module is the only place in the project that knows how to call Portkey.
All callers pass already-computed tables and receive plain InsightResult data.
"""

from __future__ import annotations

import hashlib
import json
import math
import os
import re
from typing import Any, Callable

from config import (
    AI_INSIGHT_MAX_TOKENS,
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

NUMBER USAGE CONTRACT:
The user's JSON payload contains an "allowed_numbers" field listing
every numerical value you are permitted to use. Treat this list as
an absolute whitelist. Numbers outside this list - including
arithmetic on allowed numbers (e.g. computing a multiple from two
percentages) - are FORBIDDEN. If you need a comparison ratio, use
"lift_ratio" or "lift_ratio_rounded" directly from the JSON. Never
compute it yourself.

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
11. If the user JSON includes selection-share framing rules, follow
    them over the generic examples below.

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

ABSOLUTE RULES - VIOLATION = FAILURE:
- The input JSON contains an "allowed_numbers" array.
- You MAY ONLY use numbers from "allowed_numbers". Any other number
  is forbidden. This includes rounded variants - if 67.3 is allowed,
  67 and 67.0 are also acceptable; 68 is not.
- If a number you want to write is not in allowed_numbers, REWRITE
  the sentence to use only allowed numbers, or omit the number.
- NEVER invent a number. NEVER change a number. NEVER round outside
  the allowed list.
- NEVER imply causality unless the table explicitly shows it.
- NEVER write more than two sentences.
- If no clear pattern exists, write exactly:
  "No strong differentiation visible on this dimension."

INPUT FORMAT:
You will receive a JSON object with the computed table. Use only the
numbers in that JSON. Do not use any other numbers.
"""

TABLE_INSIGHT_PROMPT = """You are a senior strategy consultant writing
THE headline for a slide that summarises an entire analysis table.

Your job: write ONE OR TWO sentences (max 35 words total) that capture
the dominant pattern across the entire table. This is the slide's
main message - what a CFO needs to know in 5 seconds.

RULES:
1. Capture the BIG PICTURE - what is the overall pattern, not one row.
2. Lead with the most striking aggregate finding.
3. Names of subjects (Winners / Laggards / Industry / Cohort) are
   required.
4. At least one number from "allowed_numbers" must appear.
5. Strong active verb required.
6. No hedging (suggest/appear/may/could/seem).
7. No meta-commentary (the data shows / survey reveals).
8. No technical jargon (Cramer's V, statistical significance).

The user's JSON contains an "allowed_numbers" array. You may ONLY
use numbers from that array. Numbers outside it are FORBIDDEN.

GOOD EXAMPLES:
- "Winners systematically out-invest Laggards across all 8 GTM
  capabilities, with the largest gap in inside sales (67% vs 29%)."
- "Machinery and Tech sectors drive 2.3x average winner-rate, while
  Retail Laggards cluster in pricing-only strategies."

If no clear table-wide pattern exists, write:
"Distribution is broadly even across the analysed dimensions."
"""

OUTLIER_INSIGHT_PROMPT = """You are a senior strategy consultant
flagging the single most surprising data point in a table.

Your job: write ONE sentence (max 25 words) calling out the ONE
finding that would make a CEO say "wait, what?"

Pick the data point with:
- The largest gap between Winners and Laggards, OR
- The most extreme lift ratio, OR
- The most counterintuitive pattern

RULES:
1. Name the specific row/option/segment.
2. State the exact numbers (winner % vs laggard %, or lift).
3. At least one number from "allowed_numbers" must appear.
4. Frame as a surprise/contrast, not a description.
5. No hedging, no meta-commentary, no jargon.

The user's JSON contains an "allowed_numbers" array. You may ONLY
use numbers from that array.

GOOD EXAMPLES:
- "Outlier: 81% of Winners use Inside Sales vs only 12% of Laggards
  - a 6.8x gap dwarfing every other capability."
- "Surprise: Machinery sector hits a 3.4x winner multiple - more than
  triple the next-closest industry."

If no clear outlier exists, write:
"No single outlier dominates the table - patterns are evenly distributed."
"""

SHORT_LABEL_PROMPT = """You are labeling survey questions for column
headers in an Excel workbook. Each question needs a SHORT label
(4-8 words max) that captures the essence of what the question asks.

Rules:
1. Keep it to 4-8 words. Max 50 characters.
2. Capture the SUBJECT, not the format.
   "How much did your organization's revenue change in 2023?"
   -> "Revenue change 2023"
3. Use noun phrases, not full sentences.
4. Do NOT include question ID (Q17, etc) - that's added separately.
5. No filler words: "How", "What", "Please", "Your organization's"
6. Preserve time references: "2023", "next year", "this quarter"

Return STRICT JSON, no markdown:
{
  "labels": {
    "Q17": "Planned revenue growth 2023",
    "Q18": "Actual revenue change 2023 vs 2022",
    "Q22": "Expected market size change 2024"
  }
}

Every question_id you receive must have a label.
"""

THEME_CATEGORIZATION_PROMPT = """You are a survey analyst organizing
questions into themes for a workbook with multiple sheets.

You will be given a list of survey questions with their IDs and text.
Your job: group these questions into 4-8 themes. Each theme becomes
one sheet in an Excel workbook.

RULES:
1. Each question must belong to exactly ONE theme.
2. Theme names must be short (2-4 words), business-friendly,
   sheet-name safe (no slashes, no quotes, no special chars).
3. Themes should reflect actual survey content. Common themes
   for business surveys:
   - "Revenue & Growth"
   - "Customer Strategy"
   - "Talent & Operations"
   - "Technology & AI"
   - "Pricing & Commercial"
   - "Geography & Market"
   - "Demographics"
4. ALL demographic questions (Industry, Region, Country, Company Size,
   Role, Tenure, Function, etc.) must be grouped under exactly one
   theme named "Demographics".
5. Output STRICT JSON, no preamble, no markdown:

{
  "themes": [
    {
      "name": "Revenue & Growth",
      "question_ids": ["Q17", "Q18", "Q19", "Q22"]
    },
    {
      "name": "Customer Strategy",
      "question_ids": ["Q30", "Q31", "Q32"]
    },
    {
      "name": "Demographics",
      "question_ids": ["Q2", "Q5", "Q6", "Q13", "Q14"]
    }
  ]
}

Every question ID provided to you must appear in exactly one theme.
"""

PRIORITY_DEMOGRAPHIC_CATEGORIES = """You are categorizing survey
questions as demographics. Given a list of survey questions, identify
which ones match each of these standard demographic categories. A
question can match at most one category.

PRIORITY CATEGORIES (in order):
Tier 1 - Basic:
  - age, gender, country, state_city, nationality, marital_status

Tier 2 - Education & Career:
  - education_level, field_of_study, employment_status, job_title,
    industry, years_experience, income_range

Tier 3 - Household:
  - household_size, dependents, living_arrangement

Tier 4 - Technology:
  - devices_used, internet_access, social_platforms, screen_time

Tier 5 - B2B:
  - company_size, department, seniority, decision_authority

Tier 6 - Lifestyle:
  - hobbies, shopping_habits, travel_frequency, fitness_habits

Return STRICT JSON:
{
  "matches": [
    {"question_id": "Q2", "category": "country", "tier": 1},
    {"question_id": "Q5", "category": "industry", "tier": 2}
  ],
  "other_demographics": ["Q47", "Q83"]
}

Only include questions that are clearly demographics. "other_demographics"
is for demographic questions that don't fit the priority list - these
appear AFTER priority matches in the workbook filters.
"""

TIER_CATEGORY_ORDER = {
    1: ["age", "gender", "country", "state_city", "nationality", "marital_status"],
    2: [
        "education_level",
        "field_of_study",
        "employment_status",
        "job_title",
        "industry",
        "years_experience",
        "income_range",
    ],
    3: ["household_size", "dependents", "living_arrangement"],
    4: ["devices_used", "internet_access", "social_platforms", "screen_time"],
    5: ["company_size", "department", "seniority", "decision_authority"],
    6: ["hobbies", "shopping_habits", "travel_frequency", "fitness_habits"],
}

AI_SYSTEM_PROMPT = SYSTEM_PROMPT

USER_PROMPT_TEMPLATE = """Here is the computed survey data table:

{table_json}

Write one PPT-ready slide headline following the system prompt rules.
Return ONLY the headline sentence. No preamble, no explanation,
no punctuation at the end except if the sentence naturally requires it.
"""

TEMPLATE_INSIGHT = "No AI insight available — review the data table above."

_NUMBER_PATTERN = re.compile(r"(\d+(?:\.\d+)?)\s*(%|x|X)?")


def generate_short_labels(
    questions: list[dict],
    cache: dict | None = None,
) -> dict[str, str]:
    """Generate compact question labels for Excel headers."""

    if not questions:
        return {}

    normalized = [
        {
            "question_id": str(question.get("question_id", "")),
            "question_text": str(question.get("question_text", "")),
        }
        for question in questions
        if question.get("question_id")
    ]
    cache_key = "short_labels_" + hashlib.md5(
        json.dumps(
            [
                {
                    "qid": question["question_id"],
                    "text": question["question_text"][:100],
                }
                for question in normalized
            ],
            sort_keys=True,
        ).encode()
    ).hexdigest()

    if cache is not None and cache_key in cache:
        return cache[cache_key]

    api_key = os.environ.get("PORTKEY_API_KEY", "").strip()
    if not api_key or OpenAI is None:
        labels = _short_label_fallback(normalized)
        if cache is not None:
            cache[cache_key] = labels
        return labels

    user_prompt = "Generate short labels for these questions:\n\n" + json.dumps(
        [
            {
                "question_id": question["question_id"],
                "question_text": question["question_text"][:200],
            }
            for question in normalized
        ],
        indent=2,
    )

    try:
        client = OpenAI(api_key=api_key, base_url=PORTKEY_BASE_URL)
        response = client.chat.completions.create(
            model=PORTKEY_DEFAULT_MODEL,
            messages=[
                {"role": "system", "content": SHORT_LABEL_PROMPT},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.0,
            max_tokens=4000,
            timeout=25,
        )
        raw = response.choices[0].message.content or ""
        text = raw.strip()
        if text.startswith("```"):
            text = text.strip("`").strip()
        if text.startswith("json"):
            text = text[4:].strip()
        parsed = json.loads(text)
        labels = {
            str(question_id): str(label)[:50]
            for question_id, label in parsed.get("labels", {}).items()
        }
        fallback = _short_label_fallback(normalized)
        for question in normalized:
            qid = question["question_id"]
            if not labels.get(qid):
                labels[qid] = fallback[qid]
        if cache is not None:
            cache[cache_key] = labels
        return labels
    except Exception:
        labels = _short_label_fallback(normalized)
        if cache is not None:
            cache[cache_key] = labels
        return labels


def _short_label_fallback(questions: list[dict]) -> dict[str, str]:
    labels: dict[str, str] = {}
    for question in questions:
        text = str(question.get("question_text", ""))
        labels[str(question["question_id"])] = text[:40] + ("..." if len(text) > 40 else "")
    return labels


def generate_insight(
    table_payload: dict,
    table_kind: str = "generic",
    title_hint: str = "",
    cache: dict | None = None,
    use_premium: bool = False,
) -> InsightResult:
    """Generate one PPT-ready headline for a computed table."""

    return _generate_with_prompt(
        table_payload=table_payload,
        table_kind=table_kind,
        system_prompt=SYSTEM_PROMPT,
        cache=cache,
        cache_prefix="insight",
        title_hint=title_hint,
        use_premium=use_premium,
    )


def generate_table_insight(
    table_payload: dict,
    table_kind: str,
    cache: dict | None = None,
) -> InsightResult:
    """Generate the dominant pattern across an entire table."""

    return _generate_with_prompt(
        table_payload=table_payload,
        table_kind=table_kind,
        system_prompt=TABLE_INSIGHT_PROMPT,
        cache=cache,
        cache_prefix="table",
    )


def generate_outlier_insight(
    table_payload: dict,
    table_kind: str,
    cache: dict | None = None,
) -> InsightResult:
    """Generate one call-out for the single most extreme data point."""

    return _generate_with_prompt(
        table_payload=table_payload,
        table_kind=table_kind,
        system_prompt=OUTLIER_INSIGHT_PROMPT,
        cache=cache,
        cache_prefix="outlier",
    )


def categorize_questions_into_themes(
    questions: list[dict],
    cache: dict | None = None,
) -> dict:
    """Categorize survey questions into workbook themes via AI."""

    questions_summary = [
        {
            "question_id": question.get("question_id"),
            "question_text": question.get("question_text", "")[:120],
            "question_type": question.get("question_type", ""),
            "is_demographic": bool(question.get("is_demographic", False)),
        }
        for question in questions
        if question.get("question_id")
    ]
    cache_key = "themes_" + hashlib.md5(
        json.dumps(questions_summary, sort_keys=True, default=str).encode()
    ).hexdigest()
    if cache is not None and cache_key in cache:
        return cache[cache_key]

    api_key = os.environ.get("PORTKEY_API_KEY", "").strip()
    if not api_key or OpenAI is None:
        result = _theme_fallback(questions, error="AI unavailable")
        if cache is not None:
            cache[cache_key] = result
        return result

    user_prompt = (
        "Categorize these survey questions into 4-8 themes:\n\n"
        + json.dumps(questions_summary, indent=2)
    )

    try:
        client = OpenAI(api_key=api_key, base_url=PORTKEY_BASE_URL)
        response = client.chat.completions.create(
            model=PORTKEY_DEFAULT_MODEL,
            messages=[
                {"role": "system", "content": THEME_CATEGORIZATION_PROMPT},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.0,
            max_tokens=2000,
            timeout=25,
        )
        raw = response.choices[0].message.content
        result = _parse_themes_response(raw, questions)
    except Exception as exc:
        result = _theme_fallback(questions, error=f"AI error: {exc}")

    if cache is not None:
        cache[cache_key] = result
    return result


def _parse_themes_response(raw: str, questions: list[dict]) -> dict:
    """Parse AI theme JSON and ensure each provided question appears once."""

    text = (raw or "").strip()
    if text.startswith("```"):
        lines = text.splitlines()
        if lines and lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].startswith("```"):
            lines = lines[:-1]
        text = "\n".join(lines).strip()
        if text.startswith("json"):
            text = text[4:].strip()

    try:
        parsed = json.loads(text)
    except json.JSONDecodeError as exc:
        return _theme_fallback(questions, error=f"Invalid JSON from AI: {exc}")

    themes = parsed.get("themes", [])
    if not isinstance(themes, list) or not themes:
        return _theme_fallback(questions, error="No themes in response")

    all_qids = [
        question.get("question_id")
        for question in questions
        if question.get("question_id")
    ]
    allowed_qids = set(all_qids)
    covered_qids: set[str] = set()
    cleaned_themes = []

    for theme in themes:
        if not isinstance(theme, dict):
            continue
        theme_name = str(theme.get("name") or "Theme").strip() or "Theme"
        cleaned_qids = []
        for qid in theme.get("question_ids", []):
            if qid in allowed_qids and qid not in covered_qids:
                cleaned_qids.append(qid)
                covered_qids.add(qid)
        if cleaned_qids:
            cleaned_themes.append({"name": theme_name, "question_ids": cleaned_qids})

    missing = [qid for qid in all_qids if qid not in covered_qids]
    if missing:
        cleaned_themes.append({"name": "Other", "question_ids": missing})

    if not cleaned_themes:
        return _theme_fallback(questions, error="No valid themes in response")

    return {
        "themes": cleaned_themes,
        "was_template": False,
        "error_message": "",
    }


def _theme_fallback(questions: list[dict], error: str = "") -> dict:
    """Deterministic fallback: split demographics from all other questions."""

    demo_qids = [
        question["question_id"]
        for question in questions
        if question.get("is_demographic") and question.get("question_id")
    ]
    other_qids = [
        question["question_id"]
        for question in questions
        if not question.get("is_demographic") and question.get("question_id")
    ]

    themes = []
    if demo_qids:
        themes.append({"name": "Demographics", "question_ids": demo_qids})
    if other_qids:
        themes.append({"name": "All Questions", "question_ids": other_qids})

    return {
        "themes": themes,
        "was_template": True,
        "error_message": error,
    }


def categorize_demographic_questions(
    questions: list[dict],
    cache: dict | None = None,
) -> dict:
    """Categorize demographic questions by priority tier and category."""

    if not questions:
        return {"priority_ordered": [], "categories": {}, "was_template": True}

    summary = [
        {
            "qid": question.get("question_id"),
            "text": question.get("question_text", "")[:80],
            "question_type": question.get("question_type", ""),
        }
        for question in questions
        if question.get("question_id")
    ]
    cache_key = "demo_priority_" + hashlib.md5(
        json.dumps(summary, sort_keys=True, default=str).encode()
    ).hexdigest()
    if cache is not None and cache_key in cache:
        return cache[cache_key]

    api_key = os.environ.get("PORTKEY_API_KEY", "").strip()
    if not api_key or OpenAI is None:
        result = {
            "priority_ordered": sorted(
                question["question_id"]
                for question in questions
                if question.get("question_id")
            ),
            "categories": {},
            "was_template": True,
        }
        if cache is not None:
            cache[cache_key] = result
        return result

    user_prompt = "Categorize these demographic questions:\n\n" + json.dumps(
        [
            {
                "question_id": question["question_id"],
                "question_text": question.get("question_text", "")[:120],
            }
            for question in questions
            if question.get("question_id")
        ],
        indent=2,
    )

    try:
        client = OpenAI(api_key=api_key, base_url=PORTKEY_BASE_URL)
        response = client.chat.completions.create(
            model=PORTKEY_DEFAULT_MODEL,
            messages=[
                {"role": "system", "content": PRIORITY_DEMOGRAPHIC_CATEGORIES},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.0,
            max_tokens=1500,
            timeout=25,
        )
        raw = response.choices[0].message.content
        result = _parse_demographic_priority(raw, questions)
    except Exception as exc:
        result = {
            "priority_ordered": sorted(
                question["question_id"]
                for question in questions
                if question.get("question_id")
            ),
            "categories": {},
            "was_template": True,
            "error_message": str(exc),
        }

    if cache is not None:
        cache[cache_key] = result
    return result


def _parse_demographic_priority(raw: str, questions: list[dict]) -> dict:
    """Parse AI demographic priority JSON and sort by tier/category."""

    text = (raw or "").strip()
    if text.startswith("```"):
        lines = text.splitlines()
        if lines and lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].startswith("```"):
            lines = lines[:-1]
        text = "\n".join(lines).strip()
        if text.startswith("json"):
            text = text[4:].strip()

    try:
        parsed = json.loads(text)
    except json.JSONDecodeError as exc:
        return {
            "priority_ordered": sorted(
                question["question_id"]
                for question in questions
                if question.get("question_id")
            ),
            "categories": {},
            "was_template": True,
            "error_message": f"Invalid JSON from AI: {exc}",
        }

    known_qids = [
        question["question_id"]
        for question in questions
        if question.get("question_id")
    ]
    known_qid_set = set(known_qids)
    categories: dict[str, str] = {}
    sortable_matches = []

    for match in parsed.get("matches", []):
        if not isinstance(match, dict):
            continue
        qid = match.get("question_id")
        category = str(match.get("category", "")).strip()
        if qid not in known_qid_set or not category:
            continue
        try:
            tier = int(match.get("tier", 99))
        except (TypeError, ValueError):
            tier = 99
        category_order = TIER_CATEGORY_ORDER.get(tier, [])
        category_pos = (
            category_order.index(category)
            if category in category_order
            else len(category_order)
        )
        categories[qid] = category
        sortable_matches.append((tier, category_pos, qid))

    sortable_matches.sort()
    ordered = [qid for _tier, _category_pos, qid in sortable_matches]
    for qid in parsed.get("other_demographics", []):
        if qid in known_qid_set and qid not in ordered:
            ordered.append(qid)
    for qid in known_qids:
        if qid not in ordered:
            ordered.append(qid)

    return {
        "priority_ordered": ordered,
        "categories": categories,
        "was_template": False,
    }


def _generate_with_prompt(
    *,
    table_payload: dict,
    table_kind: str,
    system_prompt: str,
    cache: dict | None = None,
    cache_prefix: str = "insight",
    title_hint: str = "",
    use_premium: bool = False,
) -> InsightResult:
    if cache is None:
        return _call_openai(
            table_payload=table_payload,
            table_kind=table_kind,
            system_prompt=system_prompt,
            title_hint=title_hint,
            use_premium=use_premium,
        )

    key = _payload_hash(f"{cache_prefix}:{table_kind}", table_payload)
    if key in cache:
        return cache[key]

    result = _call_openai(
        table_payload=table_payload,
        table_kind=table_kind,
        system_prompt=system_prompt,
        title_hint=title_hint,
        use_premium=use_premium,
    )
    cache[key] = result
    return result


def _call_openai(
    *,
    table_payload: dict,
    table_kind: str,
    system_prompt: str,
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
    formatted_payload = _format_payload(table_kind, table_payload)
    user_prompt = USER_PROMPT_TEMPLATE.format(table_json=formatted_payload)

    try:
        client = OpenAI(api_key=api_key, base_url=PORTKEY_BASE_URL)
        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.0,
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

    allowed = table_payload.get("allowed_numbers", []) if isinstance(table_payload, dict) else []
    if not allowed:
        try:
            formatted_payload_dict = json.loads(formatted_payload)
            allowed = formatted_payload_dict.get("allowed_numbers", [])
        except (json.JSONDecodeError, TypeError):
            allowed = []

    if allowed:
        is_valid, invalid_numbers = _validate_numbers(headline, allowed)
        if not is_valid:
            return _template_fallback(
                table_payload,
                table_kind,
                title_hint,
                error_message=(
                    f"Hallucinated numbers detected: {invalid_numbers}. "
                    f"Allowed: {allowed[:10]}{'...' if len(allowed) > 10 else ''}"
                ),
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


def _extract_numbers_from_text(text: str) -> list[float]:
    """Extract all numbers from a sentence."""

    numbers: list[float] = []
    for match in _NUMBER_PATTERN.finditer(text):
        try:
            numbers.append(float(match.group(1)))
        except ValueError:
            continue
    return numbers


def _validate_numbers(
    headline: str,
    allowed: list[float],
    tolerance: float = 0.5,
) -> tuple[bool, list[float]]:
    """Return whether every number in the headline is permitted."""

    extracted = _extract_numbers_from_text(headline)
    invalid: list[float] = []
    allowed_floats: list[float] = []
    for value in allowed:
        try:
            allowed_floats.append(float(value))
        except (TypeError, ValueError):
            continue
    for number in extracted:
        if not any(abs(number - allowed_number) <= tolerance for allowed_number in allowed_floats):
            invalid.append(number)
    return (len(invalid) == 0, invalid)


def _build_allowed_numbers(values: list[float]) -> list[float]:
    """Build list of permitted numbers, including rounded variants."""

    out: set[float] = set()
    for value in values:
        if value is None:
            continue
        try:
            float_value = float(value)
        except (TypeError, ValueError):
            continue
        if not math.isfinite(float_value):
            continue
        out.add(round(float_value, 2))
        out.add(round(float_value, 1))
        out.add(round(float_value))
    return sorted(out)


def _format_differentiator_payload(payload: dict) -> str:
    winner_rate = payload.get("winner_rate", 0)
    loser_rate = payload.get("loser_rate", 0)
    lift = payload.get("lift", 1)
    cramers_v = payload.get("cramers_v", 0)
    winner_n = payload.get("winner_n", 0)
    loser_n = payload.get("loser_n", 0)

    allowed = _build_allowed_numbers(
        [
            winner_rate * 100,
            loser_rate * 100,
            lift,
            round(winner_rate * 100),
            round(loser_rate * 100),
            round(lift, 1),
            round(lift, 2),
            winner_n,
            loser_n,
            cramers_v,
        ]
    )

    return json.dumps(
        {
            "analysis_type": "winner_vs_laggard_differentiator",
            "question": payload.get("question_text", ""),
            "top_differentiating_option": payload.get("top_option", ""),
            "winner_selection_rate": f"{winner_rate:.1%}",
            "laggard_selection_rate": f"{loser_rate:.1%}",
            "winner_selection_pct_int": round(winner_rate * 100),
            "laggard_selection_pct_int": round(loser_rate * 100),
            "lift_ratio": f"{lift:.2f}x",
            "lift_ratio_rounded": f"{lift:.1f}x",
            "winner_n": winner_n,
            "laggard_n": loser_n,
            "allowed_numbers": allowed,
        },
        indent=2,
    )


def _format_winner_profile_payload(payload: dict) -> str:
    traits = payload.get("traits", [])
    all_numbers = [payload.get("winner_n", 0), payload.get("loser_n", 0)]
    defining_traits: list[dict[str, Any]] = []
    for trait in traits[:5]:
        winner_rate = trait.get("winner_rate", 0)
        loser_rate = trait.get("loser_rate", 0)
        lift = trait.get("lift", 1)
        rate_gap = trait.get("rate_gap", 0)
        all_numbers.extend(
            [
                winner_rate * 100,
                loser_rate * 100,
                lift,
                rate_gap * 100,
                round(winner_rate * 100),
                round(loser_rate * 100),
                round(lift, 1),
            ]
        )
        defining_traits.append(
            {
                "question": trait.get("question_id", ""),
                "winning_option": trait.get("option_label", ""),
                "winner_rate": f"{winner_rate:.1%}",
                "laggard_rate": f"{loser_rate:.1%}",
                "lift": f"{lift:.2f}x",
                "rate_gap": f"+{rate_gap:.1%}",
            }
        )

    return json.dumps(
        {
            "analysis_type": "winner_profile_summary",
            "winner_label": payload.get("winner_label", "Winner"),
            "laggard_label": payload.get("loser_label", "Laggard"),
            "winner_n": payload.get("winner_n", 0),
            "laggard_n": payload.get("loser_n", 0),
            "defining_traits": defining_traits,
            "allowed_numbers": _build_allowed_numbers(all_numbers),
        },
        indent=2,
    )


def _format_winner_profile_trait_payload(payload: dict) -> str:
    """Payload for one profile trait showing both sides of the story."""
    winner_rate = payload.get("winner_rate", 0)
    loser_rate = payload.get("loser_rate", 0)
    lift = payload.get("lift", 1)
    rate_gap = payload.get("rate_gap", 0)
    lag_w_rate = payload.get("laggard_top_option_winner_rate", 0)
    lag_l_rate = payload.get("laggard_top_option_loser_rate", 0)
    lag_gap = lag_l_rate - lag_w_rate

    allowed = _build_allowed_numbers(
        [
            winner_rate * 100,
            loser_rate * 100,
            lift,
            rate_gap * 100,
            lag_w_rate * 100,
            lag_l_rate * 100,
            round(winner_rate * 100),
            round(loser_rate * 100),
            round(lift, 1),
            round(rate_gap * 100),
            round(lag_w_rate * 100),
            round(lag_l_rate * 100),
            abs(lag_gap) * 100,
            round(abs(lag_gap) * 100),
        ]
    )

    return json.dumps(
        {
            "analysis_type": "winner_profile_trait_both_sides",
            "question": payload.get("question_text", ""),
            "selection_share_contract": (
                "Each percentage is the proportion of the named cohort who "
                "selected, chose, or reported the option on the survey question. "
                "It is never a magnitude, budget percentage, allocation, spend, "
                "investment amount, revenue amount, margin, or growth rate of "
                "the option."
            ),
            "required_framing": (
                "Use selection framing only, such as 'X% of Winners chose "
                "<option>' or 'Winners are Nx more likely to select <option>'. "
                "Anchor the option to the survey question text."
            ),
            "forbidden_framing": (
                "Do not say the cohort allocated, spent, invested, dedicated, "
                "or devoted the percentage. Do not say the cohort grew, expects "
                "percentage growth, or has a percentage of budget, revenue, "
                "margin, spend, sales, portfolio, profit, or growth."
            ),
            "forbidden_verbs": [
                "allocate",
                "spend",
                "invest",
                "dedicate",
                "devote",
                "grow",
                "expect percentage growth",
            ],
            "winner_label": payload.get("winner_label", "Winner"),
            "laggard_label": payload.get("laggard_label", "Laggard"),
            "winners_top_option": payload.get("option_label", ""),
            "winners_top_option_winner_pct": f"{winner_rate:.1%}",
            "winners_top_option_laggard_pct": f"{loser_rate:.1%}",
            "winners_advantage_gap": f"+{rate_gap:.1%}",
            "laggards_top_option": payload.get("laggard_top_option_label", ""),
            "laggards_top_option_winner_pct": f"{lag_w_rate:.1%}",
            "laggards_top_option_laggard_pct": f"{lag_l_rate:.1%}",
            "laggards_concentration_gap": f"+{lag_gap:.1%}",
            "allowed_numbers": allowed,
        },
        indent=2,
    )


def _format_single_cut_payload(payload: dict) -> str:
    distribution = payload.get("distribution", {})
    all_numbers = [payload.get("valid_n", 0), payload.get("total_n", 0)]
    all_numbers.extend(_extract_numeric_values(distribution))
    all_numbers.extend(_extract_rate_percentages(distribution))
    all_numbers.extend(_extract_numeric_values(payload.get("top_option_pct", "")))

    return json.dumps(
        {
            "analysis_type": "single_question_distribution",
            "question": payload.get("question_text", ""),
            "question_type": payload.get("question_type", ""),
            "total_n": payload.get("valid_n", 0),
            "distribution": distribution,
            "top_option": payload.get("top_option", ""),
            "top_option_pct": payload.get("top_option_pct", ""),
            "allowed_numbers": _build_allowed_numbers(all_numbers),
        },
        indent=2,
    )


def _format_filtered_single_cut_payload(payload: dict) -> str:
    distribution = payload.get("distribution", {})
    all_numbers = [
        payload.get("filtered_n", 0),
        payload.get("total_n", 0),
        payload.get("valid_n", 0),
    ]
    all_numbers.extend(_extract_numeric_values(distribution))
    all_numbers.extend(_extract_rate_percentages(distribution))
    all_numbers.extend(_extract_numeric_values(payload.get("top_option_pct", "")))

    return json.dumps(
        {
            "analysis_type": "filtered_single_question_distribution",
            "question": payload.get("question_text", ""),
            "filter_applied": payload.get("filter_description", ""),
            "filtered_n": payload.get("filtered_n", 0),
            "total_n": payload.get("total_n", 0),
            "distribution": distribution,
            "top_option": payload.get("top_option", ""),
            "top_option_pct": payload.get("top_option_pct", ""),
            "allowed_numbers": _build_allowed_numbers(all_numbers),
        },
        indent=2,
    )


def _format_cross_cut_payload(payload: dict) -> str:
    table = payload.get("table", {})
    all_numbers = _extract_numeric_values(table)
    all_numbers.extend(_extract_numeric_values(payload.get("key_finding", "")))

    return json.dumps(
        {
            "analysis_type": "cross_cut_comparison",
            "title": payload.get("title", ""),
            "source_questions": payload.get("source_questions", []),
            "analysis_type_detail": payload.get("analysis_type_detail", ""),
            "table": table,
            "key_finding": payload.get("key_finding", ""),
            "allowed_numbers": _build_allowed_numbers(all_numbers),
        },
        indent=2,
    )


def _format_differentiator_table_payload(payload: dict) -> str:
    """Payload for table-level insight across all differentiators."""

    differentiators = payload.get("differentiators", [])
    all_numbers: list[float] = []
    rows_summary: list[dict[str, Any]] = []
    for differentiator in differentiators[:15]:
        winner_rate = differentiator.get("winner_rate", 0)
        loser_rate = differentiator.get("loser_rate", 0)
        lift = differentiator.get("lift", 1)
        all_numbers.extend(
            [
                winner_rate * 100,
                loser_rate * 100,
                lift,
                round(winner_rate * 100),
                round(loser_rate * 100),
                round(lift, 1),
                round(lift, 2),
            ]
        )
        rows_summary.append(
            {
                "question": differentiator.get("question_text", "")[:60],
                "top_option": differentiator.get("top_option_label", ""),
                "winner_rate": f"{winner_rate:.1%}",
                "laggard_rate": f"{loser_rate:.1%}",
                "lift": f"{lift:.2f}x",
            }
        )

    all_numbers.extend(
        [
            payload.get("winner_n", 0),
            payload.get("loser_n", 0),
            len(differentiators),
        ]
    )

    return json.dumps(
        {
            "analysis_type": "differentiator_table_summary",
            "outcome_variable": payload.get("outcome_question_id", ""),
            "winner_label": payload.get("winner_label", "Winner"),
            "laggard_label": payload.get("loser_label", "Laggard"),
            "winner_n": payload.get("winner_n", 0),
            "laggard_n": payload.get("loser_n", 0),
            "total_differentiators": len(differentiators),
            "rows": rows_summary,
            "allowed_numbers": _build_allowed_numbers(all_numbers),
        },
        indent=2,
    )


def _format_outlier_payload(payload: dict) -> str:
    """Payload for outlier callout."""

    parsed = json.loads(_format_differentiator_table_payload(payload))
    parsed["analysis_type"] = "table_outlier_callout"
    return json.dumps(parsed, indent=2)


FORMATTERS: dict[str, Callable[[dict], str]] = {
    "differentiator": _format_differentiator_payload,
    "winner_profile": _format_winner_profile_payload,
    "winner_profile_trait": _format_winner_profile_trait_payload,
    "single_cut": _format_single_cut_payload,
    "filtered_single_cut": _format_filtered_single_cut_payload,
    "cross_cut": _format_cross_cut_payload,
    "differentiator_table": _format_differentiator_table_payload,
    "outlier": _format_outlier_payload,
}


def _format_payload(table_kind: str, payload: dict) -> str:
    formatter = FORMATTERS.get(table_kind, lambda p: _format_generic_payload(p))
    return formatter(payload)


def _format_generic_payload(payload: dict) -> str:
    payload_copy = dict(payload)
    payload_copy.setdefault(
        "allowed_numbers",
        _build_allowed_numbers(_extract_numeric_values(payload)),
    )
    return json.dumps(payload_copy, indent=2, default=str)


def _payload_hash(table_kind: str, payload: dict) -> str:
    raw = json.dumps({"kind": table_kind, "payload": payload}, sort_keys=True, default=str)
    return hashlib.md5(raw.encode()).hexdigest()


def _extract_numeric_values(value: Any) -> list[float]:
    if value is None:
        return []
    if isinstance(value, bool):
        return []
    if isinstance(value, (int, float)):
        return [float(value)]
    if isinstance(value, dict):
        values: list[float] = []
        for item in value.values():
            values.extend(_extract_numeric_values(item))
        return values
    if isinstance(value, (list, tuple)):
        values = []
        for item in value:
            values.extend(_extract_numeric_values(item))
        return values
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return []
        try:
            return [float(text.rstrip("%xX"))]
        except ValueError:
            if "%" in text or "x" in text or "X" in text:
                return _extract_numbers_from_text(text)
    return []


def _extract_rate_percentages(value: Any) -> list[float]:
    if isinstance(value, dict):
        values: list[float] = []
        for key, item in value.items():
            if str(key).lower() in {"rate", "pct", "percent", "percentage"}:
                try:
                    values.append(float(item) * 100)
                except (TypeError, ValueError):
                    pass
            values.extend(_extract_rate_percentages(item))
        return values
    if isinstance(value, (list, tuple)):
        values = []
        for item in value:
            values.extend(_extract_rate_percentages(item))
        return values
    return []


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
