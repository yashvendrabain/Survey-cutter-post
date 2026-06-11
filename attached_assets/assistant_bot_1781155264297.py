"""Survey Analysis Engine — in-app assistant (chatbot) logic.

DESIGN CONTRACT (do not violate):
  Calculation-first, AI-second. The LLM NEVER computes a count, percentage,
  mean, lift, or cohort assignment, and NEVER decides whether a hypothesis is
  "true". The LLM is used for exactly three bounded jobs:
    1. Intent classification  (which of the 3 handlers)
    2. NL -> question-ID mapping (hypothesis validator)
    3. Phrasing a result that Python already computed/validated

  Every number the user sees is produced by the deterministic engines
  (compute_single_cuts / compute_cross_cuts) and re-checked against the
  source table before it is shown. If the numbers can't be verified, the
  bot shows the table and a hedged statement, never an invented claim.

This module is import-light at module load: heavy deps (the engines, Portkey)
are imported lazily inside functions so it can be unit-tested in isolation.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from typing import Any, Callable

# Minimum cell base below which we refuse to make a confident claim.
MIN_CELL_N_FOR_CONFIDENT_CLAIM = 30

# Portkey config mirrors src/ai_insights.py (single source of truth there).
PORTKEY_BASE_URL = "https://portkey.bain.dev/v1"
ROUTER_TEMPERATURE = 0.0      # deterministic intent routing
MAPPING_TEMPERATURE = 0.0     # deterministic question mapping
PHRASING_TEMPERATURE = 0.1    # light phrasing only


# ---------------------------------------------------------------------------
# Result container
# ---------------------------------------------------------------------------

@dataclass
class BotReply:
    """One assistant turn. `table` (if present) is a list-of-dict the UI renders.

    `was_grounded` is False whenever the bot fell back to a template/hedge
    because it could not verify numbers or could not reach the LLM — the UI
    should badge those replies so the user knows it's not a model assertion.
    """
    text: str
    intent: str
    table: list[dict[str, Any]] | None = None
    table_caption: str = ""
    caveats: list[str] = field(default_factory=list)
    was_grounded: bool = True
    debug: dict[str, Any] = field(default_factory=dict)


# ---------------------------------------------------------------------------
# Lazy LLM helper (delegates to the project's existing Portkey client)
# ---------------------------------------------------------------------------

def _llm_json(system: str, user: str, *, temperature: float) -> dict[str, Any] | None:
    """Call the shared LLM client and parse a JSON object, or None on any failure.

    Reuses src.ai_insights' Portkey plumbing so there is ONE place that holds
    credentials and base URL. Never raises — callers must handle None.
    """
    try:
        from src.ai_insights import _portkey_chat_json  # type: ignore
    except Exception:
        _portkey_chat_json = None  # type: ignore

    if _portkey_chat_json is not None:
        try:
            return _portkey_chat_json(system=system, user=user, temperature=temperature)
        except Exception:
            return None

    # Fallback: direct client only if the project exposes a raw caller.
    try:
        from src.ai_insights import _raw_chat  # type: ignore
        raw = _raw_chat(system=system, user=user, temperature=temperature)
        return _safe_json(raw)
    except Exception:
        return None


def _safe_json(text: str | None) -> dict[str, Any] | None:
    if not text:
        return None
    cleaned = re.sub(r"^```(?:json)?|```$", "", text.strip(), flags=re.MULTILINE).strip()
    try:
        obj = json.loads(cleaned)
        return obj if isinstance(obj, dict) else None
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Intent routing  (LLM job #1 — bounded to 3 labels)
# ---------------------------------------------------------------------------

INTENT_TOOL = "tool_help"
INTENT_HYPOTHESIS = "hypothesis"
INTENT_SURVEY = "survey_structure"
INTENT_UNKNOWN = "unknown"

_ROUTER_SYSTEM = (
    "You route a user message to exactly one label. Reply with JSON only: "
    '{"intent": "<label>"}. Labels:\n'
    "- tool_help: questions about how the Survey Analysis Engine works "
    "(features, buttons, exports, what a control does).\n"
    "- hypothesis: the user states or asks whether a relationship between two "
    "survey measures holds (e.g. 'do high-efficiency firms grow revenue more?').\n"
    "- survey_structure: questions about the CONTENT of the uploaded survey "
    "(does it contain an NPS question? how many respondents? what does Q12 ask?).\n"
    "Pick the single best label."
)


def classify_intent(message: str) -> str:
    obj = _llm_json(_ROUTER_SYSTEM, message, temperature=ROUTER_TEMPERATURE)
    intent = (obj or {}).get("intent", "")
    if intent in (INTENT_TOOL, INTENT_HYPOTHESIS, INTENT_SURVEY):
        return intent
    # Deterministic keyword fallback so routing still works without the LLM.
    low = message.lower()
    if any(w in low for w in ("how do i", "where is", "what does", "button", "export", "download", "filter", "tour")):
        return INTENT_TOOL
    if any(w in low for w in ("is there", "does the survey", "how many respondent", "what is q", "nps question")):
        return INTENT_SURVEY
    if any(w in low for w in ("is it true", "correlat", "relationship", "do companies", "are firms", "linked to", "associated with")):
        return INTENT_HYPOTHESIS
    return INTENT_UNKNOWN


# ---------------------------------------------------------------------------
# Handler A — tool help (grounded in the handover FAQ)
# ---------------------------------------------------------------------------

def answer_tool_help(message: str, faq_ground_truth: str) -> BotReply:
    """Answer using the FAQ as the ONLY source. No FAQ match -> say so."""
    system = (
        "You are the in-app help assistant for the Survey Analysis Engine. "
        "Answer ONLY from the FAQ/REFERENCE text provided. If the answer is not "
        "in the reference, say you don't have that documented and suggest the "
        "closest documented topic. Never invent feature behavior. Keep it short."
    )
    user = f"REFERENCE:\n{faq_ground_truth}\n\nUSER QUESTION:\n{message}"
    obj = _llm_json(
        system + " Reply JSON: {\"answer\": \"...\", \"grounded\": true|false}",
        user,
        temperature=PHRASING_TEMPERATURE,
    )
    if obj and obj.get("answer"):
        return BotReply(
            text=str(obj["answer"]),
            intent=INTENT_TOOL,
            was_grounded=bool(obj.get("grounded", True)),
        )
    return BotReply(
        text=("I can't reach the help model right now. The user guide and the "
              "guided tour (top-right) cover the main features."),
        intent=INTENT_TOOL,
        was_grounded=False,
    )


# ---------------------------------------------------------------------------
# Handler C — survey structure (grounded in the classified schema)
# ---------------------------------------------------------------------------

def _schema_facts(schema: Any) -> dict[str, Any]:
    """Extract deterministic facts from the schema. No LLM involved."""
    questions = list(getattr(schema, "questions", []) or [])
    by_type: dict[str, list[str]] = {}
    for q in questions:
        t = getattr(getattr(q, "question_type", None), "value", str(getattr(q, "question_type", "")))
        by_type.setdefault(t, []).append(getattr(q, "canonical_id", "?"))
    return {
        "total_questions": len(questions),
        "total_respondents": int(getattr(schema, "total_respondents", 0) or 0),
        "types": by_type,
        "ids": [getattr(q, "canonical_id", "?") for q in questions],
    }


def answer_survey_structure(message: str, schema: Any) -> BotReply:
    """Answer content questions from schema facts. Python decides facts; LLM phrases."""
    if schema is None:
        return BotReply(
            text="No survey is loaded yet. Upload a survey and run the analysis first.",
            intent=INTENT_SURVEY,
            was_grounded=True,
        )
    facts = _schema_facts(schema)

    # Direct deterministic answers for the common "is there an X question?" form.
    low = message.lower()
    type_keywords = {
        "nps": "NPS", "net promoter": "NPS",
        "rank": "RANK_ORDER", "ranking": "RANK_ORDER",
        "grid": "GRID", "matrix": "GRID",
        "multi-select": "MULTI_SELECT_BINARY", "multi select": "MULTI_SELECT_BINARY",
        "numeric": "DIRECT_NUMERIC", "number": "DIRECT_NUMERIC",
        "open text": "OPEN_TEXT", "open-ended": "OPEN_TEXT",
    }
    for kw, tcode in type_keywords.items():
        if kw in low and ("is there" in low or "any" in low or "does" in low or "have" in low):
            hits = [
                qid for t, ids in facts["types"].items() if tcode in t for qid in ids
            ]
            if hits:
                return BotReply(
                    text=f"Yes — {len(hits)} {tcode.replace('_', ' ').title()} "
                         f"question(s): {', '.join(hits[:8])}"
                         + (" …" if len(hits) > 8 else "") + ".",
                    intent=INTENT_SURVEY, was_grounded=True,
                )
            return BotReply(
                text=f"No — there is no {tcode.replace('_', ' ').title()} question in this survey.",
                intent=INTENT_SURVEY, was_grounded=True,
            )

    if "how many" in low and "respondent" in low:
        return BotReply(text=f"{facts['total_respondents']:,} respondents.",
                        intent=INTENT_SURVEY, was_grounded=True)
    if "how many" in low and "question" in low:
        return BotReply(text=f"{facts['total_questions']} questions.",
                        intent=INTENT_SURVEY, was_grounded=True)

    # Anything else: let the LLM phrase, but feed it ONLY the deterministic facts.
    system = (
        "Answer the user's question about a survey using ONLY the FACTS JSON. "
        "Do not infer questions or types that aren't listed. If the facts don't "
        "contain the answer, say so. Reply JSON: {\"answer\": \"...\"}."
    )
    obj = _llm_json(system, f"FACTS:\n{json.dumps(facts)}\n\nQUESTION:\n{message}",
                    temperature=PHRASING_TEMPERATURE)
    if obj and obj.get("answer"):
        return BotReply(text=str(obj["answer"]), intent=INTENT_SURVEY, was_grounded=True)
    # Deterministic fallback summary.
    type_summary = ", ".join(f"{len(v)} {k}" for k, v in facts["types"].items())
    return BotReply(
        text=f"This survey has {facts['total_questions']} questions "
             f"({type_summary}) across {facts['total_respondents']:,} respondents.",
        intent=INTENT_SURVEY, was_grounded=True,
    )


# ---------------------------------------------------------------------------
# Handler B — hypothesis validator  (LLM maps; PYTHON computes; LLM describes)
# ---------------------------------------------------------------------------

_MAP_SYSTEM = (
    "Map a user's hypothesis to the TWO survey questions it compares. "
    "You are given the list of available questions (id + text + type). "
    "Choose the two question IDs whose cross-tabulation would test the claim. "
    "Reply JSON only: {\"question_a\": \"<id>\", \"question_b\": \"<id>\", "
    "\"confidence\": 0.0-1.0, \"reason\": \"...\"}. "
    "If no pair fits, set both ids to empty strings."
)


def _question_catalog(schema: Any) -> list[dict[str, str]]:
    out = []
    for q in getattr(schema, "questions", []) or []:
        out.append({
            "id": getattr(q, "canonical_id", ""),
            "text": (getattr(q, "question_text", "") or "")[:140],
            "type": getattr(getattr(q, "question_type", None), "value",
                            str(getattr(q, "question_type", ""))),
        })
    return out


def _map_hypothesis_to_questions(message: str, schema: Any) -> tuple[str, str, float, str]:
    catalog = _question_catalog(schema)
    obj = _llm_json(
        _MAP_SYSTEM,
        f"QUESTIONS:\n{json.dumps(catalog)}\n\nHYPOTHESIS:\n{message}",
        temperature=MAPPING_TEMPERATURE,
    )
    if not obj:
        return "", "", 0.0, "mapping unavailable"
    valid_ids = {c["id"] for c in catalog}
    a = str(obj.get("question_a", "")) if obj.get("question_a") in valid_ids else ""
    b = str(obj.get("question_b", "")) if obj.get("question_b") in valid_ids else ""
    try:
        conf = float(obj.get("confidence", 0.0))
    except Exception:
        conf = 0.0
    return a, b, conf, str(obj.get("reason", ""))


def _pick_analysis_type(schema: Any, a: str, b: str) -> tuple[Any, tuple[str, str]]:
    """Deterministic: choose CROSS_TAB vs GROUP_COMPARISON from question types.

    Mirrors the logic already in app.py's manual cross-cut builder so the bot
    runs the SAME engine the same way.
    """
    from src.models import AnalysisType, QuestionType

    numeric = {
        QuestionType.DIRECT_NUMERIC, QuestionType.NUMERIC_ALLOCATION,
        QuestionType.NPS, QuestionType.MULTI_SELECT_BINARY, QuestionType.RANK_ORDER,
    }
    categorical = {
        QuestionType.SINGLE_SELECT, QuestionType.DEMOGRAPHIC_OR_SEGMENT,
        QuestionType.GRID_SINGLE_SELECT, QuestionType.NPS,
    }
    sa = schema.get_question(a)
    sb = schema.get_question(b)
    ta, tb = sa.question_type, sb.question_type
    if (ta in numeric) ^ (tb in numeric):
        if ta in numeric and tb in categorical:
            return AnalysisType.GROUP_COMPARISON, (b, a)   # (segment, metric)
        if tb in numeric and ta in categorical:
            return AnalysisType.GROUP_COMPARISON, (a, b)
    if ta in categorical and tb in categorical:
        return AnalysisType.CROSS_TAB, (a, b)
    return AnalysisType.CROSS_TAB, (a, b)


def _min_cell_n(result: Any) -> int:
    """Smallest cell/segment base in the result table (deterministic)."""
    rt = getattr(result, "result_table", {}) or {}
    mins: list[int] = []
    counts = rt.get("counts") or {}
    for row in counts.values():
        if isinstance(row, dict):
            mins.extend(int(v or 0) for v in row.values())
    per_seg = rt.get("per_segment") or {}
    for seg in per_seg.values():
        if isinstance(seg, dict):
            mins.append(int(seg.get("n", 0) or 0))
    return min(mins) if mins else 0


def _result_to_table(result: Any, schema: Any) -> tuple[list[dict[str, Any]], str]:
    """Render the engine result as a list-of-dict table + caption. No LLM."""
    from src.models import AnalysisType
    rt = getattr(result, "result_table", {}) or {}
    at = getattr(result, "analysis_type", None)
    rows: list[dict[str, Any]] = []
    if at is AnalysisType.CROSS_TAB and "counts" in rt:
        rlab = rt.get("row_label_map", {}) or {}
        clab = rt.get("column_label_map", {}) or {}
        rpct = rt.get("row_pct", {}) or {}
        for rc, cols in (rt.get("counts") or {}).items():
            if not isinstance(cols, dict):
                continue
            row = {"": rlab.get(rc, str(rc))}
            for cc, n in cols.items():
                pct = (rpct.get(rc, {}) or {}).get(cc, 0.0)
                row[str(clab.get(cc, cc))] = f"{int(n):,} ({pct*100:.0f}%)"
            rows.append(row)
        return rows, "Counts (row %). Rows × columns of the two questions."
    if at is AnalysisType.GROUP_COMPARISON and "per_segment" in rt:
        for sv, sd in (rt.get("per_segment") or {}).items():
            if not isinstance(sd, dict):
                continue
            rows.append({
                "Segment": sd.get("label", str(sv)),
                "N": sd.get("n", 0),
                "Mean": round(float(sd.get("mean", 0) or 0), 3),
            })
        return rows, "Mean of the numeric metric by segment."
    return rows, "Result table."


_DESCRIBE_SYSTEM = (
    "Describe what the TABLE shows, in 2-3 sentences, in response to the user's "
    "hypothesis. Rules: (1) Use ONLY numbers that appear in the TABLE. "
    "(2) Never claim causation; say 'associated with' / 'tends to' at most. "
    "(3) If told the base is small, hedge explicitly ('early signal, small base'). "
    "(4) Do not state a growth %, allocation %, or forecast that isn't a cell in "
    "the table. Reply JSON: {\"verdict\": \"supported|mixed|not supported|inconclusive\", "
    "\"explanation\": \"...\"}."
)


def validate_hypothesis(message: str, schema: Any, active_df: Any, log: Any) -> BotReply:
    """Full 4b flow: map -> compute (deterministic) -> describe (validated)."""
    if schema is None or active_df is None:
        return BotReply(text="Run an analysis first — I need the loaded survey data "
                             "to test a hypothesis.", intent=INTENT_HYPOTHESIS, was_grounded=True)

    a, b, conf, reason = _map_hypothesis_to_questions(message, schema)
    if not a or not b:
        return BotReply(
            text="I couldn't confidently match your hypothesis to two questions in "
                 "this survey. Try naming the two measures, e.g. 'efficiency gain' "
                 "and 'revenue growth'.",
            intent=INTENT_HYPOTHESIS, was_grounded=True,
            debug={"map_reason": reason},
        )

    # ---- Deterministic compute via the SAME engine app.py uses ----
    try:
        from src.cross_cut_engine import compute_cross_cuts
        from src.models import CrossCutSpec
        analysis_type, source_ids = _pick_analysis_type(schema, a, b)
        spec = CrossCutSpec(
            cross_cut_id=f"BOT_{analysis_type.value}_{source_ids[0]}_{source_ids[1]}",
            title=f"{source_ids[0]} x {source_ids[1]}",
            analysis_type=analysis_type,
            source_question_ids=source_ids,
        )
        results, _skips = compute_cross_cuts([spec], schema, active_df, log)
    except Exception as exc:  # never surface a raw trace to the user
        return BotReply(text=f"I mapped your hypothesis to {a} × {b} but couldn't "
                             f"compute the cross-cut ({type(exc).__name__}). Try the "
                             "Cross cuts screen to build it manually.",
                        intent=INTENT_HYPOTHESIS, was_grounded=False)
    if not results:
        return BotReply(text=f"I mapped your hypothesis to {a} × {b} but the engine "
                             "returned no result (often a base-size or type issue).",
                        intent=INTENT_HYPOTHESIS, was_grounded=False)

    result = results[0]
    table, caption = _result_to_table(result, schema)
    min_n = _min_cell_n(result)
    thin = min_n < MIN_CELL_N_FOR_CONFIDENT_CLAIM
    caveats: list[str] = []
    if thin:
        caveats.append(
            f"Smallest cell base is {min_n} (< {MIN_CELL_N_FOR_CONFIDENT_CLAIM}); "
            "treat this as an early signal, not a firm finding."
        )

    # ---- LLM describes the table Python produced; output is re-validated ----
    describe = _llm_json(
        _DESCRIBE_SYSTEM,
        f"HYPOTHESIS:\n{message}\n\nMAPPED TO: {a} × {b}\n"
        f"SMALL_BASE: {thin}\n\nTABLE:\n{json.dumps(table)}",
        temperature=PHRASING_TEMPERATURE,
    )
    grounded = True
    if describe and describe.get("explanation"):
        explanation = str(describe["explanation"])
        verdict = str(describe.get("verdict", "inconclusive"))
        if not _numbers_are_in_table(explanation, table) or not _framing_is_safe(explanation):
            explanation = _fallback_description(a, b, thin)
            verdict = "inconclusive"
            grounded = False
    else:
        explanation = _fallback_description(a, b, thin)
        verdict = "inconclusive"
        grounded = False

    header = {"supported": "Supported", "mixed": "Mixed evidence",
              "not supported": "Not supported", "inconclusive": "Inconclusive"}.get(
                  verdict.lower(), "Inconclusive")
    if thin and verdict.lower() == "supported":
        header = "Supported (tentatively — small base)"

    return BotReply(
        text=f"**{header}.** {explanation}",
        intent=INTENT_HYPOTHESIS,
        table=table,
        table_caption=f"{a} × {b} — {caption}",
        caveats=caveats,
        was_grounded=grounded,
        debug={"map_confidence": conf, "map_reason": reason, "min_cell_n": min_n},
    )


# ---------------------------------------------------------------------------
# Validation guards (mirror app.py's outcome-diff validators)
# ---------------------------------------------------------------------------

def _table_number_tokens(table: list[dict[str, Any]]) -> set[float]:
    nums: set[float] = set()
    for row in table:
        for v in row.values():
            for m in re.finditer(r"(\d+(?:\.\d+)?)", str(v)):
                try:
                    nums.add(float(m.group(1)))
                except ValueError:
                    pass
    return nums


def _numbers_are_in_table(text: str, table: list[dict[str, Any]]) -> bool:
    allowed = _table_number_tokens(table)
    found = [float(m.group(1)) for m in re.finditer(r"(\d+(?:\.\d+)?)\s*%?", text)]
    if not found:
        return True  # purely qualitative explanation is fine
    return all(any(abs(f - a) <= 1.0 for a in allowed) for f in found)


def _framing_is_safe(text: str) -> bool:
    unsafe = (
        r"\bcaus(e|es|ed|ing|ation)\b",
        r"\bproves?\b",
        r"\bguarantee",
        r"\d+(?:\.\d+)?\s*%\s*(growth|increase|roi|return)",
    )
    return not any(re.search(p, text, re.IGNORECASE) for p in unsafe)


def _fallback_description(a: str, b: str, thin: bool) -> str:
    base = (f"The cross-cut of {a} and {b} is shown below; read the cells directly "
            "to judge the relationship.")
    if thin:
        base += " The base size is small, so any pattern here is only a tentative signal."
    return base


# ---------------------------------------------------------------------------
# Public entrypoint
# ---------------------------------------------------------------------------

def handle_message(
    message: str,
    *,
    schema: Any = None,
    active_df: Any = None,
    log: Any = None,
    faq_ground_truth: str = "",
) -> BotReply:
    """Route a user message to the correct handler and return a BotReply."""
    intent = classify_intent(message)
    if intent == INTENT_TOOL:
        return answer_tool_help(message, faq_ground_truth)
    if intent == INTENT_SURVEY:
        return answer_survey_structure(message, schema)
    if intent == INTENT_HYPOTHESIS:
        return validate_hypothesis(message, schema, active_df, log)
    return BotReply(
        text=("I can help with three things: how the tool works, questions about "
              "your uploaded survey, and testing a hypothesis against the data. "
              "Which would you like?"),
        intent=INTENT_UNKNOWN, was_grounded=True,
    )
