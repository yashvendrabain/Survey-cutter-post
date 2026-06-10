# ===========================================================================
# Round 3 — AI winner/laggard configuration recommender.
# APPEND THIS BLOCK to src/ai_insights.py (the only file that touches Portkey).
# Reuses the module's existing imports: os, json, OpenAI, PORTKEY_BASE_URL,
# PORTKEY_PREMIUM_MODEL. AI recommends the CONFIG + reasoning only; Python
# computes the actual cohort list. No numbers are invented.
# ===========================================================================

_WINNER_CONFIG_SYSTEM = (
    "You are a survey-analytics assistant helping define a balanced "
    "winner/laggard segmentation. You are given candidate metric columns "
    "(numeric or ordinal-band performance measures) and candidate dimension "
    "columns (categorical, e.g. sector/region). Recommend a configuration:\n"
    "- pick the metrics that best define genuine performance (use ONLY column "
    "names from metric_candidates); a respondent should be a winner only if "
    "strong across the chosen metrics, not one alone.\n"
    "- for each metric give direction ('higher_is_better' or 'lower_is_better') "
    "and a weight between 0.5 and 2.0.\n"
    "- give cutoff_pct (integer 5-50): winners = top X%, laggards = bottom X%.\n"
    "- pick balance_dimensions (subset of dimension_candidates) to keep winners "
    "representative; optionally set stratify_dimension to one of them to force a "
    "proportional pick.\n"
    "- write 'reasoning': 2-4 sentences explaining the choice. Reason about the "
    "columns and their meaning. DO NOT invent any statistics or numbers; the "
    "actual counts are computed afterward in Python.\n"
    "Respond with STRICT JSON ONLY (no markdown, no prose outside JSON):\n"
    '{"metrics":[{"column":"...","direction":"higher_is_better","weight":1.0}],'
    '"cutoff_pct":25,"balance_dimensions":["..."],"stratify_dimension":null,'
    '"reasoning":"..."}'
)


def recommend_winner_config(metric_candidates, dimension_candidates, model=None):
    """Return a recommended winner-scoring config + reasoning.

    metric_candidates: list of {"column","kind"('numeric'|'band'),"detail"}
    dimension_candidates: list of {"column","categories":[...]}
    Returns dict: metrics[{column,direction,weight}], cutoff_pct,
    balance_dimensions, stratify_dimension, reasoning, was_template(bool).
    """
    api_key = os.environ.get("PORTKEY_API_KEY", "").strip()
    if not api_key or OpenAI is None:
        return _winner_config_fallback(
            metric_candidates, dimension_candidates,
            "AI unavailable — applied a balanced heuristic (equal weights, top/bottom 25%).")
    payload = json.dumps({
        "metric_candidates": metric_candidates,
        "dimension_candidates": dimension_candidates,
    })
    try:
        client = OpenAI(api_key=api_key, base_url=PORTKEY_BASE_URL)
        response = client.chat.completions.create(
            model=model or PORTKEY_PREMIUM_MODEL,
            messages=[
                {"role": "system", "content": _WINNER_CONFIG_SYSTEM},
                {"role": "user", "content": payload},
            ],
            temperature=0.1,
        )
        raw = response.choices[0].message.content or ""
        return _parse_winner_config(raw, metric_candidates, dimension_candidates)
    except Exception as exc:  # noqa: BLE001
        return _winner_config_fallback(
            metric_candidates, dimension_candidates,
            f"AI error ({exc}); applied a balanced heuristic.")


def _parse_winner_config(raw, metric_candidates, dimension_candidates):
    valid_metrics = {c["column"] for c in metric_candidates}
    valid_dims = {c["column"] for c in dimension_candidates}
    text = (raw or "").strip()
    if "```" in text:
        text = text.split("```")[1] if text.count("```") >= 2 else text
        text = text.replace("json", "", 1).strip() if text.lower().startswith("json") else text
    try:
        data = json.loads(text[text.find("{"): text.rfind("}") + 1])
    except Exception as exc:  # noqa: BLE001
        return _winner_config_fallback(
            metric_candidates, dimension_candidates,
            f"Could not parse AI response ({exc}); applied a balanced heuristic.")

    metrics = []
    for m in data.get("metrics", []):
        col = m.get("column")
        if col not in valid_metrics:
            continue
        direction = m.get("direction", "higher_is_better")
        if direction not in ("higher_is_better", "lower_is_better"):
            direction = "higher_is_better"
        try:
            weight = float(m.get("weight", 1.0))
        except (TypeError, ValueError):
            weight = 1.0
        weight = min(2.0, max(0.5, weight))
        metrics.append({"column": col, "direction": direction, "weight": weight})
    if not metrics:  # AI returned nothing usable -> heuristic
        return _winner_config_fallback(
            metric_candidates, dimension_candidates,
            "AI returned no usable metrics; applied a balanced heuristic.")

    try:
        cutoff = int(data.get("cutoff_pct", 25))
    except (TypeError, ValueError):
        cutoff = 25
    cutoff = min(50, max(5, cutoff))
    dims = [d for d in data.get("balance_dimensions", []) if d in valid_dims]
    strat = data.get("stratify_dimension")
    if strat not in valid_dims:
        strat = None
    reasoning = str(data.get("reasoning", "")).strip() or "AI recommended this configuration."
    return {
        "metrics": metrics,
        "cutoff_pct": cutoff,
        "balance_dimensions": dims,
        "stratify_dimension": strat,
        "reasoning": reasoning,
        "was_template": False,
    }


def _winner_config_fallback(metric_candidates, dimension_candidates, reason):
    metrics = [{"column": c["column"], "direction": "higher_is_better", "weight": 1.0}
               for c in metric_candidates]
    dims = [c["column"] for c in dimension_candidates]
    return {
        "metrics": metrics,
        "cutoff_pct": 25,
        "balance_dimensions": dims,
        "stratify_dimension": None,
        "reasoning": reason,
        "was_template": True,
    }
