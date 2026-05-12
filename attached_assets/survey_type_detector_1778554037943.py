"""Deterministic survey type detection from parsed survey metadata."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import pandas as pd

from src.datamap_parser import DataMap
from src.models import OutcomeVariableOption, SurveyTypeResult


SURVEY_TYPES = [
    "purchase_ltb",
    "rfp_evaluation",
    "supplier_assessment",
    "nps",
    "csat",
    "ces",
    "churn_risk",
    "growth_strategy",
    "digital_maturity",
    "innovation_readiness",
    "operational_efficiency",
    "employee_engagement",
    "pulse_survey",
    "exit_interview",
    "dei_assessment",
    "market_research",
    "brand_perception",
    "competitive_intelligence",
    "risk_assessment",
    "compliance_audit",
    "opinion",
    "unknown",
]


@dataclass(frozen=True, slots=True)
class _Rule:
    survey_type: str
    keywords: tuple[str, ...]
    threshold: int
    confidence: float
    strong_confidence: float | None = None
    strong_threshold: int | None = None


_RULES = (
    _Rule(
        "purchase_ltb",
        (
            "consideration",
            "shortlist",
            "selected vendor",
            "winner",
            "chose",
            "purchased from",
            "decided on",
            "final selection",
            "preferred supplier",
        ),
        2,
        0.75,
        0.9,
        3,
    ),
    _Rule(
        "rfp_evaluation",
        (
            "rfp",
            "proposal",
            "bid",
            "scoring criteria",
            "evaluation criteria",
            "vendor score",
            "weighted score",
        ),
        2,
        0.85,
    ),
    _Rule(
        "supplier_assessment",
        (
            "supplier performance",
            "vendor performance",
            "delivery quality",
            "sla",
            "service level",
            "supplier rating",
        ),
        2,
        0.8,
    ),
    _Rule(
        "csat",
        ("satisfied", "satisfaction", "meets expectations", "rate your experience"),
        2,
        0.85,
    ),
    _Rule(
        "churn_risk",
        (
            "likely to renew",
            "churn",
            "cancel",
            "retention",
            "continue using",
            "switch to competitor",
        ),
        2,
        0.85,
    ),
    _Rule(
        "growth_strategy",
        (
            "revenue growth",
            "growth rate",
            "growth target",
            "market share",
            "performance vs target",
            "outperform",
            "achieved growth",
            "sales growth",
        ),
        2,
        0.7,
        0.85,
        3,
    ),
    _Rule(
        "digital_maturity",
        (
            "digital transformation",
            "ai adoption",
            "technology maturity",
            "cloud migration",
            "automation",
            "digitalization",
            "tech stack",
            "digital capability",
        ),
        3,
        0.8,
    ),
    _Rule(
        "innovation_readiness",
        (
            "innovation",
            "r&d",
            "new product",
            "experimentation",
            "test and learn",
            "innovation pipeline",
            "time to market",
        ),
        3,
        0.75,
    ),
    _Rule(
        "operational_efficiency",
        (
            "cost reduction",
            "efficiency",
            "productivity",
            "process optimization",
            "waste",
            "cycle time",
            "throughput",
            "utilization",
        ),
        3,
        0.8,
    ),
    _Rule(
        "employee_engagement",
        (
            "employee satisfaction",
            "engagement",
            "enps",
            "would you recommend this company as a place to work",
            "proud to work",
            "motivated",
        ),
        2,
        0.85,
    ),
    _Rule(
        "exit_interview",
        (
            "reason for leaving",
            "exit",
            "departure",
            "resignation",
            "why are you leaving",
            "offboarding",
        ),
        2,
        0.9,
    ),
    _Rule(
        "dei_assessment",
        (
            "diversity",
            "inclusion",
            "equity",
            "belonging",
            "dei",
            "psychological safety",
            "bias",
            "discrimination",
        ),
        3,
        0.8,
    ),
    _Rule(
        "market_research",
        (
            "market size",
            "purchase intent",
            "willingness to pay",
            "unmet need",
            "buying criteria",
            "target market",
        ),
        3,
        0.75,
    ),
    _Rule(
        "brand_perception",
        (
            "brand awareness",
            "brand perception",
            "top of mind",
            "consider your brand",
            "brand attributes",
            "brand strength",
        ),
        2,
        0.8,
    ),
    _Rule(
        "competitive_intelligence",
        (
            "competitor",
            "competitive advantage",
            "market leader",
            "strengths and weaknesses",
            "win against",
            "lose to",
        ),
        3,
        0.75,
    ),
    _Rule(
        "compliance_audit",
        (
            "compliance",
            "policy adherence",
            "controls",
            "audit",
            "regulatory",
            "violation",
            "non-compliance",
        ),
        3,
        0.85,
    ),
)

_GROWTH_AGENDA_SHORT_CIRCUIT_PHRASES = (
    "revenue growth",
    "revenue in last fiscal year",
    "revenue last year",
    "annual revenue growth",
    "fiscal year revenue",
    "revenue grew",
    "revenue change",
    "how much did your organization's revenue",
    "your organization's planned revenue",
    "actual revenue",
    "expected revenue",
)

_PRIMARY_SIGNALS = {
    "purchase_ltb": {
        "selected vendor": 0.95,
        "winner": 0.95,
        "chose": 0.95,
        "final selection": 0.95,
    },
    "rfp_evaluation": {
        "final ranking": 0.95,
        "recommended vendor": 0.95,
        "total score": 0.95,
    },
    "supplier_assessment": {
        "overall performance": 0.95,
        "continue relationship": 0.95,
    },
    "nps": {"recommend": 0.95, "net promoter": 0.95},
    "csat": {"overall satisfaction": 0.95, "satisfied overall": 0.95},
    "ces": {"customer effort": 0.95, "ease of": 0.95},
    "churn_risk": {"likely to renew": 0.95, "cancel": 0.95, "churn": 0.95},
    "growth_strategy": {
        "revenue growth": 0.98,
        "revenue grew": 0.98,
        "revenue change": 0.95,
        "revenue in last fiscal year": 0.98,
        "how much did your organization's revenue": 0.98,
        "sales growth": 0.95,
        "achieved growth": 0.95,
    },
    "digital_maturity": {
        "maturity level": 0.95,
        "digital readiness": 0.95,
        "adoption stage": 0.95,
    },
    "innovation_readiness": {"innovation success": 0.95, "pipeline strength": 0.95},
    "operational_efficiency": {"cost savings": 0.95, "efficiency gains": 0.95},
    "employee_engagement": {
        "enps": 0.95,
        "engagement score": 0.95,
        "would recommend": 0.95,
    },
    "exit_interview": {"reason for leaving": 0.95, "primary reason": 0.95},
    "dei_assessment": {"inclusion score": 0.95, "belonging": 0.95},
    "market_research": {"purchase intent": 0.95, "willingness to pay": 0.95},
    "brand_perception": {"brand awareness": 0.95, "top of mind": 0.95},
    "competitive_intelligence": {"competitive position": 0.95, "win rate": 0.95},
    "risk_assessment": {"risk severity": 0.95, "aggregate risk": 0.95},
    "compliance_audit": {"compliance rate": 0.95, "control effectiveness": 0.95},
}

_SECONDARY_KEYWORDS = (
    "success",
    "satisfaction",
    "impact",
    "effectiveness",
    "result",
    "outcome",
    "performance",
    "rating",
    "score",
)

_DEMOGRAPHIC_KEYWORDS = (
    "industry",
    "region",
    "country",
    "size",
    "department",
    "role",
    "title",
    "location",
    "age",
    "tenure",
)

_MEASURABLE_TYPES = {
    "single_select",
    "direct_numeric",
    "multi_select_binary",
    "numeric_allocation",
    "grid_single_select",
}

_INELIGIBLE_TYPES = {"open_text", "metadata_or_id", "unknown"}


def detect_survey_type(schema: DataMap, decoded_df: pd.DataFrame) -> SurveyTypeResult:
    """Detect survey type and likely outcome variables using deterministic rules."""

    try:
        if decoded_df is None or len(decoded_df) == 0:
            return _unknown_result()

        questions = list(schema.get("questions", []))
        if not questions:
            return _unknown_result()

        growth_match = _detect_growth_agenda_short_circuit(questions)
        if growth_match is not None:
            outcome_id, signal = growth_match
            all_eligible = _score_all_questions(questions, "growth_strategy")
            candidates = [
                option for option in all_eligible if option.relevance_score >= 0.7
            ][:5]
            return SurveyTypeResult(
                survey_type="growth_strategy",
                outcome_question_id=outcome_id,
                confidence=0.95,
                signals=[signal],
                candidate_outcome_questions=candidates,
                all_eligible_questions=all_eligible,
            )

        survey_type, base_confidence, threshold, signals = _detect_type(questions)
        all_eligible = _score_all_questions(questions, survey_type)
        candidates = [option for option in all_eligible if option.relevance_score >= 0.7][:5]
        outcome_id = candidates[0].question_id if candidates else None

        if survey_type == "unknown":
            return _unknown_result()

        if survey_type == "opinion":
            confidence = 0.3
        else:
            confidence = min(base_confidence + max(0, len(signals) - threshold) * 0.05, 0.95)

        if outcome_id is None:
            confidence = min(confidence, 0.5)

        return SurveyTypeResult(
            survey_type=survey_type,
            outcome_question_id=outcome_id,
            confidence=confidence,
            signals=signals,
            candidate_outcome_questions=candidates,
            all_eligible_questions=all_eligible,
        )
    except Exception:
        return _unknown_result()


def _detect_growth_agenda_short_circuit(
    questions: list[dict[str, Any]],
) -> tuple[str, str] | None:
    for question in questions:
        qid = _question_id(question)
        text_lower = _question_text(question).lower()
        for phrase in _GROWTH_AGENDA_SHORT_CIRCUIT_PHRASES:
            if phrase in text_lower:
                return (
                    qid,
                    f"{qid} contains '{phrase}' — Growth Agenda short-circuit",
                )
    return None


def _detect_type(questions: list[dict[str, Any]]) -> tuple[str, float, int, list[str]]:
    for rule in _RULES[:3]:
        match = _match_rule(questions, rule)
        if match is not None:
            return match

    nps = _detect_nps(questions)
    if nps is not None:
        return nps

    csat = _detect_csat_single_overall(questions)
    if csat is not None:
        return csat

    for rule in _RULES[3:4]:
        match = _match_rule(questions, rule)
        if match is not None:
            return match

    ces = _detect_ces(questions)
    if ces is not None:
        return ces

    for rule in _RULES[4:9]:
        match = _match_rule(questions, rule)
        if match is not None:
            return match

    engagement = _detect_enps_single(questions)
    if engagement is not None:
        return engagement

    for rule in _RULES[9:10]:
        match = _match_rule(questions, rule)
        if match is not None:
            return match

    pulse = _detect_pulse(questions)
    if pulse is not None:
        return pulse

    for rule in _RULES[10:16]:
        match = _match_rule(questions, rule)
        if match is not None:
            return match

    risk = _detect_risk(questions)
    if risk is not None:
        return risk

    for rule in _RULES[16:]:
        match = _match_rule(questions, rule)
        if match is not None:
            return match

    return ("opinion", 0.3, 1, ["No strong survey-type signals detected"])


def _match_rule(
    questions: list[dict[str, Any]],
    rule: _Rule,
) -> tuple[str, float, int, list[str]] | None:
    signals = _keyword_signals(questions, rule.keywords, rule.survey_type)
    if rule.survey_type == "rfp_evaluation" and len(signals) == 1:
        if any(_question_type(q) == "numeric_allocation" for q in questions):
            signals.append("A numeric allocation question strengthens the RFP signal")

    if len(signals) < rule.threshold:
        return None

    confidence = rule.confidence
    if (
        rule.strong_confidence is not None
        and rule.strong_threshold is not None
        and len(signals) >= rule.strong_threshold
    ):
        confidence = rule.strong_confidence

    return (rule.survey_type, confidence, rule.threshold, signals)


def _detect_nps(questions: list[dict[str, Any]]) -> tuple[str, float, int, list[str]] | None:
    for question in questions:
        text = _question_text(question)
        text_lower = text.lower()
        if "net promoter" in text_lower or "nps" in text_lower:
            return (
                "nps",
                0.95,
                1,
                [_signal(question, "net promoter", "strong NPS signal")],
            )
        if "recommend" in text_lower and _has_0_10_scale(question):
            return (
                "nps",
                0.85,
                1,
                [_signal(question, "recommend", "strong NPS signal")],
            )
    return None


def _detect_ces(questions: list[dict[str, Any]]) -> tuple[str, float, int, list[str]] | None:
    for question in questions:
        text_lower = _question_text(question).lower()
        if "effort" in text_lower and _has_scale(question):
            return (
                "ces",
                0.8,
                1,
                [_signal(question, "effort", "customer effort signal")],
            )
    return None


def _detect_pulse(questions: list[dict[str, Any]]) -> tuple[str, float, int, list[str]] | None:
    if len(questions) >= 10:
        return None
    signals = _keyword_signals(
        questions,
        (
            "this week",
            "this sprint",
            "team health",
            "pulse check",
            "quick check-in",
            "how are you feeling",
        ),
        "pulse_survey",
    )
    if signals:
        return ("pulse_survey", 0.7, 1, signals)
    return None


def _detect_risk(questions: list[dict[str, Any]]) -> tuple[str, float, int, list[str]] | None:
    keywords = ("risk", "threat", "vulnerability", "likelihood", "impact", "severity", "mitigation")
    signals = _keyword_signals(questions, keywords, "risk_assessment")
    if len(signals) >= 3:
        return ("risk_assessment", 0.8, 3, signals)

    risk_scale_count = sum(
        1 for question in questions if "risk" in _question_text(question).lower() and _has_scale(question)
    )
    if risk_scale_count >= 2:
        return ("risk_assessment", 0.8, 2, signals[:2] or ["2 risk scale questions detected"])
    return None


def _detect_csat_single_overall(
    questions: list[dict[str, Any]],
) -> tuple[str, float, int, list[str]] | None:
    for question in questions:
        if "overall satisfaction" in _question_text(question).lower():
            return (
                "csat",
                0.85,
                1,
                [_signal(question, "overall satisfaction", "strong CSAT signal")],
            )
    return None


def _detect_enps_single(
    questions: list[dict[str, Any]],
) -> tuple[str, float, int, list[str]] | None:
    for question in questions:
        if "enps" in _question_text(question).lower():
            return (
                "employee_engagement",
                0.85,
                1,
                [_signal(question, "eNPS", "strong employee engagement signal")],
            )
    return None


def _keyword_signals(
    questions: list[dict[str, Any]],
    keywords: tuple[str, ...],
    survey_type: str,
) -> list[str]:
    signals: list[str] = []
    seen_questions: set[str] = set()
    for question in questions:
        qid = _question_id(question)
        text_lower = _question_text(question).lower()
        for keyword in keywords:
            if keyword.lower() in text_lower and qid not in seen_questions:
                signals.append(_signal(question, keyword, f"{survey_type.replace('_', ' ')} signal"))
                seen_questions.add(qid)
                break
    return signals


def _score_all_questions(
    questions: list[dict[str, Any]],
    survey_type: str,
) -> list[OutcomeVariableOption]:
    scored: list[OutcomeVariableOption] = []
    for question in questions:
        qid = _question_id(question)
        text = _question_text(question)
        qtype = _question_type(question)
        if qtype in _INELIGIBLE_TYPES:
            continue
        score, reason = _score_outcome_relevance(qid, text, qtype, survey_type)
        scored.append(
            OutcomeVariableOption(
                question_id=qid,
                question_text=text,
                question_type=qtype,
                relevance_score=score,
                reason=reason,
            )
        )
    return sorted(scored, key=lambda option: (-option.relevance_score, option.question_id))


def _score_outcome_relevance(
    question_id: str,
    question_text: str,
    question_type: str,
    survey_type: str,
) -> tuple[float, str]:
    text_lower = question_text.lower()

    for signal, score in _PRIMARY_SIGNALS.get(survey_type, {}).items():
        if signal in text_lower:
            return (score, f"Primary {survey_type.replace('_', ' ')} outcome")

    if any(word in text_lower for word in _SECONDARY_KEYWORDS):
        return (0.75, "Contains outcome-related term")

    if any(word in text_lower for word in _DEMOGRAPHIC_KEYWORDS):
        return (0.2, "Demographic/segment variable")

    if question_type in _MEASURABLE_TYPES:
        return (0.5, "Measurable categorical or numeric question")

    return (0.35, "Generic question")


def _question_id(question: dict[str, Any]) -> str:
    return str(question.get("canonical_id") or question.get("question_id") or question.get("raw_id") or "")


def _question_text(question: dict[str, Any]) -> str:
    return str(question.get("question_text") or "")


def _question_type(question: dict[str, Any]) -> str:
    explicit = question.get("question_type")
    if explicit:
        return str(explicit).lower()

    type_hint = question.get("type_hint")
    if type_hint == "open_text":
        return "open_text"
    if type_hint == "open_numeric":
        return "direct_numeric"
    if type_hint == "values_range":
        if question.get("sub_columns"):
            return "multi_select_binary"
        return "single_select"
    return "unknown"


def _has_scale(question: dict[str, Any]) -> bool:
    value_range = question.get("value_range")
    if isinstance(value_range, tuple) and len(value_range) == 2:
        return True
    options = question.get("options") or []
    return len(options) >= 3


def _has_0_10_scale(question: dict[str, Any]) -> bool:
    value_range = question.get("value_range")
    if isinstance(value_range, tuple) and len(value_range) == 2:
        return int(value_range[0]) == 0 and int(value_range[1]) == 10

    codes = [option[0] for option in question.get("options", []) if isinstance(option, tuple)]
    return bool(codes) and min(codes) == 0 and max(codes) == 10


def _signal(question: dict[str, Any], phrase: str, reason: str) -> str:
    qid = _question_id(question) or "(unknown question)"
    return f"{qid} contains '{phrase}' - {reason}"


def _unknown_result() -> SurveyTypeResult:
    return SurveyTypeResult(
        survey_type="unknown",
        outcome_question_id=None,
        confidence=0.0,
        signals=[],
        candidate_outcome_questions=[],
        all_eligible_questions=[],
    )
