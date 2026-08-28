"""Mission 02-A — Content Classifier.

只回答“这份内容是什么”，不回答“最终应该使用什么格式”。

classify_content(description, extracted_structure=None)

- extracted_structure 当前仅预留接口，不得因此引入 DOCX Inspector。
- 判断必须可解释：所有 confidence 来自信号评分，不允许单关键词直接 0.9+。
- 证据不足时 intent = generic，不得默认传统公文。
"""

from __future__ import annotations

import math
from typing import Any, Optional

from intelligence.signals import SIGNALS

# 证据不足的分数下限：低于该分数视为 generic
EVIDENCE_FLOOR = 1.0

# confidence 的压缩系数：score -> confidence = 1 - exp(-score / K)
CONFIDENCE_K = 4.0

CONFIRMED_THRESHOLD = 0.85
RECOMMENDED_THRESHOLD = 0.65


def _confidence(score: float) -> float:
    if score <= 0:
        return 0.0
    value = 1 - math.exp(-score / CONFIDENCE_K)
    return round(max(0.0, min(1.0, value)), 4)


def _score_intent(text: str, intent: str) -> tuple[float, list[str]]:
    score = 0.0
    matched: list[str] = []
    for signal in SIGNALS.get(intent, []):
        if any(keyword in text for keyword in signal.keywords):
            score += signal.weight
            matched.append(signal.name)
    return round(score, 4), matched


def classify_content(
    description: str,
    extracted_structure: Optional[dict[str, Any]] = None,
) -> dict[str, Any]:
    """Classify content intent from a natural-language description.

    extracted_structure: reserved for future DOCX Inspector output; ignored here.
    """
    text = description or ""
    scored: list[tuple[str, float, list[str]]] = []
    for intent in SIGNALS:
        score, matched = _score_intent(text, intent)
        scored.append((intent, score, matched))

    scored.sort(key=lambda item: (-item[1], item[0]))
    top_intent, top_score, top_signals = scored[0]

    if top_score < EVIDENCE_FLOOR:
        # 证据不足：generic，不默认传统公文
        return {
            "intent": "generic",
            "confidence": round(_confidence(top_score), 4),
            "status": "ambiguous",
            "signals": [],
            "alternatives": [
                {"intent": top_intent, "confidence": round(_confidence(top_score), 4), "signals": top_signals},
                {
                    "intent": scored[1][0],
                    "confidence": round(_confidence(scored[1][1]), 4),
                    "signals": scored[1][2],
                },
            ],
        }

    confidence = _confidence(top_score)
    if confidence >= CONFIRMED_THRESHOLD:
        status = "confirmed"
    elif confidence >= RECOMMENDED_THRESHOLD:
        status = "recommended"
    else:
        status = "ambiguous"

    alternatives = []
    for intent, score, matched in scored[1:4]:
        if score <= 0:
            continue
        alternatives.append(
            {"intent": intent, "confidence": _confidence(score), "signals": matched}
        )

    return {
        "intent": top_intent,
        "confidence": confidence,
        "status": status,
        "signals": top_signals,
        "alternatives": alternatives,
    }
