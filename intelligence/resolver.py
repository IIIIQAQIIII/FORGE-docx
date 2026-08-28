"""Mission 02-B — Format Resolver.

回答：“这一次最终应该使用什么 Format Profile？”

Content Classifier 回答“内容是什么”，Format Resolver 回答“最终怎么排”。
二者解耦：ContentIntent 不绑定唯一 FormatProfile，只能通过
CONTENT_PROFILE_RECOMMENDATIONS 推荐 profile。

决策优先级：
    1. explicit user format
    2. reference profile
    3. saved profile
    4. content recommendation
    5. system default

即：User > Reference > Saved > Recommendation > Default
"""

from __future__ import annotations

from typing import Any, Optional

from profiles import registry as profile_registry

from intelligence.format_aliases import resolve_alias
from intelligence.mappings import CONTENT_PROFILE_RECOMMENDATIONS

ERROR_NOT_FOUND = "PROFILE_NOT_FOUND"


def _profile_exists(profile_id: str) -> bool:
    return bool(profile_id) and profile_registry.get_profile(profile_id) is not None


def _result(
    *,
    profile_id: Optional[str] = None,
    status: str,
    decision_basis: str,
    content_intent: str = "",
    confidence: float = 0.0,
    reason: Optional[list[str]] = None,
    candidates: Optional[list[dict[str, Any]]] = None,
    error: Optional[str] = None,
) -> dict[str, Any]:
    return {
        "profile_id": profile_id,
        "status": status,
        "decision_basis": decision_basis,
        "content_intent": content_intent,
        "confidence": confidence,
        "reason": reason or [],
        "candidates": candidates or [],
        "error": error,
    }


def _candidates_from_classification(classification: Optional[dict[str, Any]]) -> list[dict[str, Any]]:
    candidates: list[dict[str, Any]] = []
    if not classification:
        return candidates
    alternatives = classification.get("alternatives", []) or []
    for alt in alternatives:
        intent = alt.get("intent")
        mapped = CONTENT_PROFILE_RECOMMENDATIONS.get(intent or "", "")
        if mapped and _profile_exists(mapped):
            candidates.append(
                {
                    "profile_id": mapped,
                    "content_intent": intent,
                    "confidence": alt.get("confidence", 0.0),
                }
            )
    return candidates


def resolve_format(
    classification: Optional[dict[str, Any]] = None,
    explicit_profile_id: Optional[str] = None,
    explicit_format_hint: Optional[str] = None,
    reference_profile_id: Optional[str] = None,
    saved_profile_id: Optional[str] = None,
    default_profile_id: str = "generic_document",
    allow_default: bool = False,
) -> dict[str, Any]:
    """按 User > Reference > Saved > Recommendation > Default 解析最终 FormatProfile。"""
    content_intent = (classification or {}).get("intent", "")
    confidence = float((classification or {}).get("confidence", 0.0) or 0.0)

    # 1. explicit user format（最优先）
    if explicit_profile_id:
        if not _profile_exists(explicit_profile_id):
            return _result(
                status="error",
                decision_basis="guided_required",
                content_intent=content_intent,
                confidence=confidence,
                reason=[f"用户明确指定的 profile 不存在: {explicit_profile_id}"],
                error=ERROR_NOT_FOUND,
            )
        return _result(
            profile_id=explicit_profile_id,
            status="resolved",
            decision_basis="explicit_user_choice",
            content_intent=content_intent,
            confidence=confidence,
            reason=[f"用户明确指定 profile: {explicit_profile_id}"],
        )

    if explicit_format_hint:
        alias_profile_id = resolve_alias(explicit_format_hint)
        if not alias_profile_id or not _profile_exists(alias_profile_id):
            return _result(
                status="error",
                decision_basis="guided_required",
                content_intent=content_intent,
                confidence=confidence,
                reason=[f"无法识别的格式提示: {explicit_format_hint}"],
                error=ERROR_NOT_FOUND,
            )
        return _result(
            profile_id=alias_profile_id,
            status="resolved",
            decision_basis="explicit_user_choice",
            content_intent=content_intent,
            confidence=confidence,
            reason=[f"用户格式提示“{explicit_format_hint}”→ {alias_profile_id}"],
        )

    # 2. reference profile
    if reference_profile_id:
        if not _profile_exists(reference_profile_id):
            return _result(
                status="error",
                decision_basis="guided_required",
                content_intent=content_intent,
                confidence=confidence,
                reason=[f"参考 profile 不存在: {reference_profile_id}"],
                error=ERROR_NOT_FOUND,
            )
        return _result(
            profile_id=reference_profile_id,
            status="resolved",
            decision_basis="reference_profile",
            content_intent=content_intent,
            confidence=confidence,
            reason=[f"采用参考格式 profile: {reference_profile_id}"],
        )

    # 3. saved profile
    if saved_profile_id:
        if not _profile_exists(saved_profile_id):
            return _result(
                status="error",
                decision_basis="guided_required",
                content_intent=content_intent,
                confidence=confidence,
                reason=[f"已保存 profile 不存在: {saved_profile_id}"],
                error=ERROR_NOT_FOUND,
            )
        return _result(
            profile_id=saved_profile_id,
            status="resolved",
            decision_basis="saved_profile",
            content_intent=content_intent,
            confidence=confidence,
            reason=[f"采用用户已保存 profile: {saved_profile_id}"],
        )

    # 4. content recommendation（仅在 classification 可用时）
    if classification:
        cls_status = classification.get("status", "ambiguous")
        mapped = CONTENT_PROFILE_RECOMMENDATIONS.get(content_intent, "")

        if content_intent == "generic" or cls_status == "ambiguous":
            candidates = _candidates_from_classification(classification)
            if allow_default and _profile_exists(default_profile_id):
                return _result(
                    profile_id=default_profile_id,
                    status="resolved",
                    decision_basis="system_default",
                    content_intent=content_intent,
                    confidence=confidence,
                    reason=["内容意图模糊，调用方显式允许使用系统默认 profile。"],
                    candidates=candidates,
                )
            return _result(
                status="needs_guidance",
                decision_basis="guided_required",
                content_intent=content_intent,
                confidence=confidence,
                reason=["内容意图模糊或证据不足，且无用户指定格式，需要引导确认。"],
                candidates=candidates,
            )

        if mapped and _profile_exists(mapped):
            status = "recommended"
            reason = [
                f"内容分类为 {content_intent}（confidence={confidence}），"
                f"按推荐映射采用 profile: {mapped}。"
            ]
            return _result(
                profile_id=mapped,
                status=status,
                decision_basis="content_recommendation",
                content_intent=content_intent,
                confidence=confidence,
                reason=reason,
                candidates=_candidates_from_classification(classification),
            )

        # 有分类但没有对应推荐 profile
        return _result(
            status="needs_guidance",
            decision_basis="guided_required",
            content_intent=content_intent,
            confidence=confidence,
            reason=[f"内容分类 {content_intent} 暂无推荐 profile，需要引导确认。"],
            candidates=_candidates_from_classification(classification),
        )

    # 5. system default（仅在调用方显式允许 fallback 时使用）
    if allow_default:
        if not _profile_exists(default_profile_id):
            return _result(
                status="error",
                decision_basis="guided_required",
                content_intent=content_intent,
                confidence=confidence,
                reason=[f"系统默认 profile 不存在: {default_profile_id}"],
                error=ERROR_NOT_FOUND,
            )
        return _result(
            profile_id=default_profile_id,
            status="resolved",
            decision_basis="system_default",
            content_intent=content_intent,
            confidence=confidence,
            reason=["无分类信息，调用方显式允许使用系统默认 profile。"],
        )

    return _result(
        status="needs_guidance",
        decision_basis="guided_required",
        content_intent=content_intent,
        confidence=confidence,
        reason=["无分类信息且未指定任何格式，需要引导确认。"],
    )
