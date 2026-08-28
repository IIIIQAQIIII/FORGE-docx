"""Guided Format Builder sessions persisted in FORGE_HOME/drafts."""

from __future__ import annotations

import json
import uuid
from pathlib import Path
from typing import Any, Optional

from format_model import FormatProfile
from intelligence.mappings import CONTENT_PROFILE_RECOMMENDATIONS
from open_format.config import GUIDED_FIELDS, MAX_GUIDED_QUESTIONS_PER_ROUND
from open_format.home import drafts_dir
from open_format.normalizer import validate_rules
from open_format.profile_store import save_user_profile
from profiles import registry as profile_registry


def _session_path(session_id: str) -> Path:
    return drafts_dir() / f"{session_id}.json"


def _load_session(session_id: str) -> Optional[dict]:
    path = _session_path(session_id)
    if not path.is_file():
        return None
    return json.loads(path.read_text(encoding="utf-8"))


def _save_session(session: dict) -> None:
    path = _session_path(session["session_id"])
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(session, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)


def _current_inherited_value(base_profile_id: str, field: str) -> Any:
    try:
        profile: FormatProfile = profile_registry.resolve_profile(base_profile_id)
        value = getattr(profile, field, None)
        return value if value else {}
    except KeyError:
        return {}


def create_guided_session(
    profile_id: str,
    name: Optional[str] = None,
    description: Optional[str] = None,
    base_profile_id: Optional[str] = None,
    intent: Optional[str] = None,
) -> dict:
    if profile_registry.is_builtin(profile_id):
        return {"status": "error", "error": "PROFILE_ID_CONFLICT", "reason": f"built-in profile 不允许同名覆盖: {profile_id}"}
    if _session_path(f"session") if False else None:
        pass
    # Only one live draft per profile_id: refuse duplicates unless it is an old session.
    for path in drafts_dir().glob("*.json"):
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            continue
        if data.get("profile_id") == profile_id and data.get("status") in ("needs_guidance", "ready"):
            return {"status": "error", "error": "PROFILE_ID_CONFLICT", "reason": f"已有进行中的 guided session: {data.get('session_id')}"}

    if base_profile_id is None:
        base_profile_id = "generic_document"
        if intent:
            # Use intelligence recommendation only as a starting hint; the
            # final base choice never overrides an explicit user base_profile_id.
            from intelligence.classifier import classify_content

            classification = classify_content(intent)
            recommended = CONTENT_PROFILE_RECOMMENDATIONS.get(classification.get("intent", ""), "")
            if recommended:
                base_profile_id = recommended

    try:
        profile_registry.resolve_profile(base_profile_id)
    except KeyError:
        return {"status": "error", "error": "PROFILE_NOT_FOUND", "reason": f"base profile 不存在: {base_profile_id}"}

    session = {
        "session_id": uuid.uuid4().hex[:12],
        "profile_id": profile_id,
        "name": name or profile_id,
        "description": description or "",
        "base_profile_id": base_profile_id,
        "overrides": {},
        "resolved_fields": [],
        "pending_fields": list(GUIDED_FIELDS),
        "status": "needs_guidance",
    }
    _save_session(session)
    return _session_payload(session)


def _session_payload(session: dict) -> dict:
    base = session["base_profile_id"]
    questions = []
    for field in session["pending_fields"][:MAX_GUIDED_QUESTIONS_PER_ROUND]:
        questions.append(
            {
                "field": field,
                "question": f"是否自定义 {field}？不自定义则继承 {base} 的默认值。",
                "current_inherited": _current_inherited_value(base, field),
                "accept_inherit": True,
            }
        )
    draft_profile = {
        "schema_version": 1,
        "profile_id": session["profile_id"],
        "name": session["name"],
        "description": session.get("description", ""),
        "source": "guided",
        "inherits": session["base_profile_id"],
        "rules": session.get("overrides", {}),
    }
    return {
        "session_id": session["session_id"],
        "status": session["status"],
        "profile_id": session["profile_id"],
        "base_profile_id": session["base_profile_id"],
        "draft_profile": draft_profile,
        "questions": questions,
        "pending_fields": list(session["pending_fields"]),
        "resolved_fields": list(session["resolved_fields"]),
    }


def update_guided_session(session_id: str, answers: list[dict]) -> dict:
    session = _load_session(session_id)
    if session is None:
        return {"status": "error", "error": "SESSION_NOT_FOUND", "reason": f"session 不存在: {session_id}"}
    if session["status"] not in ("needs_guidance", "ready"):
        return {"status": "error", "error": "SESSION_CLOSED", "reason": f"session 已结束: {session_id}"}

    errors = []
    for answer in answers or []:
        field = answer.get("field")
        if field not in GUIDED_FIELDS:
            errors.append(f"未知 guided field: {field}")
            continue
        if field not in session["pending_fields"]:
            continue  # already resolved
        if answer.get("inherit"):
            session["resolved_fields"].append(field)
            session["pending_fields"].remove(field)
            continue
        value = answer.get("value")
        if not isinstance(value, dict):
            errors.append(f"{field} 需要提供 value 对象或 inherit=true")
            continue
        result = validate_rules({field: value})
        if result["errors"]:
            errors.extend(result["errors"])
            continue
        session["overrides"][field] = result["rules"].get(field, {})
        session["resolved_fields"].append(field)
        session["pending_fields"].remove(field)

    if errors:
        return {"status": "error", "error": "INVALID_ANSWERS", "reason": errors, "session_id": session_id}

    session["status"] = "ready" if not session["pending_fields"] else "needs_guidance"
    _save_session(session)

    if session["status"] == "ready":
        saved = _finalize_guided_session(session)
        session["status"] = "saved"
        _save_session(session)
        return {"status": "saved", "session_id": session_id, "profile": saved}

    return _session_payload(session)


def _finalize_guided_session(session: dict) -> dict:
    save_result = save_user_profile(
        profile_id=session["profile_id"],
        name=session["name"],
        description=session.get("description", ""),
        source="guided",
        inherits=session["base_profile_id"],
        rules=session.get("overrides", {}),
    )
    if save_result.get("status") == "error":
        return save_result
    return save_result.get("profile", {})


def finalize_guided_session(session_id: str) -> dict:
    session = _load_session(session_id)
    if session is None:
        return {"status": "error", "error": "SESSION_NOT_FOUND", "reason": f"session 不存在: {session_id}"}
    if session["status"] != "ready":
        return {"status": "error", "error": "SESSION_NOT_READY", "reason": "还有 pending fields 未处理"}
    saved = _finalize_guided_session(session)
    if isinstance(saved, dict) and saved.get("status") == "error":
        return saved
    session["status"] = "saved"
    _save_session(session)
    return {"status": "saved", "session_id": session_id, "profile": saved}
