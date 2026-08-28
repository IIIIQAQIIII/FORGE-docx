"""User Profile persistence in FORGE_HOME/profiles.

Versioned JSON, atomic saves, built-in protection and in-use checks.
"""

from __future__ import annotations

import json
import os
import tempfile
from pathlib import Path
from typing import Any, Optional

from format_model import FormatMode, FormatProfile, FormatSource
from open_format.home import profiles_dir
from open_format.normalizer import validate_profile_id, validate_rules, validate_user_profile
from profiles import registry as profile_registry

SCHEMA_VERSION = 1

ERROR_BUILTIN_PROTECTED = "BUILTIN_PROFILE_PROTECTED"
ERROR_PROFILE_ID_CONFLICT = "PROFILE_ID_CONFLICT"
ERROR_PROFILE_IN_USE = "PROFILE_IN_USE"
ERROR_PROFILE_NOT_FOUND = "PROFILE_NOT_FOUND"
ERROR_INVALID_PROFILE = "INVALID_PROFILE"


def _profile_path(profile_id: str) -> Path:
    return profiles_dir() / f"{profile_id}.json"


def _atomic_write(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(prefix=path.name + ".", suffix=".tmp", dir=str(path.parent))
    try:
        with os.fdopen(fd, "wb") as fh:
            fh.write(data)
            fh.flush()
            os.fsync(fh.fileno())
        os.replace(tmp_name, path)
    except BaseException:
        try:
            os.unlink(tmp_name)
        except OSError:
            pass
        raise


def _profile_to_json(profile_id: str, name: str, description: str, source: str, inherits: str, rules: dict) -> dict:
    return {
        "schema_version": SCHEMA_VERSION,
        "profile_id": profile_id,
        "name": name,
        "description": description,
        "source": source,
        "inherits": inherits,
        "rules": rules,
    }


def _json_to_format_profile(data: dict) -> FormatProfile:
    return FormatProfile(
        profile_id=data["profile_id"],
        name=data.get("name", ""),
        description=data.get("description", ""),
        source=FormatSource(mode=FormatMode(data.get("source", "custom"))),
        inherits=data.get("inherits"),
        **{key: value for key, value in (data.get("rules") or {}).items()},
    )


def save_user_profile(
    profile_id: str,
    name: str,
    description: str,
    source: str,
    inherits: str,
    rules: dict,
) -> dict:
    """Validate and atomically persist a user profile; reload registry."""
    if profile_registry.is_builtin(profile_id):
        return {"status": "error", "error": ERROR_PROFILE_ID_CONFLICT, "reason": f"built-in profile 不允许同名覆盖: {profile_id}"}
    errors = validate_profile_id(profile_id)
    if errors:
        return {"status": "error", "error": ERROR_INVALID_PROFILE, "reason": errors}
    if source not in ("reference", "custom", "guided"):
        return {"status": "error", "error": ERROR_INVALID_PROFILE, "reason": "source 必须是 reference|custom|guided"}
    if not name:
        return {"status": "error", "error": ERROR_INVALID_PROFILE, "reason": "name 不能为空"}
    if profile_registry.resolve_profile(inherits) is None:
        return {"status": "error", "error": ERROR_PROFILE_NOT_FOUND, "reason": f"inherits profile 不存在: {inherits}"}
    rule_result = validate_rules(rules)
    if rule_result["errors"]:
        return {"status": "error", "error": ERROR_INVALID_PROFILE, "reason": rule_result["errors"]}

    data = _profile_to_json(profile_id, name, description, source, inherits, rule_result["rules"])
    _atomic_write(_profile_path(profile_id), json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8"))
    reload_user_profiles()
    return {"status": "ok", "profile": data}


def load_user_profile(profile_id: str) -> Optional[dict]:
    path = _profile_path(profile_id)
    if not path.is_file():
        return None
    raw = json.loads(path.read_text(encoding="utf-8"))
    errors = validate_user_profile(raw)
    if errors:
        return None
    return raw


def list_user_profiles() -> tuple[list[dict], list[dict]]:
    """Return (valid_profiles, corrupted) where corrupted has path+errors."""
    profiles = []
    corrupted = []
    for path in sorted(profiles_dir().glob("*.json")):
        try:
            raw = json.loads(path.read_text(encoding="utf-8"))
            errors = validate_user_profile(raw)
            if errors:
                corrupted.append({"path": str(path), "errors": errors})
                continue
            profiles.append(raw)
        except Exception as exc:  # noqa: BLE001 - corrupted files must not crash
            corrupted.append({"path": str(path), "errors": [str(exc)]})
    return profiles, corrupted


def reload_user_profiles() -> list[dict]:
    """Re-register all persisted user profiles. Returns load warnings."""
    profile_registry.clear_user_profiles()
    profiles, corrupted = list_user_profiles()
    for data in profiles:
        profile_registry.register_user_profile(_json_to_format_profile(data))
    return corrupted


def delete_user_profile(profile_id: str) -> dict:
    if profile_registry.is_builtin(profile_id):
        return {"status": "error", "error": ERROR_BUILTIN_PROTECTED, "reason": f"built-in profile 不允许删除: {profile_id}"}
    path = _profile_path(profile_id)
    if not path.is_file() and not profile_registry.is_user_profile(profile_id):
        return {"status": "error", "error": ERROR_PROFILE_NOT_FOUND, "reason": f"profile 不存在: {profile_id}"}
    # PROFILE_IN_USE: any other user profile inherits this one
    profiles, _ = list_user_profiles()
    users = [p for p in profiles if p.get("inherits") == profile_id]
    if users:
        return {
            "status": "error",
            "error": ERROR_PROFILE_IN_USE,
            "reason": f"profile 被 {len(users)} 个用户 profile 继承，不能删除",
        }
    path.unlink(missing_ok=True)
    profile_registry.unregister_user_profile(profile_id)
    return {"status": "ok", "deleted": profile_id}


def update_user_profile(profile_id: str, updates: dict) -> dict:
    current = load_user_profile(profile_id)
    if current is None:
        if profile_registry.is_builtin(profile_id):
            return {"status": "error", "error": ERROR_BUILTIN_PROTECTED, "reason": f"built-in profile 不允许修改: {profile_id}"}
        return {"status": "error", "error": ERROR_PROFILE_NOT_FOUND, "reason": f"profile 不存在: {profile_id}"}
    merged = dict(current)
    if "name" in updates and updates["name"] is not None:
        merged["name"] = updates["name"]
    if "description" in updates and updates["description"] is not None:
        merged["description"] = updates["description"]
    if "rules" in updates and updates["rules"] is not None:
        # deep merge rules on top of existing rules
        existing_rules = dict(merged.get("rules") or {})
        new_rules = dict(updates["rules"] or {})
        for slot, value in new_rules.items():
            if isinstance(value, dict) and isinstance(existing_rules.get(slot), dict):
                merged_slot = dict(existing_rules[slot])
                merged_slot.update(value)
                existing_rules[slot] = merged_slot
            else:
                existing_rules[slot] = value
        merged["rules"] = existing_rules
    if "inherits" in updates and updates["inherits"] is not None:
        merged["inherits"] = updates["inherits"]

    errors = validate_user_profile(merged)
    if errors:
        return {"status": "error", "error": ERROR_INVALID_PROFILE, "reason": errors}
    if profile_registry.resolve_profile(merged["inherits"]) is None:
        return {"status": "error", "error": ERROR_PROFILE_NOT_FOUND, "reason": f"inherits profile 不存在: {merged['inherits']}"}
    _atomic_write(_profile_path(profile_id), json.dumps(merged, ensure_ascii=False, indent=2).encode("utf-8"))
    reload_user_profiles()
    return {"status": "ok", "profile": merged}
