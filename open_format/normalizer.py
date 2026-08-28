"""Unified rule schema validation / normalization for user profiles."""

from __future__ import annotations

from typing import Any

from open_format.config import ALLOWED_RULE_SLOTS

_POSITIVE_NUMBERS = ("size_pt", "line_spacing_pt", "space_before_pt", "space_after_pt", "first_line_twips", "left_twips", "right_twips")
_NUMBERS = (
    "first_line_chars",
    "left_chars",
    "right_chars",
    "width_cm",
    "height_cm",
    "top_cm",
    "bottom_cm",
    "left_cm",
    "right_cm",
    "header_distance_cm",
    "footer_distance_cm",
    "preferred_width_cm",
    "max_width_cm",
    "max_height_cm",
    "margin_top_cm",
    "margin_bottom_cm",
    "margin_left_cm",
    "margin_right_cm",
)
_BOOLS = ("bold", "italic", "autofit", "preserve_aspect_ratio", "allow_upscale", "enabled", "show_on_first_page", "continuous")
_STRINGS = ("font", "color", "shading", "name", "style")
_ALIGNMENT_FIELDS = ("align", "alignment")
_ALIGNMENTS = ("left", "center", "right", "both", "justify", "center_right")

_PAGE_NUMBER_POSITIONS = ("footer", "header")
_TABLE_ALIGNMENTS = ("left", "center", "right")
_VERTICAL_ALIGNMENTS = ("top", "center", "bottom")


def validate_profile_id(profile_id: str) -> list[str]:
    errors = []
    if not profile_id or not isinstance(profile_id, str):
        errors.append("profile_id 必须是非空字符串")
        return errors
    allowed = set("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_.-")
    if any(ch not in allowed for ch in profile_id):
        errors.append("profile_id 仅允许字母/数字/_.-")
    return errors


def validate_rules(rules: Any) -> dict:
    """Return {"errors": [...], "rules": normalized_rules}."""
    errors = []
    if rules is None:
        return {"errors": errors, "rules": {}}
    if not isinstance(rules, dict):
        return {"errors": ["rules 必须是对象"], "rules": {}}
    normalized = {}
    for slot, value in rules.items():
        if slot not in ALLOWED_RULE_SLOTS:
            errors.append(f"未知 slot: {slot}")
            continue
        if value is None:
            continue
        if not isinstance(value, dict):
            errors.append(f"slot {slot} 必须是对象")
            continue
        slot_errors, slot_value = _validate_slot(slot, value)
        errors.extend(slot_errors)
        if not slot_errors:
            normalized[slot] = slot_value
    return {"errors": errors, "rules": normalized}


def _validate_slot(slot: str, value: dict) -> tuple[list, dict]:
    errors = []
    out = {}

    def num(key, positive=False):
        if key not in value:
            return
        v = value[key]
        if isinstance(v, bool) or not isinstance(v, (int, float)):
            errors.append(f"{slot}.{key} 必须是数字")
            return
        if positive and v <= 0:
            errors.append(f"{slot}.{key} 必须 > 0")
            return
        out[key] = float(v) if isinstance(v, int) else float(v)

    def boolean(key):
        if key not in value:
            return
        v = value[key]
        if not isinstance(v, bool):
            errors.append(f"{slot}.{key} 必须是布尔值")
            return
        out[key] = v

    def string(key):
        if key not in value:
            return
        v = value[key]
        if not isinstance(v, str) or not v.strip():
            errors.append(f"{slot}.{key} 必须是非空字符串")
            return
        out[key] = v.strip()

    def alignment(key):
        if key not in value:
            return
        v = value[key]
        if v not in _ALIGNMENTS:
            errors.append(f"{slot}.{key} 非法对齐: {v}")
            return
        out[key] = v

    for key in _POSITIVE_NUMBERS:
        num(key, positive=True)
    for key in _NUMBERS:
        num(key)
    for key in _BOOLS:
        boolean(key)
    for key in _STRINGS:
        string(key)
    alignment("align")
    alignment("alignment")

    if slot == "page_number":
        if "position" in value:
            if value["position"] not in _PAGE_NUMBER_POSITIONS:
                errors.append(f"page_number.position 非法: {value['position']}")
            else:
                out["position"] = value["position"]
        if "start_at" in value:
            v = value["start_at"]
            if isinstance(v, bool) or not isinstance(v, int) or v < 0:
                errors.append("page_number.start_at 必须是非负整数")
            else:
                out["start_at"] = v
        if "alignment" in value:
            if value["alignment"] not in _ALIGNMENTS:
                errors.append(f"page_number.alignment 非法: {value['alignment']}")
            else:
                out["alignment"] = value["alignment"]

    if slot == "table":
        if "alignment" in value:
            if value["alignment"] not in _TABLE_ALIGNMENTS:
                errors.append(f"table.alignment 非法: {value['alignment']}")
            else:
                out["alignment"] = value["alignment"]
        if "borders" in value and not isinstance(value["borders"], dict):
            errors.append("table.borders 必须是对象")
        else:
            borders = value.get("borders")
            if isinstance(borders, dict):
                out["borders"] = _validate_borders(borders, errors, slot)
        if "text" in value:
            if not isinstance(value["text"], dict):
                errors.append("table.text 必须是对象")
            else:
                sub_errors, sub = _validate_slot(slot + ".text", value["text"])
                errors.extend(sub_errors)
                if not sub_errors:
                    out["text"] = sub
        if "cell" in value:
            if not isinstance(value["cell"], dict):
                errors.append("table.cell 必须是对象")
            else:
                sub_errors, sub = _validate_slot(slot + ".cell", value["cell"])
                errors.extend(sub_errors)
                if not sub_errors:
                    out["cell"] = sub

    if slot == "image":
        if "alignment" in value:
            if value["alignment"] not in _ALIGNMENTS:
                errors.append(f"image.alignment 非法: {value['alignment']}")
            else:
                out["alignment"] = value["alignment"]

    return errors, out


def _validate_borders(borders: dict, errors: list, slot: str) -> dict:
    out = {}
    edge_keys = ("top", "bottom", "left", "right", "inside_horizontal", "inside_vertical")
    for edge in edge_keys:
        if edge not in borders:
            continue
        spec = borders[edge]
        if not isinstance(spec, dict):
            errors.append(f"{slot}.borders.{edge} 必须是对象")
            continue
        edge_out = {}
        if "val" in spec and isinstance(spec["val"], str):
            edge_out["val"] = spec["val"]
        if "sz" in spec and isinstance(spec["sz"], (int, float, str)):
            edge_out["sz"] = str(spec["sz"])
        if "color" in spec and isinstance(spec["color"], str):
            edge_out["color"] = spec["color"]
        out[edge] = edge_out
    return out


def validate_user_profile(data: dict) -> dict:
    """Validate a full user profile JSON document. Returns errors list."""
    errors = []
    if not isinstance(data, dict):
        return ["profile JSON 必须是对象"]
    if data.get("schema_version") != 1:
        errors.append("schema_version 必须为 1")
    errors.extend(validate_profile_id(data.get("profile_id", "")))
    if not data.get("name"):
        errors.append("name 不能为空")
    if data.get("source") not in ("reference", "custom", "guided"):
        errors.append("source 必须是 reference|custom|guided")
    inherits = data.get("inherits")
    if not isinstance(inherits, str) or not inherits:
        errors.append("inherits 必须是非空字符串")
    rule_result = validate_rules(data.get("rules"))
    errors.extend(rule_result["errors"])
    return errors
