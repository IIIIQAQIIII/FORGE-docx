"""Mission 04-A — Profile coverage validation."""

from __future__ import annotations

from typing import Optional

from format_model import FormatProfile
from profiles import registry as profile_registry
from reformat_engine.models import ProfileCoverage

REQUIRED_SLOTS = [
    "title",
    "subtitle",
    "organization",
    "author",
    "heading_1",
    "heading_2",
    "heading_3",
    "body",
    "caption",
    "signature",
    "date",
    "table",
    "image",
]


def validate_profile_coverage(profile_or_id) -> ProfileCoverage:
    """先 resolve 继承链，再检查全部可格式化 slot 是否都有规则。

    profile_or_id 可以是 FormatProfile 对象或 profile_id 字符串。
    """
    if isinstance(profile_or_id, str):
        profile = profile_registry.resolve_profile(profile_or_id)
    else:
        profile = profile_or_id

    missing = []
    data = profile.to_dict()
    for slot in REQUIRED_SLOTS:
        value = data.get(slot)
        if not isinstance(value, dict) or not value:
            missing.append(slot)
    return ProfileCoverage(complete=not missing, missing_slots=missing)
