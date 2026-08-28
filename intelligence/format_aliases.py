"""Mission 02-B — Format Aliases.

用户口语化格式提示 → preset profile id。

注意：FORMAT_ALIASES 不得与 classifier signals 混用。
"""

FORMAT_ALIASES: dict[str, str] = {
    "传统公文": "official_standard",
    "正式公文": "official_standard",
    "公文格式": "official_standard",
    "活动方案格式": "activity_plan_standard",
    "论文格式": "academic_standard",
    "培训通知格式": "training_notice_standard",
    "培训记录格式": "training_record_standard",
    "周报格式": "weekly_standard",
}


def resolve_alias(hint: str) -> str | None:
    """把用户格式提示解析为 preset profile id；未命中返回 None。"""
    if not hint:
        return None
    key = str(hint).strip()
    if key in FORMAT_ALIASES:
        return FORMAT_ALIASES[key]
    return None
