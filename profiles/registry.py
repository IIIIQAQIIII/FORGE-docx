"""FORGE Profile Registry (Mission 01).

内置函数：
- register_profile()  注册格式
- get_profile()       取原始 Profile（不做继承）
- list_profiles()     列出全部 Profile
- resolve_profile()   读取父 Profile → 继承 → 覆盖 → 输出最终完整 Profile

第一轮只做登记，生成器暂不使用这些 Profile。
"""

from __future__ import annotations

from typing import Optional

from format_model import AssemblyProfile, FormatProfile, FormatSource, deep_merge

_PROFILES: dict[str, FormatProfile] = {}
_ASSEMBLY_PROFILES: dict[str, AssemblyProfile] = {}
_USER_PROFILE_IDS: set[str] = set()


def register_profile(profile: FormatProfile) -> None:
    _PROFILES[profile.profile_id] = profile


def register_user_profile(profile: FormatProfile) -> None:
    _PROFILES[profile.profile_id] = profile
    _USER_PROFILE_IDS.add(profile.profile_id)


def unregister_user_profile(profile_id: str) -> None:
    _PROFILES.pop(profile_id, None)
    _USER_PROFILE_IDS.discard(profile_id)


def clear_user_profiles() -> None:
    for profile_id in list(_USER_PROFILE_IDS):
        _PROFILES.pop(profile_id, None)
    _USER_PROFILE_IDS.clear()


def is_builtin(profile_id: str) -> bool:
    return profile_id in _PROFILES and profile_id not in _USER_PROFILE_IDS


def is_user_profile(profile_id: str) -> bool:
    return profile_id in _USER_PROFILE_IDS


def get_profile(profile_id: str) -> Optional[FormatProfile]:
    """取原始 Profile（不含继承结果）。"""
    return _PROFILES.get(profile_id)


def list_profiles() -> dict[str, FormatProfile]:
    return dict(_PROFILES)


def list_profiles_detailed() -> dict[str, dict]:
    """List profiles with origin = builtin | user."""
    return {
        profile_id: {"profile": profile, "origin": "user" if profile_id in _USER_PROFILE_IDS else "builtin"}
        for profile_id, profile in _PROFILES.items()
    }


def resolve_profile(profile_id: str, _stack: Optional[tuple[str, ...]] = None) -> FormatProfile:
    """解析继承链，返回 deep merge 后的完整 Profile。

    规则：
    - 父 Profile 提供默认值，子 Profile 只覆盖自己明确提供的字段。
    - 字段级 deep merge（不是 dict.update）。
    """
    raw = _PROFILES.get(profile_id)
    if raw is None:
        raise KeyError(f"unknown profile: {profile_id}")

    stack = _stack or ()
    if profile_id in stack:
        raise ValueError(f"inheritance cycle detected: {' -> '.join(stack + (profile_id,))}")

    if raw.inherits:
        parent = resolve_profile(raw.inherits, stack + (profile_id,))
        merged_data = deep_merge(parent.to_dict(), raw.to_dict())
        merged_data.pop("inherits", None)
        merged_data["profile_id"] = profile_id
        return FormatProfile.from_dict(merged_data)

    return FormatProfile.from_dict(raw.to_dict())


# ---------------------------------------------------------------------------
# Assembly profiles
# ---------------------------------------------------------------------------

def register_assembly_profile(profile: AssemblyProfile) -> None:
    _ASSEMBLY_PROFILES[profile.profile_id] = profile


def get_assembly_profile(profile_id: str) -> Optional[AssemblyProfile]:
    return _ASSEMBLY_PROFILES.get(profile_id)


def list_assembly_profiles() -> dict[str, AssemblyProfile]:
    return dict(_ASSEMBLY_PROFILES)


# ---------------------------------------------------------------------------
# 内置 Profiles（第一轮只登记，不使用）
# ---------------------------------------------------------------------------

register_profile(FormatProfile(
    profile_id="generic_document",
    name="通用文档",
    description="开放格式体系的地基：A4、普通标题、标准正文、三级标题、普通表格、连续页码。",
    source=FormatSource(),
    page={"width_cm": 21, "height_cm": 29.7, "top_cm": 2.54, "bottom_cm": 2.54, "left_cm": 3.17, "right_cm": 3.17},
    title={"font": "宋体", "size_pt": 16, "bold": False, "align": "center"},
    subtitle={"font": "楷体", "size_pt": 14, "align": "center"},
    organization={"font": "宋体", "size_pt": 14, "bold": True, "align": "center"},
    author={"font": "宋体", "size_pt": 12, "align": "center"},
    body={"font": "宋体", "size_pt": 12, "line_spacing_pt": 18},
    heading_1={"font": "黑体", "size_pt": 16},
    heading_2={"font": "黑体", "size_pt": 14},
    heading_3={"font": "黑体", "size_pt": 12},
    caption={"font": "黑体", "size_pt": 10.5, "align": "center"},
    table={"style": "Table Grid"},
    image={"align": "center"},
    signature={"font": "宋体", "size_pt": 12, "align": "right"},
    date={"font": "宋体", "size_pt": 12, "align": "right"},
    page_number={
        "enabled": True,
        "position": "footer",
        "alignment": "center",
        "start_at": None,
        "show_on_first_page": True,
        "continuous": True,
    },
))

register_profile(FormatProfile(
    profile_id="official_standard",
    name="标准传统公文",
    description="标题两行居中、正文3号仿宋28磅、三级层次、落款中线靠右。",
    source=FormatSource(),
    inherits="generic_document",
    page={"top_cm": 3.7, "bottom_cm": 3.5, "left_cm": 2.8, "right_cm": 2.6},
    title={"font": "方正小标宋简体", "size_pt": 22, "align": "center", "line_spacing_pt": 32},
    body={"font": "仿宋_GB2312", "size_pt": 16, "line_spacing_pt": 28},
    heading_1={"font": "黑体", "size_pt": 16},
    heading_2={"font": "楷体_GB2312", "size_pt": 16},
    heading_3={"font": "仿宋_GB2312", "size_pt": 16},
    signature={"align": "center_right", "department": False},
    page_number={"position": "footer", "alignment": "center", "show_on_first_page": True, "font": "宋体", "size_pt": 14},
))

register_profile(FormatProfile(
    profile_id="activity_plan_standard",
    name="活动方案标准",
    description="活动方案/总结：继承标准传统公文，落款增加部门。",
    source=FormatSource(),
    inherits="official_standard",
    signature={"department": True},
))

register_profile(FormatProfile(
    profile_id="academic_standard",
    name="论文/长文标准",
    description="无页眉、标题不加粗、正文小四宋体18磅、三级标题、三线表/插图。",
    source=FormatSource(),
    inherits="generic_document",
    page={"top_cm": 3.0, "bottom_cm": 2.5, "left_cm": 3.0, "right_cm": 2.5},
    title={"font": "黑体", "size_pt": 16, "bold": False, "align": "center"},
    body={"font": "宋体", "size_pt": 12, "line_spacing_pt": 18},
    heading_1={"font": "黑体", "size_pt": 16},
    heading_2={"font": "黑体", "size_pt": 14},
    heading_3={"font": "黑体", "size_pt": 12},
    table={"style": "three_line"},
    image={"align": "center", "caption_font": "黑体", "caption_size_pt": 10.5},
))

register_profile(FormatProfile(
    profile_id="weekly_standard",
    name="行政周报标准",
    description="两行标题、部门分节、条目自动编号、空缺写“无”。",
    source=FormatSource(),
    inherits="official_standard",
    title={"font": "方正小标宋简体", "size_pt": 22, "align": "center", "line_spacing_pt": 32},
    body={"font": "仿宋_GB2312", "size_pt": 16, "line_spacing_pt": 28},
))

register_profile(FormatProfile(
    profile_id="training_notice_standard",
    name="培训通知标准",
    description="红色红头+双红线，正文固定28磅，默认一页。",
    source=FormatSource(),
    inherits="official_standard",
    title={"font": "方正小标宋简体", "size_pt": 36, "color": "FF0000"},
    body={"font": "仿宋_GB2312", "size_pt": 16, "line_spacing_pt": 28},
    custom_rules={"one_page": True, "flexible_masthead": True},
))

register_profile(FormatProfile(
    profile_id="training_record_standard",
    name="培训活动记录标准",
    description="5行记录表，参培人行18磅，培训内容28磅首行缩进两字符。",
    source=FormatSource(),
    inherits="generic_document",
    table={"style": "fixed_5_row"},
))

register_profile(FormatProfile(
    profile_id="activity_summary_standard",
    name="活动总结标准",
    description="活动总结：继承活动方案标准，落款单位+部门+日期。",
    source=FormatSource(),
    inherits="activity_plan_standard",
    signature={"department": True},
))

register_profile(FormatProfile(
    profile_id="activity_archive_standard",
    name="活动影像标准",
    description="活动影像：标题+信息表（时间/活动地点/负责人/活动内容）+两张照片同一页。",
    source=FormatSource(),
    inherits="generic_document",
    table={"style": "fixed_4_row"},
    image={"align": "center"},
))

register_profile(FormatProfile(
    profile_id="training_archive_standard",
    name="培训活动影像标准",
    description="培训活动影像：4行照片表，两张照片同一页。",
    source=FormatSource(),
    inherits="generic_document",
    table={"style": "fixed_4_row"},
    image={"align": "center"},
))
