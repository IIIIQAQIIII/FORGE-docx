"""FORGE Format Model (Mission 01).

Four core models:
- ContentIntent    这份内容本质是什么？（只描述内容，绝不决定格式）
- FormatSource     这个格式从哪里来的？（preset / reference / custom / guided）
- FormatProfile    单份文档具体怎么排版？
- AssemblyProfile  多份文档汇成一本时整本怎么排？

设计原则：
- ContentIntent 不能写死模板映射。内容可以推荐格式，但内容不能替用户决定格式。
- FormatProfile 支持继承（inherits），继承规则是 deep merge：
  父 Profile 提供默认值，子 Profile 只覆盖自己明确提供的字段。
"""

from __future__ import annotations

import copy
from dataclasses import asdict, dataclass, field
from enum import Enum
from typing import Any, Optional


class ContentIntentType(str, Enum):
    """内容性质，只描述“是什么”，不决定“怎么排”。"""
    OFFICIAL_PLAN = "official_plan"
    OFFICIAL_SUMMARY = "official_summary"
    ACTIVITY_PLAN = "activity_plan"
    ACTIVITY_SUMMARY = "activity_summary"
    TRAINING_NOTICE = "training_notice"
    TRAINING_RECORD = "training_record"
    WEEKLY_REPORT = "weekly_report"
    ACADEMIC_PAPER = "academic_paper"
    SPEECH = "speech"
    LONG_FORM = "long_form"
    GENERIC = "generic"


class FormatMode(str, Enum):
    PRESET = "preset"
    REFERENCE = "reference"
    CUSTOM = "custom"
    GUIDED = "guided"


@dataclass
class ContentIntent:
    """内容意图：仅描述内容性质。

    注意：它不包含 template / format_profile 字段。
    """
    type: ContentIntentType = ContentIntentType.GENERIC
    description: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {"type": self.type.value, "description": self.description}

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "ContentIntent":
        return cls(
            type=ContentIntentType(data.get("type", "generic")),
            description=data.get("description", ""),
        )


@dataclass
class FormatSource:
    """格式来源：preset / reference / custom / guided。"""
    mode: FormatMode = FormatMode.PRESET
    origin: str = ""                 # preset 名 / reference 文件路径 / custom profile 名
    reference_path: Optional[str] = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "mode": self.mode.value,
            "origin": self.origin,
            "reference_path": self.reference_path,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "FormatSource":
        return cls(
            mode=FormatMode(data.get("mode", "preset")),
            origin=data.get("origin", ""),
            reference_path=data.get("reference_path"),
        )


@dataclass
class FormatProfile:
    """单份文档的排版规则。

    继承规则：resolve_profile() 会从 inherits 父 Profile 开始 deep merge。
    子 Profile 只覆盖自己明确提供的字段。
    """
    profile_id: str
    name: str = ""
    description: str = ""
    source: FormatSource = field(default_factory=FormatSource)
    inherits: Optional[str] = None

    page: dict[str, Any] = field(default_factory=dict)
    title: dict[str, Any] = field(default_factory=dict)
    subtitle: dict[str, Any] = field(default_factory=dict)
    organization: dict[str, Any] = field(default_factory=dict)
    author: dict[str, Any] = field(default_factory=dict)
    body: dict[str, Any] = field(default_factory=dict)
    heading_1: dict[str, Any] = field(default_factory=dict)
    heading_2: dict[str, Any] = field(default_factory=dict)
    heading_3: dict[str, Any] = field(default_factory=dict)
    caption: dict[str, Any] = field(default_factory=dict)
    table: dict[str, Any] = field(default_factory=dict)
    image: dict[str, Any] = field(default_factory=dict)
    signature: dict[str, Any] = field(default_factory=dict)
    date: dict[str, Any] = field(default_factory=dict)
    page_number: dict[str, Any] = field(default_factory=dict)
    custom_rules: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {
            "profile_id": self.profile_id,
            "name": self.name,
            "description": self.description,
            "source": self.source.to_dict(),
            "inherits": self.inherits,
            "page": copy.deepcopy(self.page),
            "title": copy.deepcopy(self.title),
            "subtitle": copy.deepcopy(self.subtitle),
            "organization": copy.deepcopy(self.organization),
            "author": copy.deepcopy(self.author),
            "body": copy.deepcopy(self.body),
            "heading_1": copy.deepcopy(self.heading_1),
            "heading_2": copy.deepcopy(self.heading_2),
            "heading_3": copy.deepcopy(self.heading_3),
            "caption": copy.deepcopy(self.caption),
            "table": copy.deepcopy(self.table),
            "image": copy.deepcopy(self.image),
            "signature": copy.deepcopy(self.signature),
            "date": copy.deepcopy(self.date),
            "page_number": copy.deepcopy(self.page_number),
            "custom_rules": copy.deepcopy(self.custom_rules),
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "FormatProfile":
        return cls(
            profile_id=data["profile_id"],
            name=data.get("name", ""),
            description=data.get("description", ""),
            source=FormatSource.from_dict(data.get("source", {})),
            inherits=data.get("inherits"),
            page=copy.deepcopy(data.get("page", {})),
            title=copy.deepcopy(data.get("title", {})),
            subtitle=copy.deepcopy(data.get("subtitle", {})),
            organization=copy.deepcopy(data.get("organization", {})),
            author=copy.deepcopy(data.get("author", {})),
            body=copy.deepcopy(data.get("body", {})),
            heading_1=copy.deepcopy(data.get("heading_1", {})),
            heading_2=copy.deepcopy(data.get("heading_2", {})),
            heading_3=copy.deepcopy(data.get("heading_3", {})),
            caption=copy.deepcopy(data.get("caption", {})),
            table=copy.deepcopy(data.get("table", {})),
            image=copy.deepcopy(data.get("image", {})),
            signature=copy.deepcopy(data.get("signature", {})),
            date=copy.deepcopy(data.get("date", {})),
            page_number=copy.deepcopy(data.get("page_number", {})),
            custom_rules=copy.deepcopy(data.get("custom_rules", {})),
        )


@dataclass
class AssemblyProfile:
    """多份文档汇编成一本时的整本规则。"""
    profile_id: str
    name: str = ""
    cover: dict[str, Any] = field(default_factory=dict)
    toc: dict[str, Any] = field(default_factory=dict)
    document_order: list[str] = field(default_factory=list)
    page_break_between_items: bool = True
    continuous_page_number: bool = True
    header: dict[str, Any] = field(default_factory=dict)
    footer: dict[str, Any] = field(default_factory=dict)
    custom_rules: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {
            "profile_id": self.profile_id,
            "name": self.name,
            "cover": copy.deepcopy(self.cover),
            "toc": copy.deepcopy(self.toc),
            "document_order": list(self.document_order),
            "page_break_between_items": self.page_break_between_items,
            "continuous_page_number": self.continuous_page_number,
            "header": copy.deepcopy(self.header),
            "footer": copy.deepcopy(self.footer),
            "custom_rules": copy.deepcopy(self.custom_rules),
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "AssemblyProfile":
        return cls(
            profile_id=data["profile_id"],
            name=data.get("name", ""),
            cover=copy.deepcopy(data.get("cover", {})),
            toc=copy.deepcopy(data.get("toc", {})),
            document_order=list(data.get("document_order", [])),
            page_break_between_items=data.get("page_break_between_items", True),
            continuous_page_number=data.get("continuous_page_number", True),
            header=copy.deepcopy(data.get("header", {})),
            footer=copy.deepcopy(data.get("footer", {})),
            custom_rules=copy.deepcopy(data.get("custom_rules", {})),
        )


def deep_merge(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    """深度合并：父 Profile 提供默认值，子 Profile 只覆盖自己明确提供的字段。

    规则：
    - 两个值都是 dict → 递归合并
    - 否则子值覆盖父值（包括显式的 None / 空 dict / 空 list）
    """
    result = copy.deepcopy(base)
    for key, value in (override or {}).items():
        if key in result and isinstance(result[key], dict) and isinstance(value, dict):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = copy.deepcopy(value)
    return result
