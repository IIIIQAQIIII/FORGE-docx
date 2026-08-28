"""Mission 03-C1 — Role signals.

角色判断信号集中维护。source format 只能作为 evidence，不能单独决定 role。
"""

from __future__ import annotations

import re

# 标题层级文本结构信号（强结构信号优先）
HEADING_1_PATTERN = re.compile(r"^[一二三四五六七八九十]+、")
HEADING_2_PATTERN = re.compile(r"^（[一二三四五六七八九十]+）")
HEADING_3_PATTERN = re.compile(r"^\d+[\.．、]")
# 日期/编号防误判：2026.9.1 不是 heading_3
DATE_LIKE_PATTERN = re.compile(r"^\d{4}[\.．]\d{1,2}[\.．]\d{1,2}")
SUBTITLE_DASH_PATTERN = re.compile(r"^——")

# 结构信号权重
STRUCTURE_WEIGHTS = {
    "heading_1": 6.0,
    "heading_2": 6.0,
    "heading_3": 5.0,
}

# 辅助 evidence 权重（不能单独决定 role）
EVIDENCE_WEIGHTS = {
    "short_paragraph": 1.5,        # 短段落
    "no_sentence_end": 1.5,        # 不以句号/叹号/问号结尾
    "source_bold": 1.0,
    "source_larger_font": 1.5,     # 字号 >= 16pt
    "center_alignment": 1.0,
    "document_start": 2.5,         # 文档开头位置
    "after_title": 2.0,
    "dash_prefix": 4.0,            # —— 副标题
    "context_heading_before": 1.5, # 前一个是同类标题
    "context_heading_after": 1.5,  # 后一个是同类标题
}

# title/subtitle 专属
TITLE_WEIGHTS = {
    "document_start": 2.5,
    "short_paragraph": 2.0,
    "no_sentence_end": 1.5,
    "source_larger_font": 1.5,
    "center_alignment": 1.5,
    "source_bold": 1.0,
}

BODY_WEIGHTS = {
    "sentence_end": 1.0,
    "long_paragraph": 1.5,
}

# Mission 03-C2 新增
FULL_DATE_PATTERN = re.compile(
    r"^(\d{4})[年\-/\.](\d{1,2})[月\-/\.](\d{1,2})日?$"
)
AUTHOR_LABEL_PATTERN = re.compile(r"^(作者|撰稿|执笔|汇报人)[：:]\s*(\S.*)$")
CAPTION_PATTERN = re.compile(r"^(图|表)\s*\d+")
ORG_KEYWORDS = (
    "幼儿园", "学校", "教育局", "中心", "公司", "委员会", "办公室", "部门", "小学", "中学", "大学",
)

ORG_EVIDENCE_WEIGHTS = {
    "org_keyword": 2.0,
    "short_paragraph": 1.0,
    "near_document_start": 1.5,
    "near_document_end": 1.5,
    "before_date": 2.0,
    "after_title": 1.0,
}

DATE_EVIDENCE_WEIGHTS = {
    "full_date_pattern": 8.0,
    "near_document_end": 1.0,
    "after_signature": 1.5,
}

AUTHOR_EVIDENCE_WEIGHTS = {
    "author_label": 8.0,
    "short_paragraph": 1.0,
}

CAPTION_EVIDENCE_WEIGHTS = {
    "caption_pattern": 7.0,
    "adjacent_to_image": 1.5,
    "adjacent_to_table": 1.5,
    "short_paragraph": 1.0,
}
