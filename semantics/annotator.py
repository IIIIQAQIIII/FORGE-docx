"""Mission 03-C2 — Extended Semantic Role Annotator.

只给 ParagraphBlock 增加 semantic_role / role_confidence / role_evidence
以及必要的 metadata（如 caption_type）。绝不修改内容、顺序、table、media。
"""

from __future__ import annotations

import math
from typing import Any

from document_ir import DocumentIR, ParagraphBlock, TableBlock
from semantics.role_signals import (
    AUTHOR_EVIDENCE_WEIGHTS,
    AUTHOR_LABEL_PATTERN,
    CAPTION_EVIDENCE_WEIGHTS,
    CAPTION_PATTERN,
    DATE_EVIDENCE_WEIGHTS,
    DATE_LIKE_PATTERN,
    EVIDENCE_WEIGHTS,
    FULL_DATE_PATTERN,
    HEADING_1_PATTERN,
    HEADING_2_PATTERN,
    HEADING_3_PATTERN,
    ORG_EVIDENCE_WEIGHTS,
    ORG_KEYWORDS,
    SUBTITLE_DASH_PATTERN,
    TITLE_WEIGHTS,
)
from semantics.roles import Role

CONFIDENCE_K = 5.0


def _confidence(score: float) -> float:
    if score <= 0:
        return 0.0
    return round(max(0.0, min(1.0, 1 - math.exp(-score / CONFIDENCE_K))), 4)


def _has_visible_inline(paragraph: ParagraphBlock) -> bool:
    for inline in paragraph.inline:
        if inline.type in ("image", "tab", "line_break", "page_break"):
            return True
        if inline.type == "text" and (inline.text or "").strip():
            return True
    return False


def _has_image(paragraph: ParagraphBlock) -> bool:
    return any(inline.type == "image" for inline in paragraph.inline)


def _is_date(paragraph: ParagraphBlock) -> bool:
    return bool(FULL_DATE_PATTERN.match(paragraph.text.strip()))


def _contains_org_keyword(text: str) -> bool:
    return any(keyword in text for keyword in ORG_KEYWORDS)


def _base_evidence(paragraph: ParagraphBlock, position: int, total: int) -> list[str]:
    evidence: list[str] = []
    text = paragraph.text.strip()
    if len(text) <= 30:
        evidence.append("short_paragraph")
    if text and not text.endswith(("。", "！", "？", "!", "?", ".")):
        evidence.append("no_sentence_end")
    if paragraph.style.bold:
        evidence.append("source_bold")
    if paragraph.style.font_size_pt and paragraph.style.font_size_pt >= 16:
        evidence.append("source_larger_font")
    if paragraph.style.alignment == "center":
        evidence.append("center_alignment")
    if position <= 2:
        evidence.append("document_start")
    if total > 0 and position >= total - 2:
        evidence.append("near_document_end")
    return evidence


def annotate_document(ir: DocumentIR) -> DocumentIR:
    paragraph_positions: dict[str, int] = {}
    paragraph_blocks: list[ParagraphBlock] = []
    top_paragraph_total = 0
    for block in ir.blocks:
        if isinstance(block, ParagraphBlock):
            paragraph_positions[block.id] = top_paragraph_total
            paragraph_blocks.append(block)
            top_paragraph_total += 1

    for idx, block in enumerate(ir.blocks):
        if not isinstance(block, ParagraphBlock):
            continue
        position = paragraph_positions[block.id]
        text = block.text.strip()

        prev_block = ir.blocks[idx - 1] if idx > 0 else None
        next_block = ir.blocks[idx + 1] if idx + 1 < len(ir.blocks) else None
        next_paragraph = next_block if isinstance(next_block, ParagraphBlock) else None
        prev_paragraph = prev_block if isinstance(prev_block, ParagraphBlock) else None

        if not text and not _has_visible_inline(block):
            block.semantic_role = Role.EMPTY.value
            block.role_confidence = _confidence(8.0)
            block.role_evidence = ["no_visible_text", "no_inline_content"]
            continue

        if not text and _has_visible_inline(block):
            block.semantic_role = Role.BODY.value
            block.role_confidence = _confidence(1.0)
            block.role_evidence = ["no_text_but_inline_content"]
            continue

        evidence = _base_evidence(block, position, top_paragraph_total)

        if isinstance(prev_block, TableBlock) or isinstance(next_block, TableBlock):
            evidence.append("adjacent_to_table")
        if (prev_paragraph is not None and _has_image(prev_paragraph)) or (
            next_paragraph is not None and _has_image(next_paragraph)
        ):
            evidence.append("adjacent_to_image")
        if next_paragraph is not None and _is_date(next_paragraph):
            evidence.append("before_date")
        if prev_paragraph is not None and prev_paragraph.semantic_role == Role.SIGNATURE.value:
            evidence.append("after_signature")

        role = None
        if HEADING_1_PATTERN.match(text):
            role = Role.HEADING_1.value
            evidence.insert(0, "matches_chinese_heading_level_1")
        elif HEADING_2_PATTERN.match(text):
            role = Role.HEADING_2.value
            evidence.insert(0, "matches_chinese_heading_level_2")
        elif HEADING_3_PATTERN.match(text) and not DATE_LIKE_PATTERN.match(text):
            role = Role.HEADING_3.value
            evidence.insert(0, "matches_chinese_heading_level_3")
        elif DATE_LIKE_PATTERN.match(text):
            role = Role.BODY.value
            evidence.insert(0, "date_like_excluded_from_heading")

        if role and role.startswith("heading_"):
            score = 4.0
            score += sum(EVIDENCE_WEIGHTS.get(e, 0.0) for e in evidence if e in EVIDENCE_WEIGHTS)
            block.semantic_role = role
            block.role_confidence = _confidence(score)
            block.role_evidence = evidence
            continue

        if _is_date(block):
            evidence.insert(0, "full_date_pattern")
            score = sum(DATE_EVIDENCE_WEIGHTS.get(e, 0.0) for e in evidence if e in DATE_EVIDENCE_WEIGHTS)
            block.semantic_role = Role.DATE.value
            block.role_confidence = _confidence(score)
            block.role_evidence = evidence
            continue

        caption_match = CAPTION_PATTERN.match(text)
        if caption_match:
            evidence.insert(0, "caption_pattern")
            score = sum(CAPTION_EVIDENCE_WEIGHTS.get(e, 0.0) for e in evidence if e in CAPTION_EVIDENCE_WEIGHTS)
            block.semantic_role = Role.CAPTION.value
            block.role_confidence = _confidence(score)
            block.role_evidence = evidence
            block.metadata["caption_type"] = "figure" if caption_match.group(1) == "图" else "table"
            continue

        author_match = AUTHOR_LABEL_PATTERN.match(text)
        if author_match:
            evidence.insert(0, "author_label")
            score = sum(AUTHOR_EVIDENCE_WEIGHTS.get(e, 0.0) for e in evidence if e in AUTHOR_EVIDENCE_WEIGHTS)
            block.semantic_role = Role.AUTHOR.value
            block.role_confidence = _confidence(score)
            block.role_evidence = evidence
            continue

        if SUBTITLE_DASH_PATTERN.match(text):
            evidence.append("dash_prefix")
            block.semantic_role = Role.SUBTITLE.value
            block.role_confidence = _confidence(4.0 + _score_with_weights(TITLE_WEIGHTS, evidence))
            block.role_evidence = evidence
            continue

        title_hits = [e for e in evidence if e in TITLE_WEIGHTS]
        if (
            position <= 3
            and len(text) <= 30
            and not text.endswith(("。", "！", "？", ".", "!", "?"))
            and len(title_hits) >= 3
            and ("source_larger_font" in title_hits or "source_bold" in title_hits)
        ):
            block.semantic_role = Role.TITLE.value
            block.role_confidence = _confidence(_score_with_weights(TITLE_WEIGHTS, title_hits))
            block.role_evidence = title_hits
            continue

        if _contains_org_keyword(text) and len(text) <= 30:
            org_hits = [e for e in evidence if e in ORG_EVIDENCE_WEIGHTS]
            org_hits.append("org_keyword")
            if position <= 2 and "before_date" not in org_hits:
                block.semantic_role = Role.ORGANIZATION.value
                block.role_confidence = _confidence(
                    sum(ORG_EVIDENCE_WEIGHTS.get(e, 0.0) for e in org_hits if e in ORG_EVIDENCE_WEIGHTS)
                )
                block.role_evidence = org_hits
                continue
            if position >= top_paragraph_total - 2 and "before_date" in org_hits:
                block.semantic_role = Role.SIGNATURE.value
                block.role_confidence = _confidence(
                    sum(ORG_EVIDENCE_WEIGHTS.get(e, 0.0) for e in org_hits if e in ORG_EVIDENCE_WEIGHTS)
                )
                block.role_evidence = org_hits
                continue

        if text:
            block.semantic_role = Role.BODY.value
            block.role_confidence = _confidence(1.0)
            block.role_evidence = evidence + ["non_empty_paragraph"]
        else:
            block.semantic_role = Role.UNKNOWN.value
            block.role_confidence = _confidence(0.5)
            block.role_evidence = ["unreliable"]

    # 第二遍：连续 heading 上下文增强
    heading_ids = [
        b.id
        for b in ir.blocks
        if isinstance(b, ParagraphBlock)
        and b.semantic_role in (Role.HEADING_1.value, Role.HEADING_2.value, Role.HEADING_3.value)
    ]
    for i, block_id in enumerate(heading_ids):
        block = next(b for b in ir.blocks if isinstance(b, ParagraphBlock) and b.id == block_id)
        if i > 0:
            prev = next(
                b for b in ir.blocks if isinstance(b, ParagraphBlock) and b.id == heading_ids[i - 1]
            )
            if prev.semantic_role == block.semantic_role:
                if "context_heading_before" not in block.role_evidence:
                    block.role_evidence.append("context_heading_before")
        if i < len(heading_ids) - 1:
            nxt = next(
                b for b in ir.blocks if isinstance(b, ParagraphBlock) and b.id == heading_ids[i + 1]
            )
            if nxt.semantic_role == block.semantic_role:
                if "context_heading_after" not in block.role_evidence:
                    block.role_evidence.append("context_heading_after")

    return ir


def _score_with_weights(weights: dict[str, float], evidence: list[str]) -> float:
    return round(sum(weights.get(e, 0.0) for e in evidence), 4)
