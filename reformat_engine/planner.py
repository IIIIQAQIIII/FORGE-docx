"""Mission 04-A — Reformat Planner.

输入：Annotated Document IR + Target FormatProfile
输出：ReformatPlan

本阶段只做规划，绝对不生成 DOCX，不修改 IR / 原文档。
"""

from __future__ import annotations

from dataclasses import asdict
from typing import Optional

from document_ir import DocumentIR, OpaqueBlock, ParagraphBlock, TableBlock
from format_model import FormatProfile
from profiles import registry as profile_registry
from reformat_engine.models import Operation, ReformatPlan
from reformat_engine.profile_coverage import validate_profile_coverage

# Role -> Style Slot 集中映射（不得散落成大量 if/elif）
ROLE_STYLE_SLOT = {
    "title": "title",
    "subtitle": "subtitle",
    "organization": "organization",
    "author": "author",
    "heading_1": "heading_1",
    "heading_2": "heading_2",
    "heading_3": "heading_3",
    "body": "body",
    "caption": "caption",
    "signature": "signature",
    "date": "date",
}

ERROR_NOT_FOUND = "PROFILE_NOT_FOUND"


def _resolve_profile(target_profile_id: str) -> Optional[FormatProfile]:
    return profile_registry.resolve_profile(target_profile_id)


def build_plan(ir: DocumentIR, target_profile_id: str) -> ReformatPlan:
    """为 Document IR 中的每个 block 规划 target profile 的格式规则。"""
    try:
        profile = _resolve_profile(target_profile_id)
    except KeyError:
        plan = ReformatPlan(
            target_profile_id=target_profile_id,
            source_fingerprint=asdict(ir.content_fingerprint),
            ready=False,
            blockers=[ERROR_NOT_FOUND],
        )
        return plan

    plan = ReformatPlan(
        target_profile_id=target_profile_id,
        source_fingerprint=asdict(ir.content_fingerprint),
        source_file_sha256=ir.source_file_sha256,
        ready=True,
    )
    plan.profile_coverage = validate_profile_coverage(profile)

    for block in ir.blocks:
        if isinstance(block, ParagraphBlock):
            op = _plan_paragraph(block)
            plan.operations.append(op)
        elif isinstance(block, TableBlock):
            plan.operations.append(
                Operation(
                    block_id=block.id,
                    block_type="table",
                    semantic_role=None,
                    role_confidence=None,
                    action="apply_profile_style",
                    style_slot="table",
                    reason="表格按 target profile 的 table 规则格式化（仅规划，不实际修改）。",
                    metadata={"source_locator": block.metadata.get("source_locator")},
                )
            )
        elif isinstance(block, OpaqueBlock):
            plan.operations.append(
                Operation(
                    block_id=block.id,
                    block_type="opaque",
                    semantic_role=None,
                    role_confidence=None,
                    action="unsupported",
                    style_slot=None,
                    reason="OpaqueBlock 当前无法安全重建，必须人工处理。",
                    metadata={"source_locator": block.metadata.get("source_locator")},
                )
            )
            plan.ready = False
            plan.blockers.append(f"OpaqueBlock 无法安全重建: {block.id} ({block.xml_tag})")
        else:
            plan.operations.append(
                Operation(
                    block_id=getattr(block, "id", "?"),
                    block_type=getattr(block, "type", "unknown"),
                    semantic_role=None,
                    role_confidence=None,
                    action="unsupported",
                    style_slot=None,
                    reason="未知 block 类型，当前无法安全重建。",
                )
            )
            plan.ready = False
            plan.blockers.append(f"未知 block 类型: {block.id}")

    return plan


def _plan_paragraph(block: ParagraphBlock) -> Operation:
    role = block.semantic_role
    confidence = block.role_confidence

    if role == "empty":
        return Operation(
            block_id=block.id,
            block_type="paragraph",
            semantic_role=role,
            role_confidence=confidence,
            action="preserve_structure",
            style_slot=None,
            reason="空段落保留结构，不套样式。",
            metadata={"source_locator": block.metadata.get("source_locator")},
        )

    if role == "unknown":
        return Operation(
            block_id=block.id,
            block_type="paragraph",
            semantic_role=role,
            role_confidence=confidence,
            action="review_required",
            style_slot=None,
            reason="语义角色未知，需要人工确认。",
            metadata={"source_locator": block.metadata.get("source_locator")},
        )

    slot = ROLE_STYLE_SLOT.get(role or "")
    if slot is None:
        return Operation(
            block_id=block.id,
            block_type="paragraph",
            semantic_role=role,
            role_confidence=confidence,
            action="review_required",
            style_slot=None,
            reason=f"角色 {role} 没有对应的 style slot 映射。",
        )

    metadata = {"source_locator": block.metadata.get("source_locator")}
    if any(inline.type == "image" for inline in block.inline):
        metadata["target_image_slot"] = "image"
        metadata["image_count"] = sum(1 for inline in block.inline if inline.type == "image")
        metadata["image_note"] = "图片不移动、不删除、不替换、不重新排序。"

    return Operation(
        block_id=block.id,
        block_type="paragraph",
        semantic_role=role,
        role_confidence=confidence,
        action="apply_profile_style",
        style_slot=slot,
        reason=f"已识别角色 {role}，应用 target profile 的 {slot} 规则。",
        metadata=metadata,
    )
