"""Mission 07 — Edit Engine service + paginated Inspect service."""

from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Any, Optional

from document_ir import DocumentIR, ParagraphBlock, TableBlock, read_docx
from edit_engine.models import EditOperation, EditPlan
from edit_engine.planner import build_plan, simulate_expected_ir
from edit_engine.renderer import render_edit
from edit_engine.validation import validate_edit_contract
from semantics.annotator import annotate_document

ERROR_SOURCE_NOT_FOUND = "SOURCE_NOT_FOUND"
ERROR_SOURCE_CHANGED = "SOURCE_CHANGED"
ERROR_SOURCE_OUTPUT_PATH_CONFLICT = "SOURCE_OUTPUT_PATH_CONFLICT"
EDIT_CONTRACT_VIOLATION = "EDIT_CONTRACT_VIOLATION"


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def _block_summary(block, max_text_chars: int) -> dict:
    text = getattr(block, "text", "") or ""
    if len(text) > max_text_chars:
        preview = text[:max_text_chars] + "…"
    else:
        preview = text
    summary = {
        "block_id": block.id,
        "block_type": block.type,
        "semantic_role": getattr(block, "semantic_role", None),
        "role_confidence": getattr(block, "role_confidence", None),
        "text_preview": preview,
        "source_locator": block.metadata.get("source_locator"),
        "metadata": {k: v for k, v in block.metadata.items() if k != "source_locator"},
    }
    if isinstance(block, TableBlock):
        summary["table"] = {
            "rows": len(block.rows),
            "cols": len(block.rows[0]) if block.rows else 0,
            "cells": [
                {
                    "row": ri,
                    "cell": ci,
                    "paragraphs": [
                        {
                            "block_id": p.id,
                            "text_preview": p.text[:max_text_chars],
                            "source_locator": p.metadata.get("source_locator"),
                        }
                        for p in cell.blocks
                    ],
                }
                for ri, row in enumerate(block.rows)
                for ci, cell in enumerate(row)
            ],
        }
    return summary


def inspect_document(
    source_path: str | Path,
    offset: int = 0,
    limit: int = 100,
    query: Optional[str] = None,
    roles: Optional[list[str]] = None,
    outline_only: bool = False,
    max_text_chars: int = 4000,
) -> dict:
    source = Path(source_path).expanduser()
    if not source.is_file():
        return {"status": "error", "errors": [ERROR_SOURCE_NOT_FOUND]}
    ir: DocumentIR = read_docx(source)
    annotate_document(ir)

    blocks = []
    for block in ir.blocks:
        if isinstance(block, ParagraphBlock):
            blocks.append(block)
        elif isinstance(block, TableBlock):
            blocks.append(block)

    if outline_only:
        blocks = [b for b in blocks if getattr(b, "semantic_role", None) in ("title", "subtitle", "organization", "author", "heading_1", "heading_2", "heading_3")]
    if roles:
        blocks = [b for b in blocks if getattr(b, "semantic_role", None) in roles]
    if query:
        q = query.lower()
        blocks = [b for b in blocks if q in (getattr(b, "text", "") or "").lower()]

    total_blocks = len(blocks)
    page = blocks[offset : offset + limit]
    next_offset = offset + len(page) if offset + len(page) < total_blocks else None

    return {
        "status": "ok",
        "source": str(source),
        "source_file_sha256": ir.source_file_sha256,
        "total_blocks": total_blocks,
        "offset": offset,
        "limit": limit,
        "next_offset": next_offset,
        "blocks": [_block_summary(b, max_text_chars) for b in page],
        "warnings": ir.warnings,
    }


def edit_document(
    source_path: str | Path,
    expected_source_sha256: str,
    edits: list[dict],
    output_path: Optional[str | Path] = None,
    dry_run: bool = False,
) -> dict:
    source = Path(source_path).expanduser()
    if not source.is_file():
        return {"status": "error", "errors": [ERROR_SOURCE_NOT_FOUND]}
    actual_sha = _sha256_file(source)
    if expected_source_sha256 and actual_sha != expected_source_sha256:
        return {"status": "error", "errors": [ERROR_SOURCE_CHANGED], "source_sha256": actual_sha}

    if output_path is None:
        output = source.parent / f"{source.stem}_EDITED.docx"
    else:
        output = Path(output_path).expanduser()
        if not output.is_absolute():
            output = output.resolve()
    if output.resolve() == source.resolve():
        return {"status": "error", "errors": [ERROR_SOURCE_OUTPUT_PATH_CONFLICT]}

    operations = [EditOperation(**e) for e in edits]
    plan: EditPlan = build_plan(source, actual_sha, operations)

    planned = len(plan.operations)
    if dry_run:
        return {
            "status": "dry_run",
            "source": str(source),
            "source_sha256": actual_sha,
            "operations": {"planned": planned, "applied": 0},
            "planned_operations": [b.to_dict() for b in plan.operations],
            "matches": [b.to_dict() for b in plan.operations if b.op == "replace_text"],
            "warnings": plan.warnings,
            "blockers": plan.blockers,
            "expected_change_summary": _change_summary_from_plan(plan),
        }

    if not plan.ready:
        return {
            "status": "error",
            "source": str(source),
            "source_sha256": actual_sha,
            "output": None,
            "operations": {"planned": planned, "applied": 0},
            "warnings": plan.warnings,
            "errors": plan.blockers,
        }

    source_ir = read_docx(source)
    expected_ir = simulate_expected_ir(source_ir, plan)

    render_result = render_edit(source, output, plan)
    if render_result.get("status") != "ok":
        return {
            "status": "error",
            "source": str(source),
            "source_sha256": actual_sha,
            "output": None,
            "operations": {"planned": planned, "applied": 0},
            "warnings": plan.warnings,
            "errors": render_result.get("errors", ["EDIT_RENDER_FAILED"]),
        }

    actual_ir = read_docx(output)
    contract = validate_edit_contract(expected_ir, actual_ir)
    if not contract["passed"]:
        try:
            Path(output).unlink()
        except OSError:
            pass
        return {
            "status": "error",
            "source": str(source),
            "source_sha256": actual_sha,
            "output": None,
            "operations": {"planned": planned, "applied": len(plan.operations)},
            "change_summary": _change_summary_from_plan(plan),
            "preservation": contract,
            "warnings": plan.warnings,
            "errors": [EDIT_CONTRACT_VIOLATION],
        }

    return {
        "status": "ok",
        "source": str(source),
        "output": str(output),
        "source_sha256": actual_sha,
        "operations": {"planned": planned, "applied": len(plan.operations)},
        "change_summary": _change_summary_from_plan(plan),
        "preservation": contract,
        "warnings": plan.warnings,
        "errors": [],
    }


def _change_summary_from_plan(plan: EditPlan) -> list:
    summary = []
    for b in plan.operations:
        if b.op == "replace_text":
            summary.append(f"replace_text:{b.old_text}→{b.new_text}@{b.locator}")
        elif b.op in ("insert_paragraph_before", "insert_paragraph_after"):
            summary.append(f"{b.op}:{b.text}@{b.locator}")
        elif b.op == "append_paragraph":
            summary.append(f"append_paragraph:{b.text}")
        elif b.op == "delete_paragraph":
            summary.append(f"delete_paragraph@{b.locator}")
    return summary
