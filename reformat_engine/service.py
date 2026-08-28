"""Mission 04-C — FORGE Reformat 2.0 unified internal service.

Pipeline:
    source DOCX
    → read_docx
    → semantic annotation
    → content classification (classifier input derived from visible IR text only)
    → resolve_format
    → build_plan
    → render_reformat
    → preservation validation
    → unified result

The renderer never re-runs classification or semantic annotation.
"""

from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Any, Optional

from document_ir import DocumentIR, OpaqueBlock, ParagraphBlock, TableBlock, read_docx
from intelligence.classifier import classify_content
from intelligence.resolver import resolve_format
from reformat_engine.models import ReformatPlan
from reformat_engine.planner import build_plan
from reformat_engine.renderer import render_reformat
from semantics.annotator import annotate_document

ERROR_SOURCE_NOT_FOUND = "SOURCE_NOT_FOUND"
ERROR_SOURCE_OUTPUT_PATH_CONFLICT = "SOURCE_OUTPUT_PATH_CONFLICT"
ERROR_SOURCE_CHANGED = "SOURCE_CHANGED"


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def _classifier_text_from_ir(ir: DocumentIR) -> str:
    """Build a raw visible-text string from the Document IR.

    This is not a rewrite or summary: it is the verbatim visible text of the
    body blocks, in document order, joined by newlines.
    """
    lines: list[str] = []

    def collect_text(text: str) -> None:
        if text:
            lines.append(text)

    for block in ir.blocks:
        if isinstance(block, ParagraphBlock):
            collect_text(block.text)
        elif isinstance(block, TableBlock):
            for row in block.rows:
                for cell in row:
                    for paragraph in cell.blocks:
                        collect_text(paragraph.text)
        elif isinstance(block, OpaqueBlock):
            collect_text(block.extracted_text)
    return "\n".join(lines)


def _empty_operations() -> dict[str, int]:
    return {"planned": 0, "applied": 0, "preserved": 0, "deferred": 0}


def _base_result(source: Path, status: str) -> dict[str, Any]:
    return {
        "status": status,
        "source": str(source),
        "output": None,
        "classification": None,
        "resolution": None,
        "format": None,
        "operations": _empty_operations(),
        "content_preservation": None,
        "warnings": [],
        "errors": [],
    }


def reformat_document(
    source_path: str | Path,
    output_path: Optional[str | Path] = None,
    explicit_profile_id: Optional[str] = None,
    explicit_format_hint: Optional[str] = None,
    reference_profile_id: Optional[str] = None,
    saved_profile_id: Optional[str] = None,
    allow_default: bool = False,
    default_profile_id: str = "generic_document",
) -> dict[str, Any]:
    """Unified FORGE Reformat 2.0 entry point.

    Returns a JSON-serializable result. ``status`` is one of:
    - ``ok``              output produced and content preservation passed
    - ``needs_guidance``  classification/resolution ambiguous, no output
    - ``error``           invalid profile / preservation failed / source changed
    """
    source = Path(source_path).expanduser()
    if not source.is_absolute():
        source = source.resolve()
    if not source.is_file():
        result = _base_result(source, "error")
        result["errors"].append(ERROR_SOURCE_NOT_FOUND)
        return result

    if output_path is None:
        output = source.parent / f"{source.stem}_FORGE.docx"
    else:
        output = Path(output_path).expanduser()
        if not output.is_absolute():
            output = output.resolve()

    if output.resolve() == source.resolve():
        result = _base_result(source, "error")
        result["output"] = str(output)
        result["errors"].append(ERROR_SOURCE_OUTPUT_PATH_CONFLICT)
        return result

    source_sha_before = _sha256_file(source)

    # 1. Faithful Inspector
    ir = read_docx(source)

    # 2. Semantic Role Annotation
    annotate_document(ir)

    # 3. Content Classification (visible IR text only; never a rewrite)
    classifier_text = _classifier_text_from_ir(ir)
    classification = classify_content(classifier_text)

    # 4. Format Resolution
    resolution = resolve_format(
        classification=classification,
        explicit_profile_id=explicit_profile_id,
        explicit_format_hint=explicit_format_hint,
        reference_profile_id=reference_profile_id,
        saved_profile_id=saved_profile_id,
        default_profile_id=default_profile_id,
        allow_default=allow_default,
    )

    warnings: list[str] = []
    errors: list[str] = []

    if resolution.get("status") == "needs_guidance":
        result = _base_result(source, "needs_guidance")
        result["classification"] = classification
        result["resolution"] = resolution
        result["warnings"] = resolution.get("reason", [])
        return result

    if resolution.get("status") == "error":
        result = _base_result(source, "error")
        result["classification"] = classification
        result["resolution"] = resolution
        result["errors"].append(resolution.get("error") or "RESOLUTION_ERROR")
        result["warnings"] = resolution.get("reason", [])
        return result

    target_profile_id = resolution.get("profile_id")
    if not target_profile_id:
        result = _base_result(source, "error")
        result["classification"] = classification
        result["resolution"] = resolution
        result["errors"].append("RESOLUTION_MISSING_PROFILE_ID")
        return result

    # 5. Reformat Plan
    plan: ReformatPlan = build_plan(ir, target_profile_id)
    warnings.extend(plan.warnings)
    errors.extend(plan.blockers)

    # 6. Source-Preserving Render + 7. Preservation validation
    render_result = render_reformat(source, plan, output)

    source_sha_after = _sha256_file(source)
    if source_sha_after != source_sha_before:
        render_result["status"] = "error"
        render_result["errors"] = list(render_result.get("errors") or []) + [ERROR_SOURCE_CHANGED]
        if output.exists():
            try:
                output.unlink()
            except OSError:
                pass
        render_result["output_path"] = None

    status = render_result.get("status")
    if status not in ("ok", "needs_guidance", "error"):
        status = "error"

    result = _base_result(source, status)
    result["classification"] = classification
    result["resolution"] = resolution
    result["format"] = {
        "profile_id": target_profile_id,
        "decision_basis": resolution.get("decision_basis"),
        "content_intent": classification.get("intent"),
    }
    result["operations"] = render_result.get("operations") or _empty_operations()
    result["content_preservation"] = render_result.get("content_preservation")
    result["warnings"] = warnings + list(render_result.get("warnings") or [])
    result["errors"] = errors + list(render_result.get("errors") or [])

    if status == "ok":
        result["output"] = str(output)
    else:
        result["output"] = None

    return result
