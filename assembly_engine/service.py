"""Mission 05 — Batch Assemble Engine service."""

from __future__ import annotations

import hashlib
import json
import shutil
import tempfile
import time
import uuid
from dataclasses import asdict
from pathlib import Path
from typing import Any, Optional

from assembly_engine.importer import (
    ImportState,
    UnsupportedPackageContent,
    _append_body_child,
    _serialize_xml,
    import_item_body,
    init_import_state,
    make_separator,
    read_zip_parts,
    write_zip_parts,
)
from assembly_engine.models import AssemblyItemResult, AssemblyResult
from assembly_engine.preservation import compute_assembly_payload_sha256, verify_item_preservation
from document_ir import OpaqueBlock, read_docx
from format_model import AssemblyProfile, FormatProfile, FormatSource
from intelligence.classifier import classify_content
from intelligence.mappings import CONTENT_PROFILE_RECOMMENDATIONS
from intelligence.resolver import resolve_format
from profiles import registry as profile_registry
from reformat_engine.renderer import (
    _apply_page_format,
    _apply_page_number,
    _ensure_chrome_reference,
    _find_sect_prs,
)
from reformat_engine.service import _classifier_text_from_ir, reformat_document as reformat_single

ERROR_ASSEMBLY_PROFILE_NOT_FOUND = "ASSEMBLY_PROFILE_NOT_FOUND"
ERROR_SOURCE_OUTPUT_PATH_CONFLICT = "SOURCE_OUTPUT_PATH_CONFLICT"


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def _default_assembly_profile() -> AssemblyProfile:
    return AssemblyProfile(
        profile_id="default_assembly",
        page_break_between_items=True,
        continuous_page_number=True,
        header={},
        footer={},
    )


def _resolve_batch_format(
    sources: list[Path],
    explicit_profile_id: Optional[str],
    explicit_format_hint: Optional[str],
    reference_profile_id: Optional[str],
    saved_profile_id: Optional[str],
    allow_default: bool,
) -> dict:
    if explicit_profile_id or explicit_format_hint or reference_profile_id or saved_profile_id:
        return resolve_format(
            classification=None,
            explicit_profile_id=explicit_profile_id,
            explicit_format_hint=explicit_format_hint,
            reference_profile_id=reference_profile_id,
            saved_profile_id=saved_profile_id,
            allow_default=allow_default,
        )

    recommendations: dict[str, dict] = {}
    for source in sources:
        ir = read_docx(source)
        classification = classify_content(_classifier_text_from_ir(ir))
        if classification.get("intent") == "generic" or classification.get("status") == "ambiguous":
            continue
        mapped = CONTENT_PROFILE_RECOMMENDATIONS.get(classification.get("intent", ""), "")
        if mapped:
            if mapped not in recommendations or classification.get("confidence", 0.0) > recommendations[mapped].get(
                "confidence", 0.0
            ):
                recommendations[mapped] = classification

    if len(recommendations) == 1:
        profile_id = next(iter(recommendations))
        representative = recommendations[profile_id]
        classification = {
            "intent": representative.get("intent"),
            "confidence": representative.get("confidence", 0.0),
            "status": "recommended",
            "signals": representative.get("signals", []),
            "alternatives": [],
        }
        return resolve_format(classification=classification, allow_default=allow_default)

    return resolve_format(classification=None, allow_default=allow_default)


def _unsupported_package_check(normalized_path: Path, normalized_ir) -> list[str]:
    errors = []
    for warning in normalized_ir.warnings:
        if "检测到暂不完整支持" in warning:
            errors.append("UNSUPPORTED_PACKAGE_CONTENT:" + warning)
    if any(isinstance(block, OpaqueBlock) for block in normalized_ir.blocks):
        errors.append("UNSUPPORTED_PACKAGE_CONTENT:OpaqueBlock")
    parts = read_zip_parts(normalized_path)
    for part in ("word/footnotes.xml", "word/endnotes.xml", "word/comments.xml"):
        if part in parts:
            errors.append(f"UNSUPPORTED_PACKAGE_CONTENT:{part}")
    return errors


def _normalize_profile_for(target_profile_id: str) -> str:
    temp_id = f"_assembly_normalize_{uuid.uuid4().hex[:12]}"
    profile_registry.register_profile(
        FormatProfile(
            profile_id=temp_id,
            name=temp_id,
            source=FormatSource(),
            inherits=target_profile_id,
            page_number={"enabled": False},
        )
    )
    return temp_id


def _set_chrome_text(parts: dict, part_name: str, text: str, alignment: str) -> None:
    from lxml import etree

    root = etree.fromstring(parts[part_name])
    for p in list(root.findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p")):
        root.remove(p)
    p = etree.fromstring(
        f'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:pPr><w:jc w:val="{alignment}"/></w:pPr><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'.encode(
            "utf-8"
        )
    )
    root.append(p)
    parts[part_name] = _serialize_xml(root)


def _build_master_parts(normalized_paths: list[Path], assembly_profile: AssemblyProfile, target_profile_id: str) -> dict:
    from docx import Document as DocxDocument

    base_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    base_tmp.close()
    try:
        DocxDocument().save(base_tmp.name)
        parts = read_zip_parts(Path(base_tmp.name))
    finally:
        Path(base_tmp.name).unlink(missing_ok=True)

    state: ImportState = init_import_state(parts)
    page_break_between_items = bool(assembly_profile.page_break_between_items)
    for index, normalized_path in enumerate(normalized_paths):
        if index > 0 and page_break_between_items:
            _append_body_child(state.body, make_separator())
        import_item_body(state, normalized_path, index)

    body = state.body
    profile = profile_registry.resolve_profile(target_profile_id)
    _apply_page_format(body, profile.page or {})

    # assembly-level chrome: header / footer text + unified page number.
    sect_pr = body.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr")
    if sect_pr is None:
        sect_prs = _find_sect_prs(body)
        sect_pr = sect_prs[-1] if sect_prs else None
    if sect_pr is not None:
        header_cfg = assembly_profile.header or {}
        footer_cfg = assembly_profile.footer or {}
        if header_cfg.get("text"):
            part_name = _ensure_chrome_reference(
                parts, sect_pr, "header", "default", False, header_cfg.get("alignment", "center"), None, None
            )
            _set_chrome_text(parts, part_name, str(header_cfg.get("text")), header_cfg.get("alignment", "center"))
        if footer_cfg.get("text"):
            part_name = _ensure_chrome_reference(
                parts, sect_pr, "footer", "default", False, footer_cfg.get("alignment", "center"), None, None
            )
            _set_chrome_text(parts, part_name, str(footer_cfg.get("text")), footer_cfg.get("alignment", "center"))
        if assembly_profile.continuous_page_number:
            _apply_page_number(
                body,
                {
                    "enabled": True,
                    "position": "footer",
                    "alignment": footer_cfg.get("alignment", "center"),
                    "show_on_first_page": True,
                    "start_at": None,
                },
                parts,
            )

    parts["word/document.xml"] = _serialize_xml(state.document_root)
    return parts


def _jobs_dir() -> Path:
    from open_format.home import get_forge_home

    return get_forge_home() / "jobs"


def _job_dir(job_id: str) -> Path:
    return _jobs_dir() / job_id


def _manifest_path(job_id: str) -> Path:
    return _job_dir(job_id) / "manifest.json"


def _atomic_write_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)


def _profile_signature(target_profile_id: str) -> str:
    profile = profile_registry.resolve_profile(target_profile_id)
    payload = json.dumps(profile.to_dict(), ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _assembly_profile_signature(assembly_profile) -> str:
    payload = json.dumps(asdict(assembly_profile), ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _load_manifest(job_id: str) -> Optional[dict]:
    path = _manifest_path(job_id)
    if not path.is_file():
        return None
    return json.loads(path.read_text(encoding="utf-8"))


def _save_manifest(job_id: str, manifest: dict) -> None:
    manifest["timestamps"] = manifest.get("timestamps", {})
    manifest["timestamps"]["updated_at"] = time.time()
    _atomic_write_json(_manifest_path(job_id), manifest)


def assemble_documents(
    source_paths: list[str],
    output_path: str,
    explicit_profile_id: Optional[str] = None,
    explicit_format_hint: Optional[str] = None,
    reference_profile_id: Optional[str] = None,
    saved_profile_id: Optional[str] = None,
    assembly_profile_id: Optional[str] = None,
    output_mode: str = "assembled",
    order_mode: str = "input",
    allow_default: bool = False,
    checkpoint: bool = False,
    job_id: Optional[str] = None,
    resume: bool = False,
) -> dict[str, Any]:
    """Unified batch assembly entry point. Returns an AssemblyResult dict."""
    sources = [Path(p).expanduser().resolve() for p in source_paths]
    sources = [p for p in sources if p.is_file()]
    if order_mode == "filename":
        sources = sorted(sources, key=lambda p: p.name.lower())

    output = Path(output_path).expanduser()
    if not output.is_absolute():
        output = output.resolve()

    if output in sources:
        result = AssemblyResult(status="error", total=len(sources))
        result.errors.append(ERROR_SOURCE_OUTPUT_PATH_CONFLICT)
        return result.to_dict()

    if assembly_profile_id:
        assembly_profile = profile_registry.get_assembly_profile(assembly_profile_id)
        if assembly_profile is None:
            result = AssemblyResult(status="error", total=len(sources))
            result.errors.append(ERROR_ASSEMBLY_PROFILE_NOT_FOUND)
            return result.to_dict()
    else:
        assembly_profile = _default_assembly_profile()

    result = AssemblyResult(
        status="ok",
        total=len(sources),
        assembly_profile=asdict(assembly_profile),
    )

    # 1. unified target format resolution
    resolution = _resolve_batch_format(
        sources,
        explicit_profile_id,
        explicit_format_hint,
        reference_profile_id,
        saved_profile_id,
        allow_default,
    )
    result.resolution = resolution
    if resolution.get("status") == "needs_guidance":
        result.status = "needs_guidance"
        result.warnings.extend(resolution.get("reason", []))
        return result.to_dict()
    if resolution.get("status") == "error":
        result.status = "error"
        result.errors.append(resolution.get("error") or "RESOLUTION_ERROR")
        return result.to_dict()

    target_profile_id = resolution.get("profile_id")
    if not target_profile_id:
        result.status = "error"
        result.errors.append("RESOLUTION_MISSING_PROFILE_ID")
        return result.to_dict()
    result.target_profile_id = target_profile_id

    # 2. checkpoint setup / resume validation
    manifest = None
    if resume:
        if not job_id:
            result.status = "error"
            result.errors.append("JOB_ID_REQUIRED_FOR_RESUME")
            return result.to_dict()
        manifest = _load_manifest(job_id)
        if manifest is None:
            result.status = "error"
            result.errors.append("JOB_NOT_FOUND")
            return result.to_dict()
        config_error = _validate_manifest_config(
            manifest,
            sources=sources,
            order_mode=order_mode,
            output_mode=output_mode,
            allow_default=allow_default,
            target_profile_id=target_profile_id,
            assembly_profile=assembly_profile,
            explicit_profile_id=explicit_profile_id,
            explicit_format_hint=explicit_format_hint,
            reference_profile_id=reference_profile_id,
            saved_profile_id=saved_profile_id,
        )
        if config_error:
            result.status = "error"
            result.errors.append("JOB_CONFIGURATION_CHANGED")
            result.warnings.append(config_error)
            return result.to_dict()
        for completed in manifest.get("completed", []):
            source = sources[completed["index"]]
            if _sha256_file(source) != completed["source_sha256"]:
                result.status = "error"
                result.errors.append(f"CHECKPOINT_SOURCE_CHANGED:{completed['index']}")
                result.warnings.append(f"source 文件已变化，checkpoint 失效: {source}")
                return result.to_dict()

    if checkpoint or resume:
        if not job_id:
            job_id = uuid.uuid4().hex[:12]
        job_dir = _job_dir(job_id)
        temp_dir = job_dir
        temp_dir.mkdir(parents=True, exist_ok=True)
        if manifest is None:
            manifest = {
                "schema_version": 1,
                "job_id": job_id,
                "status": "running",
                "sources": [{"path": str(s), "sha256": _sha256_file(s)} for s in sources],
                "target_profile_id": target_profile_id,
                "profile_signature": _profile_signature(target_profile_id),
                "assembly_profile_signature": _assembly_profile_signature(assembly_profile),
                "assembly_profile": asdict(assembly_profile),
                "output_mode": output_mode,
                "order_mode": order_mode,
                "allow_default": allow_default,
                "explicit_profile_id": explicit_profile_id,
                "explicit_format_hint": explicit_format_hint,
                "reference_profile_id": reference_profile_id,
                "saved_profile_id": saved_profile_id,
                "completed": [],
                "errors": [],
                "created_at": time.time(),
            }
            _save_manifest(job_id, manifest)
    else:
        temp_dir = Path(tempfile.mkdtemp(prefix="forge_assembly_"))

    normalize_profile_id = _normalize_profile_for(target_profile_id)
    try:
        completed_by_index = {c["index"]: c for c in manifest.get("completed", [])} if manifest else {}

        # 3. normalize each source (or reuse completed checkpoint)
        for index, source in enumerate(sources):
            item = AssemblyItemResult(
                index=index,
                source_path=str(source),
                source_sha256=_sha256_file(source),
                status="failed",
            )
            source_ir = read_docx(source)
            item.source_fingerprints = asdict(source_ir.content_fingerprint)

            if index in completed_by_index:
                completed = completed_by_index[index]
                normalized_path = temp_dir / completed["normalized_file"]
                if not normalized_path.is_file():
                    item.errors = ["CHECKPOINT_NORMALIZED_MISSING"]
                    result.items.append(item)
                    continue
                item.normalized_status = completed.get("normalized_status", "ok")
                item.normalized_fingerprints = completed.get("normalized_fingerprints", {})
                item.imported_block_count = completed.get("imported_block_count", 0)
                item.normalized_path = str(normalized_path)
                item.status = completed.get("status", "success")
                item.warnings = list(completed.get("warnings", []))
                result.items.append(item)
                continue

            normalized_path = temp_dir / f"normalized_{index:02d}.docx"
            single = reformat_single(
                source_path=str(source),
                output_path=str(normalized_path),
                explicit_profile_id=normalize_profile_id,
            )
            item.normalized_status = single.get("status", "error")
            if single.get("status") != "ok":
                item.errors = list(single.get("errors") or [])
                item.warnings = list(single.get("warnings") or [])
                result.items.append(item)
                if manifest is not None:
                    manifest["status"] = "incomplete"
                    manifest["errors"] = item.errors
                    _save_manifest(job_id, manifest)
                continue

            normalized_ir = read_docx(normalized_path)
            unsupported = _unsupported_package_check(normalized_path, normalized_ir)
            if unsupported:
                item.errors = unsupported
                result.items.append(item)
                if manifest is not None:
                    manifest["status"] = "incomplete"
                    manifest["errors"] = item.errors
                    _save_manifest(job_id, manifest)
                continue

            item.normalized_fingerprints = asdict(normalized_ir.content_fingerprint)
            item.imported_block_count = len(normalized_ir.blocks)
            item.normalized_path = str(normalized_path)
            item.status = "success_with_warnings" if single.get("warnings") else "success"
            item.warnings = list(single.get("warnings") or [])
            result.items.append(item)

            if manifest is not None:
                manifest.setdefault("completed", []).append(
                    {
                        "index": index,
                        "source_sha256": item.source_sha256,
                        "normalized_status": item.normalized_status,
                        "normalized_file": normalized_path.name,
                        "normalized_fingerprints": item.normalized_fingerprints,
                        "imported_block_count": item.imported_block_count,
                        "status": item.status,
                        "warnings": item.warnings,
                    }
                )
                _save_manifest(job_id, manifest)

        result.processed = sum(1 for item in result.items if item.status in ("success", "success_with_warnings"))
        result.failed = result.total - result.processed

        if result.failed > 0:
            result.status = "ASSEMBLY_INCOMPLETE"
            result.warnings.append(f"{result.failed} item(s) failed; whole-batch atomic, no output produced.")
            if manifest is not None:
                manifest["status"] = "incomplete"
                _save_manifest(job_id, manifest)
            return result.to_dict()

        # 4. build master package + import
        assembled_tmp = temp_dir / "assembled_tmp.docx"
        try:
            master_parts = _build_master_parts(
                [Path(item.normalized_path) for item in result.items], assembly_profile, target_profile_id
            )
            write_zip_parts(master_parts, assembled_tmp)
        except UnsupportedPackageContent as exc:
            result.status = "ASSEMBLY_INCOMPLETE"
            result.errors.append("IMPORT_UNSUPPORTED:" + str(exc))
            return result.to_dict()

        # 5. per-item preservation (normalized -> assembled)
        assembled_ir = read_docx(assembled_tmp)
        assembled_blocks = assembled_ir.blocks
        pos = 0
        page_break = bool(assembly_profile.page_break_between_items)
        for index, item in enumerate(result.items):
            count = item.imported_block_count
            slice_blocks = assembled_blocks[pos : pos + count]
            pos += count
            if page_break and index < len(result.items) - 1:
                pos += 1  # assembly-generated page-break separator
            normalized_ir = read_docx(Path(item.normalized_path))
            item.per_item_preservation = verify_item_preservation(normalized_ir, assembled_ir, slice_blocks)
            if not item.per_item_preservation.get("passed"):
                item.status = "failed"
                item.errors.append("ITEM_PRESERVATION_FAILED")

        result.processed = sum(1 for item in result.items if item.status in ("success", "success_with_warnings"))
        result.failed = result.total - result.processed
        if result.failed > 0:
            result.status = "ASSEMBLY_INCOMPLETE"
            result.warnings.append("Per-item preservation failed; whole-batch atomic, no output produced.")
            return result.to_dict()

        result.assembly_payload_sha256 = compute_assembly_payload_sha256([item.to_dict() for item in result.items])

        # 6. output modes
        if output_mode in ("assembled", "both"):
            output.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(assembled_tmp, output)
            result.output = str(output)
        if output_mode in ("separate", "both"):
            normalized_dir = output.parent / (output.stem + "_normalized")
            normalized_dir.mkdir(parents=True, exist_ok=True)
            for index, item in enumerate(result.items):
                dest_name = f"{index + 1:02d}_{Path(item.source_path).stem}.docx"
                dest = normalized_dir / dest_name
                shutil.copy2(Path(item.normalized_path), dest)
                result.normalized_outputs.append(str(dest))

        if manifest is not None:
            # success: keep manifest as completed metadata, remove temp package files
            manifest["status"] = "completed"
            _save_manifest(job_id, manifest)
            for tmp_name in ("assembled_tmp.docx",):
                p = temp_dir / tmp_name
                if p.exists():
                    p.unlink()
            normalized_dir_job = temp_dir / "normalized"
            if normalized_dir_job.exists():
                shutil.rmtree(normalized_dir_job, ignore_errors=True)
            for p in temp_dir.glob("normalized_*.docx"):
                p.unlink(missing_ok=True)
        else:
            shutil.rmtree(temp_dir, ignore_errors=True)

        result.status = "ok"
        if manifest is not None:
            result.warnings.append(f"job_id={job_id}")
        return result.to_dict()
    finally:
        if not checkpoint and not resume:
            shutil.rmtree(temp_dir, ignore_errors=True)


def _validate_manifest_config(
    manifest: dict,
    sources: list,
    order_mode: str,
    output_mode: str,
    allow_default: bool,
    target_profile_id: str,
    assembly_profile,
    explicit_profile_id,
    explicit_format_hint,
    reference_profile_id,
    saved_profile_id,
) -> str:
    reasons = []
    if manifest.get("sources") and [s["path"] for s in manifest["sources"]] != [str(s) for s in sources]:
        reasons.append("source list/order changed")
    if manifest.get("order_mode") != order_mode:
        reasons.append("order_mode changed")
    if manifest.get("output_mode") != output_mode:
        reasons.append("output_mode changed")
    if manifest.get("allow_default") != allow_default:
        reasons.append("allow_default changed")
    if manifest.get("target_profile_id") != target_profile_id:
        reasons.append("target profile changed")
    if manifest.get("profile_signature") != _profile_signature(target_profile_id):
        reasons.append("profile signature changed")
    if manifest.get("assembly_profile_signature") != _assembly_profile_signature(assembly_profile):
        reasons.append("assembly profile changed")
    for key in ("explicit_profile_id", "explicit_format_hint", "reference_profile_id", "saved_profile_id"):
        if manifest.get(key) != locals().get(key):
            reasons.append(f"{key} changed")
    return "; ".join(reasons)
