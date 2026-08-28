"""Dynamic Template Registry stored under FORGE_HOME/templates."""

from __future__ import annotations

import json
import re
import shutil
from pathlib import Path
from typing import Any, Optional
from zipfile import BadZipFile, ZipFile

from open_format.home import templates_dir
from profiles import registry as profile_registry

SCHEMA_VERSION = 1

ERROR_TEMPLATE_ID_CONFLICT = "TEMPLATE_ID_CONFLICT"
ERROR_TEMPLATE_PROFILE_NOT_FOUND = "PROFILE_NOT_FOUND"
ERROR_TEMPLATE_NOT_FOUND = "TEMPLATE_NOT_FOUND"
ERROR_NO_TEMPLATE_PLACEHOLDERS = "NO_TEMPLATE_PLACEHOLDERS"
ERROR_MACRO_REJECTED = "MACRO_REJECTED"
ERROR_UNSAFE_EXTERNAL_RELATIONSHIPS = "UNSAFE_EXTERNAL_RELATIONSHIPS"
ERROR_MALFORMED_DOCX = "MALFORMED_DOCX"
ERROR_ALIAS_CONFLICT = "ALIAS_CONFLICT"

_PLACEHOLDER_RE = re.compile(r"\{\{([^{}]+)\}\}")
_LOOP_RE = re.compile(r"\{%p\s*for\s+([A-Za-z_]\w*)\s+in\s+([^%]+?)\s*%\}")


def _builtin_aliases() -> set:
    try:
        from server import DOCUMENT_TYPES

        return set(DOCUMENT_TYPES.keys())
    except Exception:  # noqa: BLE001 - registry can run without server loaded
        return set()


def _manifest_path(template_id: str) -> Path:
    return templates_dir() / f"{template_id}.json"


def _sanitize_id(template_id: str) -> bool:
    allowed = set("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_.-")
    return bool(template_id) and all(ch in allowed for ch in template_id)


def _discover_placeholders(path: Path) -> list:
    found = []
    loop_vars = set()
    with ZipFile(path) as archive:
        xml_parts = ["word/document.xml"] + [
            name for name in archive.namelist() if re.match(r"word/(header|footer)\d*\.xml", name)
        ]
        for part in xml_parts:
            try:
                xml = archive.read(part).decode("utf-8")
            except KeyError:
                continue
            for match in _PLACEHOLDER_RE.finditer(xml):
                expr = match.group(1).strip()
                if expr and expr not in found:
                    found.append(expr)
            for match in _LOOP_RE.finditer(xml):
                var = match.group(1).strip()
                seq = match.group(2).strip()
                loop_vars.add(var)
                if re.fullmatch(r"[A-Za-z_]\w*", seq) and seq not in found:
                    found.append(seq)
    return [field for field in found if field not in loop_vars]


def _validate_package(path: Path) -> list:
    errors = []
    try:
        with ZipFile(path) as archive:
            names = set(archive.namelist())
            if "word/document.xml" not in names:
                errors.append(ERROR_MALFORMED_DOCX + ":missing word/document.xml")
            if "word/vbaProject.bin" in names or path.suffix.lower() == ".docm":
                errors.append(ERROR_MACRO_REJECTED)
            for rels_name in [n for n in names if n.endswith(".rels")]:
                try:
                    from lxml import etree

                    root = etree.fromstring(archive.read(rels_name))
                except Exception:
                    continue
                ns = "http://schemas.openxmlformats.org/package/2006/relationships"
                for rel in root.findall(f"{{{ns}}}Relationship"):
                    rel_type = rel.get("Type") or ""
                    target_mode = rel.get("TargetMode")
                    if target_mode == "External":
                        if "hyperlink" not in rel_type:
                            errors.append(ERROR_UNSAFE_EXTERNAL_RELATIONSHIPS + f":{rels_name}:{rel.get('Id')}")
    except (BadZipFile, OSError):
        errors.append(ERROR_MALFORMED_DOCX)
    return errors


def register_document_template(
    template_path: str | Path,
    template_id: str,
    name: str,
    kind: str,
    profile_id: str,
    supported_intents: Optional[list] = None,
    aliases: Optional[list] = None,
) -> dict:
    if not _sanitize_id(template_id):
        return {"status": "error", "error": ERROR_TEMPLATE_ID_CONFLICT, "reason": "template_id 仅允许字母/数字/_.-"}
    if _manifest_path(template_id).exists():
        return {"status": "error", "error": ERROR_TEMPLATE_ID_CONFLICT, "reason": f"template_id 已存在: {template_id}"}

    source = Path(template_path).expanduser()
    if not source.is_file():
        return {"status": "error", "error": ERROR_TEMPLATE_NOT_FOUND, "reason": f"template 文件不存在: {source}"}

    package_errors = _validate_package(source)
    if package_errors:
        return {"status": "error", "error": package_errors[0], "reason": package_errors}

    try:
        profile_registry.resolve_profile(profile_id)
    except KeyError:
        return {"status": "error", "error": ERROR_TEMPLATE_PROFILE_NOT_FOUND, "reason": f"profile_id 不存在: {profile_id}"}

    if kind == "docxtpl":
        placeholders = _discover_placeholders(source)
        if not placeholders:
            return {
                "status": "error",
                "error": ERROR_NO_TEMPLATE_PLACEHOLDERS,
                "reason": "docxtpl template 未发现任何 placeholders；可改为 kind=reference 注册。",
            }
    elif kind == "reference":
        placeholders = _discover_placeholders(source)
    else:
        return {"status": "error", "error": "INVALID_KIND", "reason": "kind 必须是 docxtpl|reference"}

    aliases = [a for a in (aliases or []) if a]
    builtin_aliases = _builtin_aliases()
    dangerous = [a for a in aliases if a in builtin_aliases]
    if dangerous:
        return {
            "status": "error",
            "error": ERROR_ALIAS_CONFLICT,
            "reason": f"alias 与 built-in 冲突: {dangerous}",
        }

    dest_dir = templates_dir()
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest_file = dest_dir / f"{template_id}.docx"
    shutil.copy2(source, dest_file)

    manifest = {
        "schema_version": SCHEMA_VERSION,
        "template_id": template_id,
        "name": name,
        "kind": kind,
        "file": dest_file.name,
        "profile_id": profile_id,
        "supported_intents": supported_intents or [],
        "aliases": aliases,
        "schema": {"placeholders": placeholders} if kind == "docxtpl" else {},
    }
    _write_manifest(manifest)
    return {"status": "ok", "template": manifest}


def _write_manifest(manifest: dict) -> None:
    path = _manifest_path(manifest["template_id"])
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)


def load_template_manifest(template_id: str) -> Optional[dict]:
    path = _manifest_path(template_id)
    if not path.is_file():
        return None
    return json.loads(path.read_text(encoding="utf-8"))


def list_user_templates() -> list[dict]:
    templates = []
    for path in sorted(templates_dir().glob("*.json")):
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            if data.get("schema_version") == 1:
                templates.append(data)
        except Exception:
            continue
    return templates


def _builtin_template_entries() -> list[dict]:
    try:
        from server import DOCUMENT_TYPES

        entries_by_file = {}
        for type_name, template_file in DOCUMENT_TYPES.items():
            entries_by_file.setdefault(template_file, []).append(type_name)
        entries = []
        for template_file, aliases in entries_by_file.items():
            entries.append(
                {
                    "template_id": Path(template_file).stem,
                    "name": Path(template_file).stem,
                    "kind": "docxtpl",
                    "file": template_file,
                    "profile_id": _builtin_profile_for(template_file),
                    "supported_intents": [],
                    "aliases": aliases,
                    "origin": "builtin",
                }
            )
        return entries
    except Exception:  # noqa: BLE001
        return []


def _builtin_profile_for(template_file: str) -> Optional[str]:
    mapping = {
        "传统公文.docx": "official_standard",
        "传统公文-活动方案.docx": "activity_plan_standard",
        "论文.docx": "academic_standard",
        "行政周报.docx": "weekly_standard",
        "活动总结.docx": "activity_summary_standard",
        "活动影像.docx": "activity_archive_standard",
        "培训通知.docx": "training_notice_standard",
        "培训活动记录.docx": "training_record_standard",
        "培训活动影像.docx": "training_archive_standard",
        "培训通知记录.docx": "training_notice_standard",
        "sample_template.docx": "generic_document",
    }
    return mapping.get(template_file)


def list_document_templates() -> list[dict]:
    entries = _builtin_template_entries()
    for manifest in list_user_templates():
        entries.append(
            {
                "template_id": manifest["template_id"],
                "name": manifest["name"],
                "kind": manifest["kind"],
                "file": manifest["file"],
                "profile_id": manifest["profile_id"],
                "supported_intents": manifest.get("supported_intents", []),
                "aliases": manifest.get("aliases", []),
                "origin": "user",
            }
        )
    return entries


def resolve_document_template(document_type: str) -> Optional[dict]:
    """Resolve a type name / template id to builtin legacy or user template."""
    # user alias or id first (data-driven layer above legacy)
    for manifest in list_user_templates():
        if document_type in (manifest.get("aliases") or []) or document_type == manifest["template_id"]:
            return {"origin": "user", "manifest": manifest, "file": templates_dir() / manifest["file"]}
    try:
        from server import DOCUMENT_TYPES

        if document_type in DOCUMENT_TYPES:
            return {"origin": "builtin", "template": DOCUMENT_TYPES[document_type]}
    except Exception:  # noqa: BLE001
        pass
    return None
