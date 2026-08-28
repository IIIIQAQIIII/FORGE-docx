"""Mission 05 — Package-aware DOCX importer.

Imports body content from normalized DOCX packages into a master package,
handling package-level conflicts:

- media part name collisions + r:embed / r:id remap
- external hyperlink relationships + r:id remap
- numbering abstractNumId / numId remap
- styleId collision-safe rename + reference patching
- wp:docPr/@id uniqueness
- bookmark w:id uniqueness

The importer never rebuilds tables or paragraphs; it deep-copies body XML and
only patches package references / IDs.
"""

from __future__ import annotations

import hashlib
from copy import deepcopy
from dataclasses import dataclass, field
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

from document_ir import A, R, V, W, WP, W_NS

PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
PKG_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PR = "{%s}" % PKG_REL_NS
CT = "{%s}" % PKG_CT_NS

IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
HYPERLINK_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
NUMBERING_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
NUMBERING_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"

IMAGE_CT_BY_EXT = {
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "tiff": "image/tiff",
    "svg": "image/svg+xml",
}


class UnsupportedPackageContent(Exception):
    pass


@dataclass
class ImportState:
    parts: dict
    rels_name: str
    rels_root: object
    body: object
    document_root: object = None
    next_rid: int = 0
    next_docpr: int = 0
    next_bookmark: int = 0
    next_abstract: int = 0
    next_num: int = 0
    next_media: int = 0
    style_remap: dict = field(default_factory=dict)
    style_seq: int = 0
    warnings: list = field(default_factory=list)
    item_media_shas: list = field(default_factory=list)
    item_media_part_names: list = field(default_factory=list)


def read_zip_parts(path) -> dict:
    with ZipFile(path) as archive:
        return {info.filename: archive.read(info.filename) for info in archive.infolist()}


def write_zip_parts(parts: dict, path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(path, "w", ZIP_DEFLATED) as zout:
        for name, data in parts.items():
            zout.writestr(name, data)


def _serialize_xml(root) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _parse_xml(data: bytes):
    return etree.fromstring(data)


def _save_rels(state: ImportState) -> None:
    state.parts[state.rels_name] = _serialize_xml(state.rels_root)


def _next_rid(state: ImportState) -> str:
    state.next_rid += 1
    return f"rId{state.next_rid}"


def _add_rel(state: ImportState, rel_type: str, target: str, target_mode: str = None) -> str:
    rid = _next_rid(state)
    rel = etree.SubElement(state.rels_root, PR + "Relationship")
    rel.set("Id", rid)
    rel.set("Type", rel_type)
    rel.set("Target", target)
    if target_mode:
        rel.set("TargetMode", target_mode)
    _save_rels(state)
    return rid


def _full_part_name(target: str) -> str:
    target = (target or "").strip()
    if target.startswith("word/"):
        return target
    if target.startswith("/"):
        return "word/" + target.lstrip("/")
    return "word/" + target


def _ensure_content_type_default(parts: dict, ext: str) -> None:
    ct_name = "[Content_Types].xml"
    root = _parse_xml(parts[ct_name])
    ext_l = ext.lower()
    for default in root.findall(CT + "Default"):
        if (default.get("Extension") or "").lower() == ext_l:
            return
    content_type = IMAGE_CT_BY_EXT.get(ext_l)
    if content_type is None:
        return
    default = etree.SubElement(root, CT + "Default")
    default.set("Extension", ext_l)
    default.set("ContentType", content_type)
    parts[ct_name] = _serialize_xml(root)


def _ensure_content_type_override(parts: dict, part_name: str, content_type: str) -> None:
    ct_name = "[Content_Types].xml"
    root = _parse_xml(parts[ct_name])
    part_name_l = "/" + part_name
    for override in root.findall(CT + "Override"):
        if override.get("PartName") == part_name_l:
            return
    override = etree.SubElement(root, CT + "Override")
    override.set("PartName", part_name_l)
    override.set("ContentType", content_type)
    parts[ct_name] = _serialize_xml(root)


def _source_rels_map(source_parts: dict) -> dict:
    rels_name = "word/_rels/document.xml.rels"
    raw = source_parts.get(rels_name)
    result = {}
    if raw is None:
        return result
    root = _parse_xml(raw)
    for rel in root.findall(PR + "Relationship"):
        result[rel.get("Id")] = (rel.get("Type"), rel.get("Target"), rel.get("TargetMode"))
    return result


def _import_media(state: ImportState, source_parts: dict, target: str, item_index: int) -> str:
    full = _full_part_name(target)
    if full not in source_parts:
        raise UnsupportedPackageContent(f"media part not found: {full}")
    data = source_parts[full]
    sha = hashlib.sha256(data).hexdigest()
    ext = (full.rsplit(".", 1)[-1] if "." in full else "bin").lower()
    new_name = f"word/media/asm_{state.next_media}_{item_index}.{ext}"
    state.next_media += 1
    state.parts[new_name] = data
    _ensure_content_type_default(state.parts, ext)
    rel_target = new_name[len("word/") :]
    rid = _add_rel(state, IMAGE_REL_TYPE, rel_target)
    state.item_media_shas.append(sha)
    state.item_media_part_names.append(new_name)
    return rid


def _remap_relationships(copy, state: ImportState, source_parts: dict, source_rels: dict, item_index: int) -> None:
    for element in copy.iter():
        for attr in (R + "embed", R + "id", R + "link"):
            old_rid = element.get(attr)
            if not old_rid:
                continue
            rel = source_rels.get(old_rid)
            if rel is None:
                state.warnings.append("UNMAPPED_RELATIONSHIP_REFERENCE:" + old_rid)
                continue
            rel_type, target, target_mode = rel
            if "image" in rel_type:
                element.set(attr, _import_media(state, source_parts, target, item_index))
            elif "hyperlink" in rel_type:
                element.set(attr, _add_rel(state, HYPERLINK_REL_TYPE, target, target_mode))
            else:
                raise UnsupportedPackageContent(f"unsupported body relationship type: {rel_type}")


def _strip_paragraph_sectpr(copy) -> None:
    for sect_pr in list(copy.iter(W + "sectPr")):
        parent = sect_pr.getparent()
        if parent is not None and parent.tag == W + "pPr":
            parent.remove(sect_pr)


def _import_numbering(state: ImportState, source_parts: dict) -> dict:
    """Import source numbering definitions; return old numId -> new numId map."""
    source_raw = source_parts.get("word/numbering.xml")
    if source_raw is None:
        return {}
    source_root = _parse_xml(source_raw)

    if "word/numbering.xml" not in state.parts:
        master_root = etree.fromstring(
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:numbering xmlns:w="{W_NS}"></w:numbering>'.encode("utf-8")
        )
        state.parts["word/numbering.xml"] = _serialize_xml(master_root)
        _ensure_content_type_override(state.parts, "word/numbering.xml", NUMBERING_CT)
    else:
        master_root = _parse_xml(state.parts["word/numbering.xml"])

    abs_map = {}
    for abstract in source_root.findall(W + "abstractNum"):
        old_abs = abstract.get(W + "abstractNumId")
        new_abs = str(state.next_abstract)
        state.next_abstract += 1
        copied = deepcopy(abstract)
        copied.set(W + "abstractNumId", new_abs)
        master_root.append(copied)
        abs_map[old_abs] = new_abs

    num_map = {}
    for num in source_root.findall(W + "num"):
        old_num = num.get(W + "numId")
        new_num = str(state.next_num)
        state.next_num += 1
        copied = deepcopy(num)
        copied.set(W + "numId", new_num)
        abs_ref = copied.find(W + "abstractNumId")
        if abs_ref is not None and abs_ref.get(W + "val") in abs_map:
            abs_ref.set(W + "val", abs_map[abs_ref.get(W + "val")])
        master_root.append(copied)
        num_map[old_num] = new_num

    state.parts["word/numbering.xml"] = _serialize_xml(master_root)
    return num_map


def _patch_numbering_refs(copy, num_map: dict) -> None:
    if not num_map:
        return
    for num_id in copy.iter(W + "numId"):
        val = num_id.get(W + "val")
        if val in num_map:
            num_id.set(W + "val", num_map[val])


def _style_key(style_element) -> str:
    clone = deepcopy(style_element)
    clone.attrib.pop(W + "styleId", None)
    return etree.tostring(clone)


def _collect_style_refs(element) -> set:
    used = set()
    for child in element.iter():
        for tag in (W + "pStyle", W + "rStyle", W + "tblStyle"):
            val = child.get(W + "val")
            if val:
                used.add(val)
    return used


def _import_styles(state: ImportState, source_parts: dict, used_ids: set) -> dict:
    """Import styles used by the current source. Returns old styleId -> new styleId."""
    if not used_ids:
        return {}
    source_raw = source_parts.get("word/styles.xml")
    source_root = _parse_xml(source_raw) if source_raw is not None else None
    source_by_id = {}
    if source_root is not None:
        for style in source_root.findall(W + "style"):
            source_by_id[style.get(W + "styleId")] = style

    master_raw = state.parts.get("word/styles.xml")
    master_root = _parse_xml(master_raw) if master_raw is not None else etree.fromstring(
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{W_NS}"></w:styles>'.encode("utf-8")
    )
    master_by_id = {}
    for style in master_root.findall(W + "style"):
        master_by_id[style.get(W + "styleId")] = style

    ordered = []
    visited = set()
    style_map = {}

    def visit(style_id: str) -> None:
        if style_id in visited or style_id in style_map:
            return
        visited.add(style_id)
        style_el = source_by_id.get(style_id)
        if style_el is None:
            state.warnings.append("STYLE_NOT_FOUND_IN_SOURCE:" + style_id)
            return
        based_on = style_el.find(W + "basedOn")
        if based_on is not None and based_on.get(W + "val"):
            visit(based_on.get(W + "val"))
        ordered.append(style_id)

    for style_id in sorted(used_ids):
        visit(style_id)

    for style_id in ordered:
        style_el = source_by_id[style_id]
        if style_id in master_by_id:
            if _style_key(master_by_id[style_id]) == _style_key(style_el):
                style_map[style_id] = style_id
            else:
                new_id = f"Assembly_{style_id}_{state.style_seq}"
                state.style_seq += 1
                copied = deepcopy(style_el)
                copied.set(W + "styleId", new_id)
                based_on = copied.find(W + "basedOn")
                if based_on is not None and based_on.get(W + "val") in style_map:
                    based_on.set(W + "val", style_map[based_on.get(W + "val")])
                master_root.append(copied)
                master_by_id[new_id] = copied
                style_map[style_id] = new_id
        else:
            copied = deepcopy(style_el)
            copied.set(W + "styleId", style_id)
            based_on = copied.find(W + "basedOn")
            if based_on is not None and based_on.get(W + "val") in style_map:
                based_on.set(W + "val", style_map[based_on.get(W + "val")])
            master_root.append(copied)
            master_by_id[style_id] = copied
            style_map[style_id] = style_id

    state.parts["word/styles.xml"] = _serialize_xml(master_root)
    return style_map


def _patch_style_refs(copy, style_map: dict) -> None:
    for child in copy.iter():
        for tag in (W + "pStyle", W + "rStyle", W + "tblStyle"):
            val = child.get(W + "val")
            if val and val in style_map:
                child.set(W + "val", style_map[val])


def _remap_docpr_ids(copy, state: ImportState) -> None:
    for doc_pr in copy.iter(WP + "docPr"):
        doc_pr.set("id", str(state.next_docpr))
        state.next_docpr += 1


def _remap_bookmark_ids(copy, state: ImportState) -> None:
    local_map = {}
    for tag in (W + "bookmarkStart", W + "bookmarkEnd"):
        for bookmark in copy.iter(tag):
            old_id = bookmark.get(W + "id")
            if old_id is None:
                continue
            if old_id not in local_map:
                local_map[old_id] = str(state.next_bookmark)
                state.next_bookmark += 1
            bookmark.set(W + "id", local_map[old_id])


def _append_body_child(body, child) -> None:
    children = list(body)
    insert_at = len(children)
    for index, existing in enumerate(children):
        if existing.tag == W + "sectPr":
            insert_at = index
            break
    body.insert(insert_at, child)


def make_separator() -> object:
    return etree.fromstring(
        f'<w:p xmlns:w="{W_NS}"><w:pPr><w:rPr/></w:pPr>'
        f'<w:r><w:br w:type="page"/></w:r></w:p>'.encode("utf-8")
    )


def import_item_body(state: ImportState, source_path, item_index: int) -> int:
    """Import one normalized DOCX into the master body. Returns imported child count."""
    source_parts = read_zip_parts(source_path)
    source_root = _parse_xml(source_parts["word/document.xml"])
    source_body = source_root.find(W + "body")
    if source_body is None:
        raise UnsupportedPackageContent("source has no w:body")

    source_rels = _source_rels_map(source_parts)
    num_map = _import_numbering(state, source_parts)

    used_ids = set()
    body_children = []
    for child in source_body:
        if child.tag == W + "sectPr":
            continue
        body_children.append(child)
        used_ids |= _collect_style_refs(child)

    style_map = _import_styles(state, source_parts, used_ids)

    state.item_media_shas = []
    state.item_media_part_names = []

    imported = 0
    for child in body_children:
        copy = deepcopy(child)
        _strip_paragraph_sectpr(copy)
        _remap_relationships(copy, state, source_parts, source_rels, item_index)
        _patch_style_refs(copy, style_map)
        _patch_numbering_refs(copy, num_map)
        _remap_docpr_ids(copy, state)
        _remap_bookmark_ids(copy, state)
        _append_body_child(state.body, copy)
        imported += 1
    return imported


def _collect_max_docpr_id(document_root) -> int:
    max_id = 0
    for doc_pr in document_root.iter(WP + "docPr"):
        try:
            max_id = max(max_id, int(doc_pr.get("id") or "0"))
        except ValueError:
            pass
    return max_id


def _collect_max_bookmark_id(document_root) -> int:
    max_id = 0
    for tag in (W + "bookmarkStart", W + "bookmarkEnd"):
        for bookmark in document_root.iter(tag):
            try:
                max_id = max(max_id, int(bookmark.get(W + "id") or "0"))
            except ValueError:
                pass
    return max_id


def init_import_state(master_parts: dict) -> ImportState:
    document_root = _parse_xml(master_parts["word/document.xml"])
    body = document_root.find(W + "body")
    rels_name, rels_root = _parse_relationships(master_parts)
    max_rid = _max_rid(rels_root)
    return ImportState(
        parts=master_parts,
        rels_name=rels_name,
        rels_root=rels_root,
        body=body,
        document_root=document_root,
        next_rid=max_rid,
        next_docpr=_collect_max_docpr_id(document_root) + 1,
        next_bookmark=_collect_max_bookmark_id(document_root) + 1,
        next_abstract=_max_abstract_id(master_parts) + 1,
        next_num=_max_num_id(master_parts) + 1,
    )


def _parse_relationships(parts: dict):
    rels_name = "word/_rels/document.xml.rels"
    raw = parts.get(rels_name)
    if raw is not None:
        return rels_name, _parse_xml(raw)
    root = etree.fromstring(
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL_NS}"></Relationships>'.encode("utf-8")
    )
    parts[rels_name] = _serialize_xml(root)
    return rels_name, root


def _max_rid(rels_root) -> int:
    max_n = 0
    for rel in rels_root.findall(PR + "Relationship"):
        rid = rel.get("Id") or ""
        if rid.startswith("rId"):
            try:
                max_n = max(max_n, int(rid[3:]))
            except ValueError:
                pass
    return max_n


def _max_abstract_id(parts: dict) -> int:
    raw = parts.get("word/numbering.xml")
    if raw is None:
        return -1
    root = _parse_xml(raw)
    max_n = -1
    for abstract in root.findall(W + "abstractNum"):
        try:
            max_n = max(max_n, int(abstract.get(W + "abstractNumId") or "0"))
        except ValueError:
            pass
    return max_n


def _max_num_id(parts: dict) -> int:
    raw = parts.get("word/numbering.xml")
    if raw is None:
        return -1
    root = _parse_xml(raw)
    max_n = -1
    for num in root.findall(W + "num"):
        try:
            max_n = max(max_n, int(num.get(W + "numId") or "0"))
        except ValueError:
            pass
    return max_n


def save_document_from_parts(parts: dict) -> None:
    document_root = _parse_xml(parts["word/document.xml"])
    parts["word/document.xml"] = _serialize_xml(document_root)
