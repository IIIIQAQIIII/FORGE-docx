"""Mission 04-B2 — Source-Preserving DOCX Renderer (Table & Image Formatting)."""

from __future__ import annotations

import hashlib
import os
import shutil
import tempfile
from pathlib import Path
from typing import Any, Optional
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

from document_ir import A, R, V, W, WP, W_NS, DocumentIR, read_docx
from profiles import registry as profile_registry
from reformat_engine.models import ReformatPlan

SOURCE_CHANGED = "SOURCE_CHANGED"
SOURCE_OUTPUT_PATH_CONFLICT = "SOURCE_OUTPUT_PATH_CONFLICT"
CONTENT_PRESERVATION_FAILED = "CONTENT_PRESERVATION_FAILED"
PROFILE_NOT_FOUND = "PROFILE_NOT_FOUND"
PLAN_NOT_READY = "PLAN_NOT_READY"

ALIGN_MAP = {
    "center": "center",
    "right": "right",
    "left": "left",
    "both": "both",
    "justify": "both",
}

EMU_PER_CM = 360000

PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
PKG_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PR = "{%s}" % PKG_REL_NS
CT = "{%s}" % PKG_CT_NS

FTR_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
HDR_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
FTR_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
HDR_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"

CH = chr(10)


def _locate(body_element, locator: str):
    if not locator:
        return None
    parts = locator.split("/")
    if parts[0] != "body":
        return None
    current = body_element
    for part in parts[1:]:
        if part == "sdtContent":
            current = current.find(W + "sdtContent")
        else:
            try:
                index = int(part)
            except ValueError:
                return None
            children = [c for c in current]
            if index < 0 or index >= len(children):
                return None
            current = children[index]
    return current


def _set_pPr_child(p_pr, tag, attrs):
    element = p_pr.find(tag)
    if element is None:
        element = etree.SubElement(p_pr, tag)
    for key, value in attrs.items():
        element.set(key, value)
    return element


def _apply_run_format(run, font=None, size_pt=None, bold=None, italic=None):
    if font is None and size_pt is None and bold is None and italic is None:
        return
    r_pr = run.find(W + "rPr")
    if r_pr is None:
        r_pr = etree.Element(W + "rPr")
        run.insert(0, r_pr)
    if font is not None:
        rfonts = r_pr.find(W + "rFonts")
        if rfonts is None:
            rfonts = etree.SubElement(r_pr, W + "rFonts")
        rfonts.set(W + "ascii", font)
        rfonts.set(W + "hAnsi", font)
        rfonts.set(W + "eastAsia", font)
    if size_pt is not None:
        half = str(int(round(float(size_pt) * 2)))
        for tag in (W + "sz", W + "szCs"):
            sz = r_pr.find(tag)
            if sz is None:
                sz = etree.SubElement(r_pr, tag)
            sz.set(W + "val", half)
    if bold is not None:
        b = r_pr.find(W + "b")
        if b is None:
            b = etree.SubElement(r_pr, W + "b")
        b.set(W + "val", "1" if bold else "0")
    if italic is not None:
        i = r_pr.find(W + "i")
        if i is None:
            i = etree.SubElement(r_pr, W + "i")
        i.set(W + "val", "1" if italic else "0")


def _apply_paragraph_format(p_element, slot: dict[str, Any]) -> None:
    if not slot:
        return
    p_pr = p_element.find(W + "pPr")
    if p_pr is None:
        p_pr = etree.Element(W + "pPr")
        p_element.insert(0, p_pr)
    align = slot.get("align")
    if align in ALIGN_MAP:
        _set_pPr_child(p_pr, W + "jc", {W + "val": ALIGN_MAP[align]})
    if "line_spacing_pt" in slot:
        spacing = _set_pPr_child(
            p_pr,
            W + "spacing",
            {
                W + "line": str(int(round(float(slot["line_spacing_pt"]) * 20))),
                W + "lineRule": "exact",
            },
        )
        if "space_before_pt" in slot:
            spacing.set(W + "before", str(int(round(float(slot["space_before_pt"]) * 20))))
        if "space_after_pt" in slot:
            spacing.set(W + "after", str(int(round(float(slot["space_after_pt"]) * 20))))
    indent_updates = {}
    if "first_line_chars" in slot:
        indent_updates[W + "firstLineChars"] = str(slot["first_line_chars"])
        if "first_line_twips" in slot:
            indent_updates[W + "firstLine"] = str(slot["first_line_twips"])
    if "left_chars" in slot:
        indent_updates[W + "leftChars"] = str(slot["left_chars"])
        if "left_twips" in slot:
            indent_updates[W + "left"] = str(slot["left_twips"])
    if "right_chars" in slot:
        indent_updates[W + "rightChars"] = str(slot["right_chars"])
        if "right_twips" in slot:
            indent_updates[W + "right"] = str(slot["right_twips"])
    if indent_updates:
        _set_pPr_child(p_pr, W + "ind", indent_updates)
    font = slot.get("font")
    size_pt = slot.get("size_pt")
    bold = slot.get("bold")
    italic = slot.get("italic")
    for run in p_element.findall(W + "r"):
        _apply_run_format(run, font, size_pt, bold, italic)
    for hyperlink in p_element.findall(W + "hyperlink"):
        for run in hyperlink.findall(W + "r"):
            _apply_run_format(run, font, size_pt, bold, italic)


def _apply_page_format(body_element, page: dict[str, Any]) -> None:
    if not page:
        return
    sect_pr = body_element.find(W + "sectPr")
    if sect_pr is None:
        return
    pg_sz = sect_pr.find(W + "pgSz")
    if pg_sz is None:
        pg_sz = etree.SubElement(sect_pr, W + "pgSz")
    if "width_cm" in page:
        pg_sz.set(W + "w", str(int(round(float(page["width_cm"]) * 567))))
    if "height_cm" in page:
        pg_sz.set(W + "h", str(int(round(float(page["height_cm"]) * 567))))
    pg_mar = sect_pr.find(W + "pgMar")
    if pg_mar is None:
        pg_mar = etree.SubElement(sect_pr, W + "pgMar")
    mar_map = {
        "top_cm": W + "top",
        "bottom_cm": W + "bottom",
        "left_cm": W + "left",
        "right_cm": W + "right",
        "header_distance_cm": W + "header",
        "footer_distance_cm": W + "footer",
    }
    for key, attr in mar_map.items():
        if key in page:
            pg_mar.set(attr, str(int(round(float(page[key]) * 567))))


# ---------------------------------------------------------------------------
# Document chrome: source-preserving PAGE field handling
# ---------------------------------------------------------------------------

def _find_sect_prs(body_element):
    sect_prs = []
    for p in body_element.findall(W + "p"):
        p_pr = p.find(W + "pPr")
        if p_pr is not None:
            sect_pr = p_pr.find(W + "sectPr")
            if sect_pr is not None:
                sect_prs.append(sect_pr)
    body_sect_pr = body_element.find(W + "sectPr")
    if body_sect_pr is not None:
        sect_prs.append(body_sect_pr)
    return sect_prs


def _serialize_xml(root) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _parse_relationships(parts) -> tuple:
    rels_name = "word/_rels/document.xml.rels"
    raw = parts.get(rels_name)
    if raw is not None:
        root = etree.fromstring(raw)
    else:
        root = etree.fromstring(
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="{PKG_REL_NS}"></Relationships>'.encode("utf-8")
        )
    return rels_name, root


def _next_rid(rels_root) -> str:
    max_n = 0
    for rel in rels_root.findall(PR + "Relationship"):
        rid = rel.get("Id") or ""
        if rid.startswith("rId"):
            try:
                max_n = max(max_n, int(rid[3:]))
            except ValueError:
                pass
    return f"rId{max_n + 1}"


def _add_relationship(parts, rels_name, rels_root, rel_type: str, target: str) -> str:
    rid = _next_rid(rels_root)
    rel = etree.SubElement(rels_root, PR + "Relationship")
    rel.set("Id", rid)
    rel.set("Type", rel_type)
    rel.set("Target", target)
    parts[rels_name] = _serialize_xml(rels_root)
    return rid


def _ensure_content_type(parts, part_name: str, content_type: str) -> None:
    ct_name = "[Content_Types].xml"
    raw = parts.get(ct_name)
    if raw is None:
        parts[ct_name] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Types xmlns="{PKG_CT_NS}"><Default Extension="rels" '
            'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/></Types>'
        ).encode("utf-8")
        raw = parts[ct_name]
    root = etree.fromstring(raw)
    part_name_l = "/" + part_name
    for override in root.findall(CT + "Override"):
        if override.get("PartName") == part_name_l:
            return
    override = etree.SubElement(root, CT + "Override")
    override.set("PartName", part_name_l)
    override.set("ContentType", content_type)
    parts[ct_name] = _serialize_xml(root)


def _next_part_index(parts, prefix: str) -> int:
    import re as _re

    max_n = 0
    pattern = _re.compile(r"^word/(footer|header)(\d+)\.xml$")
    for name in parts:
        match = pattern.match(name)
        if match and match.group(1) == prefix:
            max_n = max(max_n, int(match.group(2)))
    return max_n + 1


def _page_field_paragraph_xml(alignment, font, size_pt) -> str:
    jc = ALIGN_MAP.get(alignment)
    ppr = f'<w:pPr><w:jc w:val="{jc}"/>' if jc else "<w:pPr>"
    rpr = "<w:rPr>"
    if font:
        rpr += (
            f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:eastAsia="{font}"/>'
        )
    if size_pt is not None:
        half = int(round(float(size_pt) * 2))
        rpr += f'<w:sz w:val="{half}"/><w:szCs w:val="{half}"/>'
    rpr += "</w:rPr>"
    ppr += f"{rpr}</w:pPr>"
    return (
        f'<w:p xmlns:w="{W_NS}">{ppr}'
        f"<w:r>{rpr}<w:fldChar w:fldCharType=\"begin\"/></w:r>"
        f"<w:r>{rpr}<w:instrText xml:space=\"preserve\"> PAGE </w:instrText></w:r>"
        f"<w:r>{rpr}<w:fldChar w:fldCharType=\"end\"/></w:r>"
        f"</w:p>"
    )


def _make_chrome_part_xml(position: str, alignment, font, size_pt, with_page_field: bool) -> bytes:
    root_tag = "w:ftr" if position == "footer" else "w:hdr"
    page_paragraph = _page_field_paragraph_xml(alignment, font, size_pt) if with_page_field else ""
    if not page_paragraph:
        page_paragraph = "<w:p/>"
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<{root_tag} xmlns:w="{W_NS}">{page_paragraph}</{root_tag}>'
    ).encode("utf-8")


def _paragraph_has_page_field(p_element) -> bool:
    for instr in p_element.iter(W + "instrText"):
        if (instr.text or "").strip() == "PAGE":
            return True
    return False


def _apply_rpr(element, font, size_pt) -> None:
    if font is None and size_pt is None:
        return
    r_pr = element.find(W + "rPr")
    if r_pr is None:
        r_pr = etree.Element(W + "rPr")
        element.insert(0, r_pr)
    if font is not None:
        rfonts = r_pr.find(W + "rFonts")
        if rfonts is None:
            rfonts = etree.SubElement(r_pr, W + "rFonts")
        rfonts.set(W + "ascii", font)
        rfonts.set(W + "hAnsi", font)
        rfonts.set(W + "eastAsia", font)
    if size_pt is not None:
        half = str(int(round(float(size_pt) * 2)))
        for tag in (W + "sz", W + "szCs"):
            sz = r_pr.find(tag)
            if sz is None:
                sz = etree.SubElement(r_pr, tag)
            sz.set(W + "val", half)


def _is_page_field_run(run) -> bool:
    if run.find(W + "fldChar") is not None:
        return True
    instr = run.find(W + "instrText")
    return instr is not None and (instr.text or "").strip() == "PAGE"


def _ensure_page_field_in_part(parts, part_name: str, alignment, font, size_pt) -> None:
    raw = parts.get(part_name)
    if raw is None:
        return
    root = etree.fromstring(raw)
    page_paragraphs = [p for p in root.findall(W + "p") if _paragraph_has_page_field(p)]
    if page_paragraphs:
        for p in page_paragraphs:
            if alignment in ALIGN_MAP:
                p_pr = p.find(W + "pPr")
                if p_pr is None:
                    p_pr = etree.Element(W + "pPr")
                    p.insert(0, p_pr)
                _set_pPr_child(p_pr, W + "jc", {W + "val": ALIGN_MAP[alignment]})
            for run in p.findall(W + "r"):
                if _is_page_field_run(run):
                    _apply_rpr(run, font, size_pt)
    else:
        p_element = etree.fromstring(_page_field_paragraph_xml(alignment, font, size_pt).encode("utf-8"))
        root.append(p_element)
    parts[part_name] = _serialize_xml(root)


def _remove_page_field_from_part(parts, part_name: str) -> None:
    raw = parts.get(part_name)
    if raw is None:
        return
    root = etree.fromstring(raw)
    for p in list(root.findall(W + "p")):
        if not _paragraph_has_page_field(p):
            continue
        for run in list(p.findall(W + "r")):
            if _is_page_field_run(run):
                p.remove(run)
    parts[part_name] = _serialize_xml(root)


def _insert_sect_pr_child(sect_pr, child) -> None:
    """Insert a sectPr child respecting the common CT_SectPr sequence."""
    header_footer_tags = (W + "headerReference", W + "footerReference")
    # header/footer references belong at the front of the reference group
    if child.tag in header_footer_tags:
        insert_after = None
        for existing in sect_pr:
            if existing.tag in header_footer_tags:
                insert_after = existing
        if insert_after is not None:
            insert_after.addnext(child)
        else:
            sect_pr.insert(0, child)
        return
    if child.tag == W + "pgNumType":
        cols = sect_pr.find(W + "cols")
        if cols is not None:
            cols.addprevious(child)
            return
        pg_mar = sect_pr.find(W + "pgMar")
        if pg_mar is not None:
            pg_mar.addnext(child)
            return
        pg_sz = sect_pr.find(W + "pgSz")
        if pg_sz is not None:
            pg_sz.addnext(child)
            return
        sect_pr.append(child)
        return
    if child.tag == W + "titlePg":
        cols = sect_pr.find(W + "cols")
        if cols is not None:
            cols.addnext(child)
            return
        pg_num_type = sect_pr.find(W + "pgNumType")
        if pg_num_type is not None:
            pg_num_type.addnext(child)
            return
        pg_mar = sect_pr.find(W + "pgMar")
        if pg_mar is not None:
            pg_mar.addnext(child)
            return
        sect_pr.append(child)
        return
    sect_pr.append(child)


def _ensure_chrome_reference(
    parts,
    sect_pr,
    position: str,
    ref_type: str,
    with_page_field: bool,
    alignment,
    font,
    size_pt,
) -> str:
    """Ensure a header/footer reference exists; return the part name."""
    ref_tag = W + "footerReference" if position == "footer" else W + "headerReference"
    rel_type = FTR_REL_TYPE if position == "footer" else HDR_REL_TYPE
    content_type = FTR_CT if position == "footer" else HDR_CT
    prefix = "footer" if position == "footer" else "header"

    existing_ref = None
    for candidate in sect_pr.findall(ref_tag):
        candidate_type = candidate.get(W + "type", "default")
        if candidate_type == ref_type:
            existing_ref = candidate
            break

    rels_name, rels_root = _parse_relationships(parts)

    if existing_ref is not None:
        rid = existing_ref.get(R + "id")
        target = None
        if rid:
            for rel in rels_root.findall(PR + "Relationship"):
                if rel.get("Id") == rid:
                    target = rel.get("Target")
                    break
        if target:
            full_target = target if target.startswith("word/") else "word/" + target
            if full_target in parts:
                if with_page_field:
                    _ensure_page_field_in_part(parts, full_target, alignment, font, size_pt)
                else:
                    _remove_page_field_from_part(parts, full_target)
                return full_target
        # Existing reference but missing part: create a fresh part and repoint.
        part_name = f"word/{prefix}{_next_part_index(parts, prefix)}.xml"
        parts[part_name] = _make_chrome_part_xml(position, alignment, font, size_pt, with_page_field)
        _ensure_content_type(parts, part_name, content_type)
        if rid is not None:
            for rel in rels_root.findall(PR + "Relationship"):
                if rel.get("Id") == rid:
                    rel.set("Target", part_name.split("/", 1)[1])
                    parts[rels_name] = _serialize_xml(rels_root)
                    return part_name
        new_rid = _add_relationship(parts, rels_name, rels_root, rel_type, part_name.split("/", 1)[1])
        existing_ref.set(R + "id", new_rid)
        return part_name

    # No reference yet: create a part, relationship and reference element.
    part_name = f"word/{prefix}{_next_part_index(parts, prefix)}.xml"
    parts[part_name] = _make_chrome_part_xml(position, alignment, font, size_pt, with_page_field)
    _ensure_content_type(parts, part_name, content_type)
    new_rid = _add_relationship(parts, rels_name, rels_root, rel_type, part_name.split("/", 1)[1])
    ref = etree.Element(ref_tag)
    ref.set(W + "type", ref_type)
    ref.set(R + "id", new_rid)
    _insert_sect_pr_child(sect_pr, ref)
    return part_name


def _apply_page_number(body_element, page_number: dict[str, Any], parts) -> list[str]:
    warnings: list[str] = []
    if not page_number:
        return warnings
    if page_number.get("enabled") is False:
        return warnings
    position = page_number.get("position", "footer")
    if position not in ("footer", "header"):
        warnings.append("PAGE_NUMBER_POSITION_UNSUPPORTED")
        return warnings
    alignment = page_number.get("alignment", "center")
    if alignment not in ALIGN_MAP:
        warnings.append("PAGE_NUMBER_ALIGNMENT_UNSUPPORTED")
        alignment = "center"
    start_at = page_number.get("start_at")
    show_on_first = bool(page_number.get("show_on_first_page", True))
    font = page_number.get("font")
    size_pt = page_number.get("size_pt")

    sect_prs = _find_sect_prs(body_element)
    if not sect_prs:
        warnings.append("PAGE_NUMBER_NO_SECTPR")
        return warnings

    if start_at is not None:
        pg_num_type = sect_prs[0].find(W + "pgNumType")
        if pg_num_type is None:
            pg_num_type = etree.Element(W + "pgNumType")
            _insert_sect_pr_child(sect_prs[0], pg_num_type)
        pg_num_type.set(W + "start", str(start_at))

    for sect_pr in sect_prs:
        _ensure_chrome_reference(parts, sect_pr, position, "default", True, alignment, font, size_pt)
        if not show_on_first:
            if sect_pr.find(W + "titlePg") is None:
                title_pg = etree.Element(W + "titlePg")
                _insert_sect_pr_child(sect_pr, title_pg)
            _ensure_chrome_reference(parts, sect_pr, position, "first", False, alignment, font, size_pt)
        else:
            if sect_pr.find(W + "titlePg") is not None:
                _ensure_chrome_reference(parts, sect_pr, position, "first", True, alignment, font, size_pt)
    return warnings


def _set_tblPr_child(tbl_pr, tag, attrs):
    element = tbl_pr.find(tag)
    if element is None:
        element = etree.SubElement(tbl_pr, tag)
    for key, value in attrs.items():
        element.set(key, value)
    return element


def _apply_table_format(tbl_element, slot: dict[str, Any]) -> None:
    if not slot:
        return
    tbl_pr = tbl_element.find(W + "tblPr")
    if tbl_pr is None:
        tbl_pr = etree.Element(W + "tblPr")
        tbl_element.insert(0, tbl_pr)
    if "alignment" in slot and slot["alignment"] in ("left", "center", "right"):
        _set_tblPr_child(tbl_pr, W + "jc", {W + "val": slot["alignment"]})
    if "preferred_width_cm" in slot:
        _set_tblPr_child(
            tbl_pr,
            W + "tblW",
            {W + "w": str(int(round(float(slot["preferred_width_cm"]) * 567))), W + "type": "dxa"},
        )
    if "autofit" in slot:
        _set_tblPr_child(
            tbl_pr,
            W + "tblLayout",
            {W + "type": "autofit" if slot["autofit"] else "fixed"},
        )
    borders = slot.get("borders")
    if isinstance(borders, dict) and borders:
        tbl_borders = tbl_pr.find(W + "tblBorders")
        if tbl_borders is None:
            tbl_borders = etree.SubElement(tbl_pr, W + "tblBorders")
        edge_map = {
            "top": W + "top",
            "bottom": W + "bottom",
            "left": W + "left",
            "right": W + "right",
            "inside_horizontal": W + "insideH",
            "inside_vertical": W + "insideV",
        }
        for key, tag in edge_map.items():
            if key in borders:
                spec = borders[key]
                if isinstance(spec, dict):
                    edge = tbl_borders.find(tag)
                    if edge is None:
                        edge = etree.SubElement(tbl_borders, tag)
                    edge.set(W + "val", str(spec.get("val", "single")))
                    if "sz" in spec:
                        edge.set(W + "sz", str(spec["sz"]))
                    if "color" in spec:
                        edge.set(W + "color", str(spec["color"]))
    text_slot = slot.get("text")
    cell_slot = slot.get("cell")
    if isinstance(text_slot, dict) or isinstance(cell_slot, dict):
        for tr in tbl_element.findall(W + "tr"):
            for tc in tr.findall(W + "tc"):
                if isinstance(cell_slot, dict):
                    tc_pr = tc.find(W + "tcPr")
                    if tc_pr is None:
                        tc_pr = etree.Element(W + "tcPr")
                        tc.insert(0, tc_pr)
                    if "vertical_alignment" in cell_slot:
                        _set_tblPr_child(tc_pr, W + "vAlign", {W + "val": str(cell_slot["vertical_alignment"])})
                    if any(k in cell_slot for k in ("margin_top_cm", "margin_bottom_cm", "margin_left_cm", "margin_right_cm")):
                        mar = tc_pr.find(W + "tcMar")
                        if mar is None:
                            mar = etree.SubElement(tc_pr, W + "tcMar")
                        mar_map = {
                            "margin_top_cm": W + "top",
                            "margin_bottom_cm": W + "bottom",
                            "margin_left_cm": W + "left",
                            "margin_right_cm": W + "right",
                        }
                        for key, tag in mar_map.items():
                            if key in cell_slot:
                                m = mar.find(tag)
                                if m is None:
                                    m = etree.SubElement(mar, tag)
                                m.set(W + "w", str(int(round(float(cell_slot[key]) * 567))))
                                m.set(W + "type", "dxa")
                    if "shading" in cell_slot:
                        shd = tc_pr.find(W + "shd")
                        if shd is None:
                            shd = etree.SubElement(tc_pr, W + "shd")
                        shd.set(W + "fill", str(cell_slot["shading"]))
                if isinstance(text_slot, dict):
                    for p in tc.findall(W + "p"):
                        _apply_paragraph_format(p, text_slot)


def _media_only_paragraph(p_element) -> bool:
    return p_element.find(f".//{W}t") is None


def _apply_inline_image_format(p_element, slot: dict[str, Any]) -> list[str]:
    warnings = []
    if not slot:
        return warnings
    max_width_cm = slot.get("max_width_cm")
    max_height_cm = slot.get("max_height_cm")
    preserve_aspect_ratio = slot.get("preserve_aspect_ratio", True)
    allow_upscale = slot.get("allow_upscale", False)

    for inline in p_element.iter(WP + "inline"):
        extents = inline.findall(WP + "extent")
        if not extents:
            continue
        extent = extents[0]
        cx = int(extent.get("cx", "0"))
        cy = int(extent.get("cy", "0"))
        if cx <= 0 or cy <= 0:
            continue
        max_cx = int(round(float(max_width_cm) * EMU_PER_CM)) if max_width_cm else None
        max_cy = int(round(float(max_height_cm) * EMU_PER_CM)) if max_height_cm else None
        if preserve_aspect_ratio:
            scale_x = max_cx / cx if max_cx else None
            scale_y = max_cy / cy if max_cy else None
            scale = 1.0
            if scale_x is not None:
                scale = min(scale, scale_x)
            if scale_y is not None:
                scale = min(scale, scale_y)
            if not allow_upscale:
                scale = min(1.0, scale)
            new_cx = int(round(cx * scale))
            new_cy = int(round(cy * scale))
        else:
            new_cx = min(cx, max_cx) if max_cx else cx
            new_cy = min(cy, max_cy) if max_cy else cy
            if not allow_upscale:
                new_cx = min(cx, new_cx)
                new_cy = min(cy, new_cy)
        extent.set("cx", str(new_cx))
        extent.set("cy", str(new_cy))
        for xfrm_ext in inline.iter(A + "ext"):
            xfrm_ext.set("cx", str(new_cx))
            xfrm_ext.set("cy", str(new_cy))

    if "alignment" in slot:
        if _media_only_paragraph(p_element):
            p_pr = p_element.find(W + "pPr")
            if p_pr is None:
                p_pr = etree.Element(W + "pPr")
                p_element.insert(0, p_pr)
            align = slot["alignment"]
            if align in ALIGN_MAP:
                _set_pPr_child(p_pr, W + "jc", {W + "val": ALIGN_MAP[align]})
        else:
            warnings.append("IMAGE_ALIGNMENT_DEFERRED_FOR_MIXED_PARAGRAPH")

    for _ in p_element.iter(WP + "anchor"):
        warnings.append("ANCHOR_IMAGE_FORMAT_DEFERRED")
    if p_element.findall(f".//{V}imagedata"):
        warnings.append("VML_IMAGE_FORMAT_DEFERRED")
    return warnings


def _table_structure_signature(ir: DocumentIR):
    parts = []
    for block in ir.blocks:
        if block.__class__.__name__ == "TableBlock":
            rows = []
            for row in block.rows:
                cells = []
                for cell in row:
                    cells.append(
                        str(cell.metadata.get("grid_span", ""))
                        + ":"
                        + str(cell.metadata.get("v_merge", ""))
                        + ":"
                        + ",".join(p.text for p in cell.blocks)
                    )
                rows.append("|".join(cells))
            parts.append("T(" + ";".join(rows) + ")")
    return CH.join(parts)


def _media_relationship_signature(ir: DocumentIR):
    lines = []
    for item in ir.media:
        lines.append(str(item.relationship_id or "") + ":" + item.part_name + ":" + item.sha256)
    return CH.join(lines)


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def render_reformat(source_path, plan: ReformatPlan, output_path):
    source = Path(source_path)
    output = Path(output_path)

    if output.resolve() == source.resolve():
        return {
            "status": "error",
            "source_path": str(source),
            "output_path": str(output),
            "target_profile_id": plan.target_profile_id,
            "operations": {"planned": len(plan.operations), "applied": 0, "preserved": 0, "deferred": 0},
            "warnings": [],
            "errors": [SOURCE_OUTPUT_PATH_CONFLICT],
            "content_preservation": {"passed": False},
        }

    actual_sha = _sha256_file(source)
    if actual_sha != plan.source_file_sha256:
        return {
            "status": "error",
            "source_path": str(source),
            "output_path": str(output),
            "target_profile_id": plan.target_profile_id,
            "operations": {"planned": len(plan.operations), "applied": 0, "preserved": 0, "deferred": 0},
            "warnings": [],
            "errors": [SOURCE_CHANGED],
            "content_preservation": {"passed": False},
        }

    if not plan.ready or any("OpaqueBlock" in b or "unsupported" in b for b in plan.blockers):
        return {
            "status": "error",
            "source_path": str(source),
            "output_path": str(output),
            "target_profile_id": plan.target_profile_id,
            "operations": {"planned": len(plan.operations), "applied": 0, "preserved": 0, "deferred": 0},
            "warnings": [],
            "errors": [PLAN_NOT_READY],
            "content_preservation": {"passed": False},
        }

    try:
        profile = profile_registry.resolve_profile(plan.target_profile_id)
    except KeyError:
        return {
            "status": "error",
            "source_path": str(source),
            "output_path": str(output),
            "target_profile_id": plan.target_profile_id,
            "operations": {"planned": len(plan.operations), "applied": 0, "preserved": 0, "deferred": 0},
            "warnings": [],
            "errors": [PROFILE_NOT_FOUND],
            "content_preservation": {"passed": False},
        }

    source_ir = read_docx(source)
    tmp = Path(tempfile.mktemp(suffix=".docx", prefix="forge_render_"))
    shutil.copy2(source, tmp)

    applied = preserved = deferred = 0
    warnings: list[str] = []

    try:
        with ZipFile(tmp) as zin:
            parts = {info.filename: zin.read(info.filename) for info in zin.infolist()}
        document_xml = parts["word/document.xml"]
        root = etree.fromstring(document_xml)
        body = root.find(W + "body")

        _apply_page_format(body, profile.page)
        warnings.extend(_apply_page_number(body, profile.page_number or {}, parts))

        for op in plan.operations:
            if op.action == "apply_profile_style":
                locator = (op.metadata or {}).get("source_locator")
                element = _locate(body, locator) if locator else None
                if element is None:
                    warnings.append("定位失败，保留源格式: " + op.block_id)
                    preserved += 1
                    continue
                if op.block_type == "table":
                    _apply_table_format(element, profile.table)
                    applied += 1
                    continue
                slot = getattr(profile, op.style_slot, None) if op.style_slot else None
                if isinstance(slot, dict):
                    _apply_paragraph_format(element, slot)
                    applied += 1
                    image_warnings = _apply_inline_image_format(element, profile.image or {})
                    warnings.extend(image_warnings)
                    continue
                warnings.append("profile 缺少 slot: " + str(op.style_slot))
                preserved += 1
            elif op.action == "preserve_structure":
                preserved += 1
            elif op.action == "review_required":
                preserved += 1
                warnings.append("SOURCE_FORMAT_PRESERVED_FOR_UNRESOLVED_BLOCK")
            elif op.action == "unsupported":
                deferred += 1

        parts["word/document.xml"] = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        with ZipFile(tmp, "w", ZIP_DEFLATED) as zout:
            for name, data in parts.items():
                zout.writestr(name, data)

        re_ir = read_docx(tmp)
        re_fp = re_ir.content_fingerprint
        expected = plan.source_fingerprint or {}
        preservation = {
            "passed": False,
            "text": re_fp.text_sha256 == expected.get("text_sha256"),
            "structure": re_fp.structure_sha256 == expected.get("structure_sha256"),
            "media": re_fp.media_sha256 == expected.get("media_sha256"),
            "sequence": re_fp.content_sequence_sha256 == expected.get("content_sequence_sha256"),
            "table_structure": _table_structure_signature(source_ir) == _table_structure_signature(re_ir),
            "media_relationships": _media_relationship_signature(source_ir) == _media_relationship_signature(re_ir),
        }
        preservation["passed"] = all(
            preservation[k] for k in ("text", "structure", "media", "sequence", "table_structure", "media_relationships")
        )

        if not preservation["passed"]:
            return {
                "status": "error",
                "source_path": str(source),
                "output_path": str(output),
                "target_profile_id": plan.target_profile_id,
                "operations": {"planned": len(plan.operations), "applied": applied, "preserved": preserved, "deferred": deferred},
                "warnings": warnings,
                "errors": [CONTENT_PRESERVATION_FAILED],
                "content_preservation": preservation,
            }

        output.parent.mkdir(parents=True, exist_ok=True)
        os.replace(tmp, output)
        return {
            "status": "ok",
            "source_path": str(source),
            "output_path": str(output),
            "target_profile_id": plan.target_profile_id,
            "operations": {"planned": len(plan.operations), "applied": applied, "preserved": preserved, "deferred": deferred},
            "warnings": warnings,
            "errors": [],
            "content_preservation": preservation,
        }
    finally:
        if tmp.exists():
            try:
                tmp.unlink()
            except OSError:
                pass
