"""FORGE Mission 03-B — Document IR + Faithful Ordered DOCX Reader.

核心原则：Preserve the content. Reforge the format.

只读、顺序保真。认识的内容完整读取；暂时不认识的内容不得静默丢失。
"""

from __future__ import annotations

import hashlib
import mimetypes
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional
from zipfile import ZipFile

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
V_NS = "urn:schemas-microsoft-com:vml"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

W = "{%s}" % W_NS
A = "{%s}" % A_NS
R = "{%s}" % R_NS
V = "{%s}" % V_NS
WP = "{%s}" % WP_NS


@dataclass
class Inline:
    """段落内联内容，保持段落内部真实顺序。"""
    type: str
    text: Optional[str] = None
    media_id: Optional[str] = None
    relationship_id: Optional[str] = None
    break_type: Optional[str] = None
    target: Optional[str] = None
    children: list["Inline"] = field(default_factory=list)
    metadata: dict[str, Any] = field(default_factory=dict)


@dataclass
class ParagraphStyle:
    style_name: Optional[str] = None
    alignment: Optional[str] = None
    font_name: Optional[str] = None
    font_size_pt: Optional[float] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None


@dataclass
class ParagraphBlock:
    id: str
    type: str = "paragraph"
    source_index: int = 0
    text: str = ""
    inline: list[Inline] = field(default_factory=list)
    semantic_role: Optional[str] = None
    role_confidence: Optional[float] = None
    role_evidence: list[str] = field(default_factory=list)
    style: ParagraphStyle = field(default_factory=ParagraphStyle)
    metadata: dict[str, Any] = field(default_factory=dict)


@dataclass
class CellBlock:
    blocks: list[ParagraphBlock] = field(default_factory=list)
    metadata: dict[str, Any] = field(default_factory=dict)


@dataclass
class TableBlock:
    id: str
    type: str = "table"
    source_index: int = 0
    rows: list[list[CellBlock]] = field(default_factory=list)
    metadata: dict[str, Any] = field(default_factory=dict)


@dataclass
class OpaqueBlock:
    """未知/可能含内容的 body 元素：不静默丢失。"""
    id: str
    type: str = "opaque"
    source_index: int = 0
    xml_tag: str = ""
    extracted_text: str = ""
    raw_xml_sha256: str = ""
    metadata: dict[str, Any] = field(default_factory=dict)


@dataclass
class MediaItem:
    media_id: str
    relationship_id: Optional[str]
    part_name: str
    content_type: Optional[str]
    sha256: str
    size_bytes: int


@dataclass
class SectionInfo:
    index: int
    page_width: Optional[float] = None
    page_height: Optional[float] = None
    margin_top: Optional[float] = None
    margin_bottom: Optional[float] = None
    margin_left: Optional[float] = None
    margin_right: Optional[float] = None
    header_distance: Optional[float] = None
    footer_distance: Optional[float] = None


@dataclass
class Statistics:
    block_count: int = 0
    paragraph_count: int = 0
    table_count: int = 0
    image_count: int = 0
    section_count: int = 0


@dataclass
class ContentFingerprint:
    text_sha256: str = ""
    structure_sha256: str = ""
    media_sha256: str = ""
    content_sequence_sha256: str = ""


@dataclass
class DocumentIR:
    source: str
    source_file_sha256: str = ""
    blocks: list[Any] = field(default_factory=list)
    media: list[MediaItem] = field(default_factory=list)
    sections: list[SectionInfo] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    statistics: Statistics = field(default_factory=Statistics)
    content_fingerprint: ContentFingerprint = field(default_factory=ContentFingerprint)


def _element_text(element) -> str:
    return "".join(t.text or "" for t in element.iter(W + "t"))


def _parse_style(p_element) -> ParagraphStyle:
    style = ParagraphStyle()
    ppr = p_element.find(W + "pPr")
    if ppr is not None:
        pstyle = ppr.find(W + "pStyle")
        if pstyle is not None:
            style.style_name = pstyle.get(W + "val")
        jc = ppr.find(W + "jc")
        if jc is not None:
            style.alignment = jc.get(W + "val")
        rpr = ppr.find(W + "rPr")
        if rpr is not None:
            _apply_rpr(style, rpr)
    first_run = p_element.find(W + "r")
    if first_run is not None:
        rpr = first_run.find(W + "rPr")
        if rpr is not None:
            _apply_rpr(style, rpr)
    return style


def _apply_rpr(style: ParagraphStyle, rpr) -> None:
    rfonts = rpr.find(W + "rFonts")
    if rfonts is not None:
        style.font_name = rfonts.get(W + "eastAsia") or rfonts.get(W + "ascii")
    sz = rpr.find(W + "sz")
    if sz is not None:
        try:
            style.font_size_pt = int(sz.get(W + "val")) / 2.0
        except (TypeError, ValueError):
            pass
    bold = rpr.find(W + "b")
    if bold is not None:
        val = bold.get(W + "val")
        style.bold = val not in ("0", "false")
    italic = rpr.find(W + "i")
    if italic is not None:
        val = italic.get(W + "val")
        style.italic = val not in ("0", "false")


def _find_embeds(element) -> list[tuple[str, str, str]]:
    embeds = []
    for blip in element.iter(A + "blip"):
        rid = blip.get(R + "embed")
        if rid:
            layout = "inline"
            ancestor = blip
            while ancestor is not None:
                if ancestor.tag == WP + "anchor":
                    layout = "anchor"
                    break
                if ancestor.tag == WP + "inline":
                    layout = "inline"
                    break
                ancestor = ancestor.getparent()
            embeds.append((rid, layout, "blip"))
    for imagedata in element.iter(V + "imagedata"):
        rid = imagedata.get(R + "id")
        if rid:
            embeds.append((rid, "vml", "vml"))
    return embeds


def _parse_inline(container, inline_list, media_by_rid, rid_to_target) -> None:
    for child in container:
        tag = child.tag
        if tag == W + "r":
            for run_child in child:
                run_tag = run_child.tag
                if run_tag == W + "t":
                    text = run_child.text or ""
                    if text:
                        inline_list.append(Inline(type="text", text=text))
                elif run_tag == W + "tab":
                    inline_list.append(Inline(type="tab"))
                elif run_tag == W + "br":
                    break_type = run_child.get(W + "type") or "line"
                    inline_list.append(
                        Inline(
                            type="page_break" if break_type == "page" else "line_break",
                            break_type=break_type,
                        )
                    )
                elif run_tag == W + "drawing":
                    for rid, layout, _kind in _find_embeds(run_child):
                        inline_list.append(
                            Inline(
                                type="image",
                                relationship_id=rid,
                                media_id=media_by_rid.get(rid),
                                metadata={"layout": layout},
                            )
                        )
                elif run_tag == W + "pict":
                    for rid, layout, _kind in _find_embeds(run_child):
                        inline_list.append(
                            Inline(
                                type="image",
                                relationship_id=rid,
                                media_id=media_by_rid.get(rid),
                                metadata={"layout": layout},
                            )
                        )
        elif tag == W + "hyperlink":
            rid = child.get(R + "id")
            hyperlink = Inline(
                type="hyperlink",
                relationship_id=rid,
                target=rid_to_target.get(rid) if rid else None,
            )
            _parse_inline(child, hyperlink.children, media_by_rid, rid_to_target)
            hyperlink.text = "".join(c.text or "" for c in hyperlink.children if c.type == "text")
            inline_list.append(hyperlink)
        elif tag == W + "sdt":
            sdt_content = child.find(W + "sdtContent")
            if sdt_content is not None:
                _parse_inline(sdt_content, inline_list, media_by_rid, rid_to_target)


def _build_paragraph(p_element, block_id, source_index, media_by_rid, rid_to_target, metadata=None) -> ParagraphBlock:
    inline: list[Inline] = []
    _parse_inline(p_element, inline, media_by_rid, rid_to_target)
    block_metadata = dict(metadata or {})
    ppr = p_element.find(W + "pPr")
    if ppr is not None:
        numpr = ppr.find(W + "numPr")
        if numpr is not None:
            num_id_el = numpr.find(W + "numId")
            ilvl_el = numpr.find(W + "ilvl")
            block_metadata["numbering"] = {
                "num_id": num_id_el.get(W + "val") if num_id_el is not None else None,
                "ilvl": ilvl_el.get(W + "val") if ilvl_el is not None else None,
            }
    return ParagraphBlock(
        id=block_id,
        type="paragraph",
        source_index=source_index,
        text=_element_text(p_element),
        inline=inline,
        semantic_role=None,
        role_confidence=None,
        role_evidence=[],
        style=_parse_style(p_element),
        metadata=block_metadata,
    )


def _build_table(tbl_element, block_id, source_index, media_by_rid, rid_to_target, metadata=None, table_locator=None) -> TableBlock:
    rows: list[list[CellBlock]] = []
    for ri, tr in enumerate(tbl_element.findall(W + "tr")):
        row_cells: list[CellBlock] = []
        for ci, tc in enumerate(tr.findall(W + "tc")):
            tc_pr = tc.find(W + "tcPr")
            cell_metadata: dict[str, Any] = {}
            if tc_pr is not None:
                grid_span = tc_pr.find(W + "gridSpan")
                if grid_span is not None:
                    cell_metadata["grid_span"] = grid_span.get(W + "val")
                v_merge = tc_pr.find(W + "vMerge")
                if v_merge is not None:
                    cell_metadata["v_merge"] = v_merge.get(W + "val") or "continue"
            cell_blocks: list[ParagraphBlock] = []
            for pi, p in enumerate(tc.findall(W + "p")):
                paragraph_metadata = None
                if table_locator:
                    paragraph_metadata = {
                        "source_locator": f"{table_locator}/table/row/{ri}/cell/{ci}/p/{pi}"
                    }
                cell_blocks.append(
                    _build_paragraph(
                        p,
                        f"{block_id}_r{ri}_c{ci}_p{pi}",
                        source_index,
                        media_by_rid,
                        rid_to_target,
                        paragraph_metadata,
                    )
                )
            row_cells.append(CellBlock(blocks=cell_blocks, metadata=cell_metadata))
        rows.append(row_cells)
    return TableBlock(id=block_id, type="table", source_index=source_index, rows=rows, metadata=dict(metadata or {}))


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _read_sections(docx_path: Path) -> list[SectionInfo]:
    from docx import Document

    document = Document(docx_path)
    sections: list[SectionInfo] = []
    for index, section in enumerate(document.sections):
        sections.append(
            SectionInfo(
                index=index,
                page_width=round(section.page_width.cm, 4) if section.page_width else None,
                page_height=round(section.page_height.cm, 4) if section.page_height else None,
                margin_top=round(section.top_margin.cm, 4) if section.top_margin else None,
                margin_bottom=round(section.bottom_margin.cm, 4) if section.bottom_margin else None,
                margin_left=round(section.left_margin.cm, 4) if section.left_margin else None,
                margin_right=round(section.right_margin.cm, 4) if section.right_margin else None,
                header_distance=round(section.header_distance.cm, 4) if section.header_distance else None,
                footer_distance=round(section.footer_distance.cm, 4) if section.footer_distance else None,
            )
        )
    return sections


def _may_contain_content(element) -> bool:
    return (
        element.find(f".//{W}t") is not None
        or element.find(f".//{W}p") is not None
        or element.find(f".//{W}tbl") is not None
        or element.find(f".//{A}blip") is not None
        or element.find(f".//{V}imagedata") is not None
        or element.find(f".//{W}hyperlink") is not None
        or element.find(f".//{W}sdt") is not None
    )


def read_docx(path: str | Path) -> DocumentIR:
    docx_path = Path(path)
    source_bytes = docx_path.read_bytes()
    ir = DocumentIR(
        source=str(docx_path.resolve()),
        source_file_sha256=hashlib.sha256(source_bytes).hexdigest(),
    )

    with ZipFile(docx_path) as archive:
        names = archive.namelist()
        document_xml = archive.read("word/document.xml")
        try:
            rels_xml = archive.read("word/_rels/document.xml.rels").decode("utf-8")
        except KeyError:
            rels_xml = ""
            ir.warnings.append("缺少 word/_rels/document.xml.rels")

        import re as _re

        rid_to_target: dict[str, str] = {}
        for rid, target in _re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels_xml):
            rid_to_target[rid] = target
        rid_to_media: dict[str, str] = {
            rid: target for rid, target in rid_to_target.items() if target.startswith("media/")
        }

        media_by_rid: dict[str, str] = {}
        media_parts = sorted(n for n in names if n.startswith("word/media/"))
        for media_index, part_name in enumerate(media_parts):
            data = archive.read(part_name)
            media_id = f"media_{media_index}"
            ir.media.append(
                MediaItem(
                    media_id=media_id,
                    relationship_id=None,
                    part_name=part_name,
                    content_type=mimetypes.guess_type(part_name)[0],
                    sha256=_sha256_bytes(data),
                    size_bytes=len(data),
                )
            )
        for rid, target in rid_to_media.items():
            full_part = target if target.startswith("word/") else "word/" + target
            for media_item in ir.media:
                if media_item.part_name == full_part:
                    media_item.relationship_id = rid
                    media_by_rid[rid] = media_item.media_id
                    break

        advanced_markers = {
            f"{W}footnoteReference": "脚注",
            f"{W}endnoteReference": "尾注",
            f"{W}commentReference": "批注",
            f"{W}object": "OLE 对象",
            "{urn:schemas-microsoft-com:office:office}OLEObject": "OLE 对象",
            f"{W}fldChar": "复杂域",
        }
        root = etree.fromstring(document_xml)
        for tag, label in advanced_markers.items():
            if root.find(f".//{tag}") is not None:
                ir.warnings.append(f"检测到暂不完整支持的 {label}")

        body = root.find(W + "body")
        if body is None:
            ir.warnings.append("缺少 w:body")
            body = root

        block_index = 0

        def walk_container(
            container,
            container_metadata: Optional[dict[str, Any]] = None,
            locator_prefix: str = "body",
        ):
            nonlocal block_index
            for child_index, child in enumerate(container):
                child_locator = f"{locator_prefix}/{child_index}"
                if child.tag == W + "p":
                    block_id = f"b{block_index}"
                    paragraph = _build_paragraph(
                        child, block_id, block_index, media_by_rid, rid_to_target, container_metadata
                    )
                    paragraph.metadata["source_locator"] = child_locator
                    ir.blocks.append(paragraph)
                    ir.statistics.paragraph_count += 1
                    block_index += 1
                elif child.tag == W + "tbl":
                    block_id = f"b{block_index}"
                    table = _build_table(
                        child, block_id, block_index, media_by_rid, rid_to_target, container_metadata,
                        table_locator=child_locator,
                    )
                    table.metadata["source_locator"] = child_locator
                    ir.blocks.append(table)
                    ir.statistics.table_count += 1
                    block_index += 1
                elif child.tag == W + "sdt":
                    sdt_content = child.find(W + "sdtContent")
                    if sdt_content is not None:
                        walk_container(
                            sdt_content,
                            {"container": "sdt"},
                            f"{child_locator}/sdtContent",
                        )
                    else:
                        ir.warnings.append("w:sdt 缺少 w:sdtContent")
                elif child.tag == W + "sectPr":
                    continue
                else:
                    if _may_contain_content(child):
                        opaque = OpaqueBlock(
                            id=f"b{block_index}",
                            type="opaque",
                            source_index=block_index,
                            xml_tag=child.tag.replace(W, "w:"),
                            extracted_text=_element_text(child),
                            raw_xml_sha256=_sha256_bytes(etree.tostring(child)),
                            metadata={"container": container_metadata.get("container")} if container_metadata else {},
                        )
                        opaque.metadata["source_locator"] = child_locator
                        ir.blocks.append(opaque)
                        block_index += 1
                    else:
                        ir.warnings.append(f"忽略无内容 body 元素: {child.tag.replace(W, 'w:')}")

        walk_container(body)

    ir.statistics.block_count = len(ir.blocks)
    ir.statistics.image_count = _count_images(ir.blocks)
    ir.sections = _read_sections(docx_path)
    ir.statistics.section_count = len(ir.sections)
    ir.content_fingerprint = _compute_fingerprint(ir)
    return ir


def _count_images(blocks) -> int:
    count = 0
    for block in blocks:
        if isinstance(block, ParagraphBlock):
            count += sum(1 for inline in block.inline if inline.type == "image")
        elif isinstance(block, TableBlock):
            for row in block.rows:
                for cell in row:
                    count += _count_images(cell.blocks)
    return count


def _paragraph_sequence_events(paragraph: ParagraphBlock) -> list[str]:
    events = ["p"]
    for inline in paragraph.inline:
        if inline.type == "text":
            events.append(f"t:{inline.text}")
        elif inline.type == "tab":
            events.append("tab")
        elif inline.type == "line_break":
            events.append("lb")
        elif inline.type == "page_break":
            events.append("pb")
        elif inline.type == "image":
            media_sha = inline.metadata.get("media_sha", "") if inline.metadata else ""
            events.append(f"img:{media_sha or inline.media_id}")
        elif inline.type == "hyperlink":
            events.append(f"hl:{inline.text}:{inline.target or ''}")
    events.append("/p")
    return events


def _block_sequence_events(block) -> list[str]:
    if isinstance(block, ParagraphBlock):
        return _paragraph_sequence_events(block)
    if isinstance(block, TableBlock):
        events = ["table"]
        for row in block.rows:
            events.append("row")
            for cell in row:
                events.append("cell")
                for paragraph in cell.blocks:
                    events.extend(_paragraph_sequence_events(paragraph))
                events.append("/cell")
            events.append("/row")
        events.append("/table")
        return events
    if isinstance(block, OpaqueBlock):
        return [f"opaque:{block.xml_tag}:{block.raw_xml_sha256}"]
    return []


def _compute_fingerprint(ir: DocumentIR) -> ContentFingerprint:
    text_parts: list[str] = []
    structure_parts: list[str] = []
    sequence_events: list[str] = []

    for block in ir.blocks:
        if isinstance(block, ParagraphBlock):
            structure_parts.append("P")
            text_parts.append(block.text)
        elif isinstance(block, TableBlock):
            row_count = len(block.rows)
            col_count = len(block.rows[0]) if block.rows else 0
            structure_parts.append(f"T({row_count},{col_count})")
            for row in block.rows:
                for cell in row:
                    text_parts.append("\n".join(p.text for p in cell.blocks))
        elif isinstance(block, OpaqueBlock):
            structure_parts.append(f"O({block.xml_tag})")
            text_parts.append(block.extracted_text)
        sequence_events.extend(_block_sequence_events(block))

    media_lines = sorted(f"{item.part_name}:{item.sha256}" for item in ir.media)
    return ContentFingerprint(
        text_sha256=hashlib.sha256("\n".join(text_parts).encode("utf-8")).hexdigest(),
        structure_sha256=hashlib.sha256(";".join(structure_parts).encode("utf-8")).hexdigest(),
        media_sha256=hashlib.sha256("\n".join(media_lines).encode("utf-8")).hexdigest(),
        content_sequence_sha256=hashlib.sha256("\n".join(sequence_events).encode("utf-8")).hexdigest(),
    )
