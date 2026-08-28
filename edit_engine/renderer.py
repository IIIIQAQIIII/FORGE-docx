"""Mission 07 — Source-preserving transactional Edit Renderer."""

from __future__ import annotations

import shutil
import tempfile
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

from document_ir import W, W_NS
from edit_engine.models import EditPlan


def _serialize_xml(root) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _make_paragraph(text: str, ppr_template=None) -> object:
    p = etree.fromstring(f'<w:p xmlns:w="{W_NS}"><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'.encode("utf-8"))
    if ppr_template is not None:
        ppr = ppr_template.find(W + "pPr")
        if ppr is not None:
            copied = etree.fromstring(etree.tostring(ppr))
            # never copy numbering relationships: safe insert only copies formatting
            for num_pr in copied.findall(W + "numPr"):
                copied.remove(num_pr)
            existing = p.find(W + "pPr")
            if existing is not None:
                p.replace(existing, copied)
            else:
                p.insert(0, copied)
    return p


def apply_plan_to_document_xml(document_root, plan: EditPlan) -> None:
    for bound in plan.operations:
        if bound.op == "replace_text":
            _apply_replace(bound)
        elif bound.op == "insert_paragraph_before":
            p = _make_paragraph(bound.text or "", bound.element)
            bound.element.addprevious(p)
        elif bound.op == "insert_paragraph_after":
            p = _make_paragraph(bound.text or "", bound.element)
            bound.element.addnext(p)
        elif bound.op == "append_paragraph":
            body = bound.element
            p = _make_paragraph(bound.text or "")
            sect_pr = body.find(W + "sectPr")
            if sect_pr is not None:
                sect_pr.addprevious(p)
            else:
                body.append(p)
        elif bound.op == "delete_paragraph":
            bound.element.getparent().remove(bound.element)


def _apply_replace(bound) -> None:
    paragraph = bound.element
    # Rebuild the text-node map from the CURRENT paragraph (bound once, but
    # replacements on different paragraphs don't affect this one).
    from edit_engine.planner import _collect_text_nodes

    nodes = _collect_text_nodes(paragraph)
    full_text = "".join(n["text"] for n in nodes)
    for plan in bound.match_plan:
        first = nodes[plan["first_idx"]]
        last = nodes[plan["last_idx"]]
        first["element"].text = plan["prefix"] + (bound.new_text or "") + plan["suffix"]
        for idx in range(plan["first_idx"] + 1, plan["last_idx"] + 1):
            nodes[idx]["element"].text = ""
        # nodes before first keep their prefix; nodes after last keep their suffix


def render_edit(source_path, output_path, plan: EditPlan):
    source = Path(source_path)
    output = Path(output_path)
    if output.resolve() == source.resolve():
        return {"status": "error", "errors": ["SOURCE_OUTPUT_PATH_CONFLICT"]}
    tmp = Path(tempfile.mktemp(suffix=".docx", prefix="forge_edit_"))
    shutil.copy2(source, tmp)
    try:
        with ZipFile(tmp) as zin:
            parts = {info.filename: zin.read(info.filename) for info in zin.infolist()}
        document_root = etree.fromstring(parts["word/document.xml"])

        # Re-bind each operation to the temp document BEFORE any mutation.
        from edit_engine.planner import _locate_from_body

        body = document_root.find(W + "body")
        for bound in plan.operations:
            if bound.op == "append_paragraph":
                bound.element = body
            else:
                bound.element = _locate_from_body(body, bound.locator)
                if bound.element is None:
                    return {"status": "error", "errors": ["EDIT_TARGET_NOT_FOUND:" + str(bound.locator)]}

        apply_plan_to_document_xml(document_root, plan)
        parts["word/document.xml"] = _serialize_xml(document_root)
        with ZipFile(tmp, "w", ZIP_DEFLATED) as zout:
            for name, data in parts.items():
                zout.writestr(name, data)

        output.parent.mkdir(parents=True, exist_ok=True)
        Path(tmp).replace(output)
        return {"status": "ok", "output_path": str(output)}
    finally:
        if Path(tmp).exists():
            try:
                Path(tmp).unlink()
            except OSError:
                pass
