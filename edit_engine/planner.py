"""Mission 07 — Edit Planner.

Binds every operation to its target element BEFORE any mutation happens, so
insertions/deletions cannot shift later locator lookups. Also detects conflicts
and simulates the expected IR block list for Edit Contract validation.
"""

from __future__ import annotations

from copy import deepcopy
from typing import Any, Optional

from lxml import etree

from document_ir import W, DocumentIR, ParagraphBlock, TableBlock, read_docx
from edit_engine.models import BoundEditOperation, EditOperation, EditPlan

ERROR_TARGET_NOT_FOUND = "EDIT_TARGET_NOT_FOUND"
ERROR_OCCURRENCE_MISMATCH = "EDIT_OCCURRENCE_MISMATCH"
ERROR_COMPLEX_BOUNDARY = "EDIT_COMPLEX_BOUNDARY_UNSUPPORTED"
ERROR_CONFLICT = "EDIT_OPERATION_CONFLICT"
ERROR_UNSAFE_DELETE = "EDIT_UNSAFE_DELETE_TARGET"


def _locate_from_body(body, locator: str):
    parts = (locator or "").split("/")
    if not parts or parts[0] != "body":
        return None
    current = body
    idx = 1
    if idx < len(parts) and parts[idx].isdigit():
        children = [c for c in current]
        index = int(parts[idx])
        if index < 0 or index >= len(children):
            return None
        current = children[index]
        idx += 1
    while idx < len(parts):
        part = parts[idx]
        if part == "sdtContent":
            current = current.find(W + "sdtContent")
            idx += 1
        elif part == "table":
            idx += 1
        elif part == "row":
            ri = int(parts[idx + 1])
            rows = current.findall(W + "tr")
            if ri < 0 or ri >= len(rows):
                return None
            current = rows[ri]
            idx += 2
        elif part == "cell":
            ci = int(parts[idx + 1])
            cells = current.findall(W + "tc")
            if ci < 0 or ci >= len(cells):
                return None
            current = cells[ci]
            idx += 2
        elif part == "p":
            pi = int(parts[idx + 1])
            paragraphs = current.findall(W + "p")
            if pi < 0 or pi >= len(paragraphs):
                return None
            current = paragraphs[pi]
            idx += 2
        else:
            return None
    return current


def _iter_paragraph_blocks(ir: DocumentIR):
    for block in ir.blocks:
        if isinstance(block, ParagraphBlock):
            yield block
        elif isinstance(block, TableBlock):
            for row in block.rows:
                for cell in row:
                    for paragraph in cell.blocks:
                        yield paragraph


def _build_id_locator_map(ir: DocumentIR) -> dict:
    mapping = {}
    for block in _iter_paragraph_blocks(ir):
        mapping[block.id] = block.metadata.get("source_locator")
    return mapping


def _collect_text_nodes(paragraph_element):
    """Collect w:t text nodes in document order with barrier flags."""
    nodes = []
    last_container_key = None
    protected_since_last = False

    def walk(element, container_key):
        nonlocal last_container_key, protected_since_last
        for child in element:
            tag = child.tag
            if tag == W + "r":
                for run_child in child:
                    if run_child.tag == W + "t":
                        text = run_child.text or ""
                        key = container_key or "run"
                        barrier = False
                        if last_container_key is not None and key != last_container_key:
                            barrier = True
                        if protected_since_last:
                            barrier = True
                        nodes.append({"element": run_child, "text": text, "key": key, "barrier": barrier})
                        last_container_key = key
                        protected_since_last = False
                    elif run_child.tag in (W + "drawing", W + "fldChar", W + "object", W + "pict"):
                        protected_since_last = True
            elif tag == W + "hyperlink":
                walk(child, "hyperlink")
                protected_since_last = True
            elif tag == W + "sdt":
                walk(child, "sdt")
                protected_since_last = True
            elif tag in (W + "bookmarkStart", W + "bookmarkEnd"):
                pass

    walk(paragraph_element, "run")
    return nodes


def _find_text_occurrences(paragraph_element, old_text: str, expected_occurrences: int):
    nodes = _collect_text_nodes(paragraph_element)
    full_text = "".join(n["text"] for n in nodes)
    start = 0
    occurrences = []
    while True:
        pos = full_text.find(old_text, start)
        if pos < 0:
            break
        occurrences.append((pos, pos + len(old_text)))
        start = pos + max(1, len(old_text))
    return nodes, full_text, occurrences


def _match_plan_for_occurrence(nodes, full_text, occ):
    start, end = occ
    # find nodes overlapping [start, end)
    idx = 0
    pos = 0
    plan_nodes = []
    for i, node in enumerate(nodes):
        node_start = pos
        node_end = pos + len(node["text"])
        pos = node_end
        if node_end <= start:
            continue
        if node_start >= end:
            break
        if node_start < end and node_end > start:
            plan_nodes.append(i)
    first_idx = plan_nodes[0]
    last_idx = plan_nodes[-1]
    for i in plan_nodes:
        if i != first_idx and nodes[i]["barrier"]:
            return None
    prefix = nodes[first_idx]["text"][: start - _node_start(nodes, first_idx, full_text)]
    suffix = nodes[last_idx]["text"][end - _node_start(nodes, last_idx, full_text) :]
    return {"first_idx": first_idx, "last_idx": last_idx, "prefix": prefix, "suffix": suffix}


def _node_start(nodes, idx, full_text):
    pos = 0
    for i in range(idx):
        pos += len(nodes[i]["text"])
    return pos


def _paragraph_has_unsafe_content(paragraph_element) -> bool:
    for child in paragraph_element.iter():
        if child.tag in (W + "drawing", W + "fldChar", W + "object", W + "pict", W + "hyperlink"):
            return True
    return False


def _paragraph_is_text_only(paragraph_element) -> bool:
    for child in paragraph_element.iter():
        if child.tag in (W + "drawing", W + "fldChar", W + "object", W + "pict", W + "hyperlink"):
            return False
    return True


def build_plan(source_path, source_sha256: str, operations: list[EditOperation]) -> EditPlan:
    ir = read_docx(source_path)
    id_locator = _build_id_locator_map(ir)
    with __import__("zipfile").ZipFile(source_path) as archive:
        document_root = etree.fromstring(archive.read("word/document.xml"))
    body = document_root.find(W + "body")

    plan = EditPlan(source_sha256=source_sha256)
    bound: list[BoundEditOperation] = []

    for op in operations:
        if op.op == "append_paragraph":
            bound.append(BoundEditOperation(op=op.op, element=body, block_id=None, locator="body", text=op.text or ""))
            continue
        locator = op.source_locator
        if not locator and op.target_block_id:
            locator = id_locator.get(op.target_block_id)
        if not locator:
            plan.ready = False
            plan.blockers.append(ERROR_TARGET_NOT_FOUND + ":" + str(op.target_block_id or op.source_locator))
            continue
        element = _locate_from_body(body, locator)
        if element is None:
            plan.ready = False
            plan.blockers.append(ERROR_TARGET_NOT_FOUND + ":" + locator)
            continue
        if op.op == "replace_text":
            nodes, full_text, occurrences = _find_text_occurrences(element, op.old_text or "", op.expected_occurrences)
            if len(occurrences) == 0:
                plan.ready = False
                plan.blockers.append(ERROR_TARGET_NOT_FOUND + f":{op.old_text}@{locator}")
                continue
            if len(occurrences) != op.expected_occurrences:
                plan.ready = False
                plan.blockers.append(
                    ERROR_OCCURRENCE_MISMATCH + f":expected={op.expected_occurrences},actual={len(occurrences)}@{locator}"
                )
                continue
            match_plans = []
            for occ in occurrences:
                match_plan = _match_plan_for_occurrence(nodes, full_text, occ)
                if match_plan is None:
                    plan.ready = False
                    plan.blockers.append(ERROR_COMPLEX_BOUNDARY + f":{op.old_text}@{locator}")
                    break
                match_plans.append(match_plan)
            if match_plans is None or len(match_plans) != len(occurrences):
                continue
            bound.append(
                BoundEditOperation(
                    op=op.op, element=element, block_id=op.target_block_id, locator=locator,
                    old_text=op.old_text, new_text=op.new_text,
                    expected_occurrences=op.expected_occurrences, match_plan=match_plans,
                )
            )
        elif op.op in ("insert_paragraph_before", "insert_paragraph_after"):
            bound.append(
                BoundEditOperation(op=op.op, element=element, block_id=op.target_block_id, locator=locator, text=op.text or "")
            )
        elif op.op == "delete_paragraph":
            if not _paragraph_is_text_only(element):
                plan.ready = False
                plan.blockers.append(ERROR_UNSAFE_DELETE + ":" + locator)
                continue
            # table cell: must not be the only paragraph in its cell
            parent = element.getparent()
            if parent is not None and parent.tag == W + "tc":
                if len(parent.findall(W + "p")) <= 1:
                    plan.ready = False
                    plan.blockers.append(ERROR_UNSAFE_DELETE + ":only-cell-paragraph@" + locator)
                    continue
            # top-level: must not delete the only body paragraph
            grand = parent.getparent() if parent is not None else None
            if parent is not None and parent.tag == W + "body" and len(parent.findall(W + "p")) <= 1:
                plan.ready = False
                plan.blockers.append(ERROR_UNSAFE_DELETE + ":only-body-paragraph@" + locator)
                continue
            bound.append(BoundEditOperation(op=op.op, element=element, block_id=op.target_block_id, locator=locator))
        else:
            plan.ready = False
            plan.blockers.append("EDIT_UNSUPPORTED_OPERATION:" + op.op)

    # conflict detection
    if plan.ready:
        conflicts = _detect_conflicts(bound)
        if conflicts:
            plan.ready = False
            plan.blockers.extend(ERROR_CONFLICT + ":" + c for c in conflicts)

    plan.operations = bound
    return plan


def _detect_conflicts(bound: list[BoundEditOperation]) -> list:
    conflicts = []
    deleted = [b.locator for b in bound if b.op == "delete_paragraph"]
    for b in bound:
        if b.op in ("replace_text", "insert_paragraph_before", "insert_paragraph_after"):
            if b.locator in deleted:
                conflicts.append(f"{b.op} targets deleted paragraph {b.locator}")
    # overlapping replacements within same paragraph
    replacements_by_locator = {}
    for b in bound:
        if b.op == "replace_text":
            replacements_by_locator.setdefault(b.locator, []).append(b)
    for locator, repls in replacements_by_locator.items():
        for i in range(len(repls)):
            for j in range(i + 1, len(repls)):
                conflicts.append(f"overlapping replacements on {locator}")
    # duplicate delete
    if len(deleted) != len(set(deleted)):
        conflicts.append("duplicate delete on same paragraph")
    return conflicts


def simulate_expected_ir(ir: DocumentIR, plan: EditPlan) -> DocumentIR:
    """Apply block-level edit simulation to a fresh IR for contract validation."""
    new_ir = deepcopy(ir)

    # map block ids to live objects and their containing list
    top = new_ir.blocks

    def find(block_id, locator):
        for i, block in enumerate(top):
            if isinstance(block, ParagraphBlock):
                if block.id == block_id or (locator and block.metadata.get("source_locator") == locator):
                    return block, top, i
            if isinstance(block, TableBlock):
                for row in block.rows:
                    for cell in row:
                        for j, paragraph in enumerate(cell.blocks):
                            if paragraph.id == block_id or (locator and paragraph.metadata.get("source_locator") == locator):
                                return paragraph, cell.blocks, j
        return None, None, None

    for b in plan.operations:
        if b.op == "replace_text":
            block, container, idx = find(b.block_id, b.locator)
            if block is None:
                continue
            block.text = block.text.replace(b.old_text or "", b.new_text or "", b.expected_occurrences)
        elif b.op == "insert_paragraph_before":
            block, container, idx = find(b.block_id, b.locator)
            if block is None:
                continue
            new_p = ParagraphBlock(id=f"edit_{b.locator}_ins_before", type="paragraph", text=b.text or "")
            new_p.inline = [__import__("document_ir", fromlist=["Inline"]).Inline(type="text", text=b.text or "")]
            container.insert(idx, new_p)
        elif b.op == "insert_paragraph_after":
            block, container, idx = find(b.block_id, b.locator)
            if block is None:
                continue
            new_p = ParagraphBlock(id=f"edit_{b.locator}_ins_after", type="paragraph", text=b.text or "")
            new_p.inline = [__import__("document_ir", fromlist=["Inline"]).Inline(type="text", text=b.text or "")]
            container.insert(idx + 1, new_p)
        elif b.op == "append_paragraph":
            new_p = ParagraphBlock(id=f"edit_append_{len(top)}", type="paragraph", text=b.text or "")
            new_p.inline = [__import__("document_ir", fromlist=["Inline"]).Inline(type="text", text=b.text or "")]
            top.append(new_p)
        elif b.op == "delete_paragraph":
            block, container, idx = find(b.block_id, b.locator)
            if block is None:
                continue
            container.pop(idx)
    return new_ir
