"""Mission 05 — Per-item import preservation + assembly payload fingerprint."""

from __future__ import annotations

import hashlib
from typing import Any

from document_ir import DocumentIR, OpaqueBlock, ParagraphBlock, TableBlock, _block_sequence_events, _compute_fingerprint
from reformat_engine.renderer import _table_structure_signature


def _media_id_to_sha(ir: DocumentIR) -> dict:
    return {item.media_id: item.sha256 for item in ir.media}


def _media_shas_in_order(blocks, sha_map: dict) -> list:
    shas = []

    def walk_paragraph(paragraph: ParagraphBlock) -> None:
        for inline in paragraph.inline:
            if inline.type == "image":
                shas.append(sha_map.get(inline.media_id, inline.media_id or ""))

    for block in blocks:
        if isinstance(block, ParagraphBlock):
            walk_paragraph(block)
        elif isinstance(block, TableBlock):
            for row in block.rows:
                for cell in row:
                    for paragraph in cell.blocks:
                        walk_paragraph(paragraph)
        elif isinstance(block, OpaqueBlock):
            pass
    return shas


def _sequence_events_with_sha(blocks, sha_map: dict) -> list:
    events = []
    for block in blocks:
        for event in _block_sequence_events(block):
            if event.startswith("img:"):
                media_id = event[4:]
                events.append("img:" + sha_map.get(media_id, media_id))
            else:
                events.append(event)
    return events


def _partial_ir(blocks) -> DocumentIR:
    return DocumentIR(source="", blocks=list(blocks), media=[])


def verify_item_preservation(
    normalized_ir: DocumentIR,
    assembled_ir: DocumentIR,
    assembled_blocks: list,
) -> dict[str, Any]:
    """Compare a normalized item IR against its imported block slice in the master.

    Media ids in the assembled slice are resolved against the full assembled IR
    media table (relationship IDs legitimately change during package import).
    """
    norm_partial = _partial_ir(normalized_ir.blocks)
    asm_partial = _partial_ir(assembled_blocks)

    norm_fp = _compute_fingerprint(norm_partial)
    asm_fp = _compute_fingerprint(asm_partial)

    norm_sha_map = _media_id_to_sha(normalized_ir)
    asm_sha_map = _media_id_to_sha(assembled_ir)

    text_ok = asm_fp.text_sha256 == norm_fp.text_sha256
    structure_ok = asm_fp.structure_sha256 == norm_fp.structure_sha256
    sequence_ok = _sequence_events_with_sha(normalized_ir.blocks, norm_sha_map) == _sequence_events_with_sha(
        assembled_blocks, asm_sha_map
    )
    table_ok = _table_structure_signature(norm_partial) == _table_structure_signature(asm_partial)

    norm_shas = _media_shas_in_order(normalized_ir.blocks, norm_sha_map)
    asm_shas = _media_shas_in_order(assembled_blocks, asm_sha_map)
    media_sha_ok = norm_shas == asm_shas
    media_rel_ok = norm_shas == asm_shas

    passed = all((text_ok, structure_ok, sequence_ok, table_ok, media_sha_ok, media_rel_ok))
    return {
        "passed": bool(passed),
        "text": bool(text_ok),
        "structure": bool(structure_ok),
        "sequence": bool(sequence_ok),
        "table_structure": bool(table_ok),
        "media_sha256": bool(media_sha_ok),
        "media_relationships": bool(media_rel_ok),
    }


def compute_assembly_payload_sha256(items: list[dict[str, Any]]) -> str:
    """Fingerprint over item content + order.

    Assembly-generated page numbers / page breaks are not part of this payload.
    """
    h = hashlib.sha256()
    for item in items:
        fp = item.get("normalized_fingerprints") or {}
        line = (
            f"{item.get('index')}:{item.get('source_sha256')}:"
            f"{fp.get('text_sha256', '')}:{fp.get('structure_sha256', '')}:"
            f"{fp.get('media_sha256', '')}:{fp.get('content_sequence_sha256', '')}"
        )
        h.update(line.encode("utf-8"))
        h.update(b"\n")
    return h.hexdigest()
