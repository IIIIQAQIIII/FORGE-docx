"""Mission 07 — Edit Contract validation.

Expected payload = Source IR + simulated EditPlan (block level).
Actual payload = Inspector output of the edited document.

We do not require full fingerprint equality (edits intentionally change text);
we compare every block's visible text sequence, document structure hash, table
structure, media bytes, and media relationships against the simulated payload.
"""

from __future__ import annotations

from typing import Any

from document_ir import DocumentIR, OpaqueBlock, ParagraphBlock, TableBlock, _compute_fingerprint
from reformat_engine.renderer import _media_relationship_signature, _table_structure_signature


def _all_paragraph_texts(ir: DocumentIR) -> list:
    texts = []
    for block in ir.blocks:
        if isinstance(block, ParagraphBlock):
            texts.append(block.text)
        elif isinstance(block, TableBlock):
            for row in block.rows:
                for cell in row:
                    for paragraph in cell.blocks:
                        texts.append(paragraph.text)
        elif isinstance(block, OpaqueBlock):
            texts.append(block.extracted_text)
    return texts


def _block_structure_events(ir: DocumentIR) -> list:
    events = []
    for block in ir.blocks:
        if isinstance(block, ParagraphBlock):
            events.append("P")
        elif isinstance(block, TableBlock):
            events.append(f"T({len(block.rows)},{len(block.rows[0]) if block.rows else 0})")
        elif isinstance(block, OpaqueBlock):
            events.append(f"O({block.xml_tag})")
    return events


def payload_from_ir(ir: DocumentIR) -> dict:
    fp = _compute_fingerprint(ir)
    return {
        "texts": _all_paragraph_texts(ir),
        "text_sha256": fp.text_sha256,
        "structure_sha256": fp.structure_sha256,
        "sequence_events": _block_structure_events(ir),
        "media_shas": [m.sha256 for m in ir.media],
        "media_relationships": _media_relationship_signature(ir),
        "table_structure": _table_structure_signature(ir),
    }


def validate_edit_contract(expected_ir: DocumentIR, actual_ir: DocumentIR) -> dict:
    expected = payload_from_ir(expected_ir)
    actual = payload_from_ir(actual_ir)

    texts_ok = expected["texts"] == actual["texts"]
    structure_ok = expected["structure_sha256"] == actual["structure_sha256"]
    text_sha_ok = expected["text_sha256"] == actual["text_sha256"]
    media_ok = expected["media_shas"] == actual["media_shas"]
    media_rel_ok = expected["media_relationships"] == actual["media_relationships"]
    table_ok = expected["table_structure"] == actual["table_structure"]

    passed = all((texts_ok, structure_ok, text_sha_ok, media_ok, media_rel_ok, table_ok))
    return {
        "passed": bool(passed),
        "texts": bool(texts_ok),
        "structure": bool(structure_ok),
        "text_sha256": bool(text_sha_ok),
        "media": bool(media_ok),
        "media_relationships": bool(media_rel_ok),
        "table_structure": bool(table_ok),
    }
