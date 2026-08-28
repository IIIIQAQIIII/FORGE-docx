"""Reference DOCX → Draft FormatProfile builder.

Faithful Inspector + Semantic Annotator + effective formatting evidence.
Never stores document body text; only formatting statistics.
"""

from __future__ import annotations

from collections import Counter
from pathlib import Path
from typing import Any, Optional

from lxml import etree

from document_ir import W, W_NS, DocumentIR, ParagraphBlock, read_docx
from open_format.config import DOMINANCE_RATIO_THRESHOLD, REFERENCE_ROLE_SLOTS
from profiles import registry as profile_registry
from semantics.annotator import annotate_document

EXTRACTED_PROPERTIES = (
    "font",
    "size_pt",
    "bold",
    "italic",
    "align",
    "line_spacing_pt",
    "space_before_pt",
    "space_after_pt",
    "first_line_chars",
    "left_chars",
    "right_chars",
)


class EffectiveFormatReader:
    """Resolve direct + Word style inheritance to effective pPr/rPr values."""

    def __init__(self, document_root, styles_root):
        self.document_root = document_root
        self.styles_root = styles_root
        self.styles_by_id = {}
        if styles_root is not None:
            for style in styles_root.findall(W + "style"):
                self.styles_by_id[style.get(W + "styleId")] = style
        self.doc_defaults = styles_root.find(W + "docDefaults") if styles_root is not None else None

    def locate(self, locator: str):
        if not locator:
            return None
        parts = locator.split("/")
        if parts[0] != "body":
            return None
        current = self.document_root.find(W + "body")
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

    def _style_chain(self, style_id: str) -> list:
        chain = []
        seen = set()
        current_id = style_id
        while current_id and current_id not in seen:
            seen.add(current_id)
            style = self.styles_by_id.get(current_id)
            if style is None:
                break
            chain.append(style)
            based_on = style.find(W + "basedOn")
            current_id = based_on.get(W + "val") if based_on is not None else None
        chain.reverse()
        return chain

    @staticmethod
    def _apply_rpr(rpr, props: dict) -> None:
        if rpr is None:
            return
        rfonts = rpr.find(W + "rFonts")
        if rfonts is not None:
            props["font"] = rfonts.get(W + "eastAsia") or rfonts.get(W + "ascii") or props.get("font")
        sz = rpr.find(W + "sz")
        if sz is not None:
            try:
                props["size_pt"] = int(sz.get(W + "val")) / 2.0
            except (TypeError, ValueError):
                pass
        bold = rpr.find(W + "b")
        if bold is not None:
            props["bold"] = bold.get(W + "val") not in ("0", "false")
        italic = rpr.find(W + "i")
        if italic is not None:
            props["italic"] = italic.get(W + "val") not in ("0", "false")

    @staticmethod
    def _apply_ppr(ppr, props: dict) -> None:
        if ppr is None:
            return
        jc = ppr.find(W + "jc")
        if jc is not None:
            val = jc.get(W + "val")
            props["align"] = "both" if val in ("both", "justify") else val
        spacing = ppr.find(W + "spacing")
        if spacing is not None:
            line = spacing.get(W + "line")
            line_rule = spacing.get(W + "lineRule")
            if line and line_rule in ("exact", "atLeast"):
                try:
                    props["line_spacing_pt"] = int(line) / 20.0
                except ValueError:
                    pass
            before = spacing.get(W + "before")
            if before:
                try:
                    props["space_before_pt"] = int(before) / 20.0
                except ValueError:
                    pass
            after = spacing.get(W + "after")
            if after:
                try:
                    props["space_after_pt"] = int(after) / 20.0
                except ValueError:
                    pass
        ind = ppr.find(W + "ind")
        if ind is not None:
            for key in ("first_line_chars", "left_chars", "right_chars"):
                val = ind.get(W + key)
                if val:
                    try:
                        props[key] = int(val)
                    except ValueError:
                        pass

    def effective_props(self, paragraph_element) -> dict:
        props = {}
        if self.doc_defaults is not None:
            ppr_default = self.doc_defaults.find(W + "pPrDefault")
            if ppr_default is not None:
                self._apply_ppr(ppr_default.find(W + "pPr"), props)
            rpr_default = self.doc_defaults.find(W + "rPrDefault")
            if rpr_default is not None:
                self._apply_rpr(rpr_default.find(W + "rPr"), props)
        ppr = paragraph_element.find(W + "pPr")
        pstyle_id = None
        if ppr is not None:
            pstyle = ppr.find(W + "pStyle")
            if pstyle is not None:
                pstyle_id = pstyle.get(W + "val")
        for style in self._style_chain(pstyle_id) if pstyle_id else []:
            self._apply_ppr(style.find(W + "pPr"), props)
            self._apply_rpr(style.find(W + "rPr"), props)
        if ppr is not None:
            self._apply_ppr(ppr, props)
            self._apply_rpr(ppr.find(W + "rPr"), props)
        first_run = paragraph_element.find(W + "r")
        if first_run is not None:
            self._apply_rpr(first_run.find(W + "rPr"), props)
        return props


def analyze_reference_docx(
    reference_path: str | Path,
    base_profile_id: str = "generic_document",
    profile_id: Optional[str] = None,
    name: Optional[str] = None,
) -> dict[str, Any]:
    """Analyze a reference DOCX and build a Draft FormatProfile.

    The returned draft contains no document text, only formatting statistics.
    """
    path = Path(reference_path)
    ir: DocumentIR = read_docx(path)
    annotate_document(ir)

    try:
        profile_registry.resolve_profile(base_profile_id)
    except KeyError:
        return {"status": "error", "error": "PROFILE_NOT_FOUND", "reason": f"base profile 不存在: {base_profile_id}"}

    with __import__("zipfile").ZipFile(path) as archive:
        document_root = etree.fromstring(archive.read("word/document.xml"))
        try:
            styles_root = etree.fromstring(archive.read("word/styles.xml"))
        except KeyError:
            styles_root = None

    reader = EffectiveFormatReader(document_root, styles_root)

    # Page evidence from section 1
    page_rules = {}
    if ir.sections:
        section = ir.sections[0]
        page_map = {
            "width_cm": section.page_width,
            "height_cm": section.page_height,
            "top_cm": section.margin_top,
            "bottom_cm": section.margin_bottom,
            "left_cm": section.margin_left,
            "right_cm": section.margin_right,
            "header_distance_cm": section.header_distance,
            "footer_distance_cm": section.footer_distance,
        }
        page_rules = {k: round(v, 2) for k, v in page_map.items() if v is not None}

    rules = {}
    evidence = {}
    conflicts = []
    if page_rules:
        rules["page"] = page_rules
        evidence["page"] = {"sample_count": 1, "dominant_count": 1, "dominance_ratio": 1.0, "confidence": 1.0}

    role_paragraphs = {}
    for block in ir.blocks:
        if isinstance(block, ParagraphBlock) and block.semantic_role in REFERENCE_ROLE_SLOTS:
            role_paragraphs.setdefault(block.semantic_role, []).append(block)

    for role, slot in REFERENCE_ROLE_SLOTS.items():
        samples = role_paragraphs.get(role, [])
        slot_evidence = {"sample_count": len(samples), "source": "inherited" if not samples else "evidence"}
        if not samples:
            evidence[slot] = slot_evidence
            continue

        prop_values = {prop: [] for prop in EXTRACTED_PROPERTIES}
        for block in samples:
            element = reader.locate(block.metadata.get("source_locator", ""))
            if element is None:
                continue
            props = reader.effective_props(element)
            for prop in EXTRACTED_PROPERTIES:
                if props.get(prop) is not None:
                    prop_values[prop].append(props[prop])

        adopted = {}
        slot_evidence["dominant_count"] = 0
        slot_evidence["adopted_properties"] = []
        for prop, values in prop_values.items():
            if not values:
                continue
            counter = Counter(values)
            dominant_value, dominant_count = counter.most_common(1)[0]
            ratio = dominant_count / len(values)
            if ratio >= DOMINANCE_RATIO_THRESHOLD:
                adopted[prop] = dominant_value
                slot_evidence["adopted_properties"].append(prop)
                slot_evidence["dominant_count"] = max(slot_evidence["dominant_count"], dominant_count)
            else:
                conflicts.append(
                    {
                        "slot": slot,
                        "property": prop,
                        "candidates": [
                            {"value": value, "count": count, "ratio": round(count / len(values), 4)}
                            for value, count in counter.most_common()
                        ],
                    }
                )
        if adopted:
            rules[slot] = adopted
        ratio_all = (slot_evidence["dominant_count"] / len(samples)) if slot_evidence["dominant_count"] else 0.0
        slot_evidence["dominance_ratio"] = round(ratio_all, 4)
        slot_evidence["confidence"] = round(ratio_all, 4)
        evidence[slot] = slot_evidence

    status = "needs_review" if conflicts else "ok"
    draft = {
        "schema_version": 1,
        "profile_id": profile_id or Path(path).stem,
        "name": name or f"Reference {Path(path).stem}",
        "description": "reference learning draft",
        "source": "reference",
        "inherits": base_profile_id,
        "rules": rules,
        "evidence": evidence,
    }
    return {"status": status, "draft": draft, "conflicts": conflicts}
