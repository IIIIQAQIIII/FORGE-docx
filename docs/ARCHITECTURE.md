# FORGE Architecture

Version: v1.2.0

## High-Level Pipeline

```
INPUT(内容/DOCX) → Inspector → Document IR → Content Classifier
→ Format Resolver → Format Profile → Renderer → Validator → DOCX
```

Generate / Reformat / Assemble / Edit / Validate are entry points into the
same underlying pipeline.

## Subsystems

| Subsystem | Location | Responsibility |
|---|---|---|
| Format Model | `format_model.py`, `profiles/registry.py` | Profiles, inheritance, registry |
| Inspector | `document_ir.py` | Ordered, faithful DOCX → Document IR + fingerprints |
| Semantic Roles | `semantics/` | Block-level role annotation |
| Intelligence | `intelligence/` | Classifier + resolver + mappings |
| Reformat | `reformat_engine/` | Planner + source-preserving renderer + service |
| Assembly | `assembly_engine/` | Package-aware importer, preservation, service, checkpoint |
| Edit | `edit_engine/` | Planner + renderer + contract validation + service |
| Open Format | `open_format/` | FORGE_HOME, profile store, reference builder, guided, template registry |
| Utilities | `forge_utils.py`, `forge_errors.py`, `forge_version.py` | Shared helpers / error codes / version |

## Generate Pipeline

```
JSON content + Template (builtin or user docxtpl)
→ docxtpl render
→ optional media insertion / masthead adjustment / one-page guard
→ DOCX
```

## Inspect Pipeline

```
DOCX → Inspector → Document IR
     → Semantic Annotator
     → paginated block summaries (block_id, role, locator, text preview)
```

Deterministic `source_locator` paths (e.g. `body/3/table/row/1/cell/0/p/0`)
are used everywhere; text matching is never used to locate blocks.

## Reformat Pipeline

```
source DOCX
→ Inspector (Document IR)
→ Semantic Annotator
→ Content Classifier
→ Format Resolver (User > Reference > Saved > Recommendation > Default)
→ Reformat Planner
→ Source-Preserving Renderer (only w:pPr / w:rPr / w:tblPr / w:tcPr / drawing extents)
→ Preservation Validation
→ atomic output
```

Renderer never re-runs classification or semantic annotation.

## Assemble Pipeline

```
source_paths
→ resolve unified target profile
→ per source: Reformat 2.0 normalization (page chrome disabled)
→ normalized preservation PASS
→ Package-Aware Importer:
    media / hyperlink / numbering / style collisions / docPr / bookmark ids
→ assembly-level chrome (page breaks, continuous page numbers, header/footer)
→ per-item normalized→assembled preservation
→ assembly_payload_sha256
→ assembled | separate | both
```

With `checkpoint=True`, each completed item is written to
`FORGE_HOME/jobs/<job_id>/manifest.json` atomically; `resume=True` reuses
valid completed items after config/source verification.

## Edit Pipeline

```
source DOCX
→ inspect
→ source sha validation (SOURCE_CHANGED on mismatch)
→ EditPlan: bind every locator on the ORIGINAL temp document
→ conflict detection
→ apply explicit edits in user order
→ inspect temp
→ Edit Contract: Expected Payload (Source IR + simulated EditPlan)
                 vs Actual Payload (edited DOCX IR)
→ PASS → atomic output
```

Edit Contract compares per-block visible text, structure, content sequence,
table structure, media bytes, and media relationships. Unauthorized mutation
is rejected with `EDIT_CONTRACT_VIOLATION`.

## Transactional Output

All DOCX-producing subsystems work on a temporary copy and atomically move the
result into place with `os.replace`. Source files are never modified.
