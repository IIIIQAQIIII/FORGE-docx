# FORGE

**Format-Oriented Rendering & Generation Engine**

> **Preserve the content. Reforge the format.**

FORGE is a schema-driven DOCX formatting and transformation engine for AI agents.
It generates, inspects, reformats, assembles, and safely edits Word documents —
without rewriting the user's content.

FORGE 是面向 AI Agent 的、由 Schema 驱动的 DOCX 格式化与文档转换引擎。
它可以生成、检查、重排、汇编和安全定点编辑可继续修改的 Word 文档，
并通过内容指纹与保真校验确保用户原文不被擅自改写。

FORGE is **not** a text-generation system. The server never summarizes, rewrites,
or "improves" document text. It only changes formatting and document structure,
under explicit instructions, with hard preservation guarantees.

FORGE **不是**文本生成系统。除非用户明确执行编辑操作，否则它不会总结、
润色或重写正文，只按照明确指令调整格式与文档呈现结构。

---

## Quick Start

```bash
cd forge-word-docx-mcp
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt        # runtime deps
.venv/bin/pip install -r requirements-dev.txt    # tests (optional)

.venv/bin/python server.py                       # start MCP server (stdio)
.venv/bin/python test_server.py                  # self-test (25 checks)
.venv/bin/python -m pytest -q                    # full test suite
```

Python support: **3.11.x** (verified in the release environment; see
`INSTALL.md` for installer-based setup).

Configure the MCP client with:

```json
{
  "mcpServers": {
    "forge": {
      "command": "/absolute/path/to/forge-word-docx-mcp/.venv/bin/python",
      "args": ["/absolute/path/to/forge-word-docx-mcp/server.py"]
    }
  }
}
```

---

## Core Capabilities

| Capability | Tool | Description |
|---|---|---|
| Generate | `generate_docx` / `generate_by_type` / `generate_document_set` | Fill built-in or user-registered DOCX templates. |
| Inspect | `inspect_document` | Paginated structural view of a DOCX (block ids, roles, locators). |
| Reformat | `reformat_document` | Source-preserving reformat with format intelligence + preservation validation. |
| Assemble | `assemble_documents` | Batch normalize + package-aware merge + preservation + checkpoint/resume. |
| Edit | `edit_document` | Explicit, source-preserving, contract-validated targeted edits. |
| Format Intelligence | `resolve_document_format` / `recommend_document_type` | Classify content intent and resolve a target format. |
| Open Format System | `list_format_profiles` / `analyze_reference_format` / `create_format_profile` / `register_document_template` … | Preset / Reference / Custom / Guided profiles and dynamic templates. |

Legacy compatibility:

- `reformat_docx` is the **legacy** template-based reformat path and remains
  unchanged.
- Built-in legacy generation types and templates remain unchanged.

---

## Content Preservation (high level)

FORGE treats the user's content as immutable payload:

1. **Inspector** reads the DOCX into an ordered Document IR with deterministic
   `source_locator` paths (never text matching).
2. **Source-Preserving Renderer** only patches formatting XML
   (`w:pPr` / `w:rPr` / `w:tblPr` / `w:tcPr` / drawing extents).
3. **Preservation Validation** compares content fingerprints before/after:
   visible text, structure, media bytes, content sequence, table structure,
   and media relationships.
4. If any fingerprint fails, no output is produced.

Edit intentionally allows text changes, so it validates against an
**Edit Contract**: Source IR + EditPlan → Expected Payload vs Actual Payload.
Unauthorized changes are rejected with `EDIT_CONTRACT_VIOLATION`.

---

## Format Model

```
ContentIntent != SemanticRole != FormatProfile != Template
```

- **ContentIntent** says *what the content is* (plan, summary, notice…).
- **SemanticRole** says *what a block does* (title, heading, body, caption…).
- **FormatProfile** says *how to format* each role.
- **Template** is only a DOCX skeleton for generation.

Profiles can be:

- **Preset** — built-in profiles (generic, official, academic, weekly, …).
- **Reference** — learned from a sample DOCX (`analyze_reference_format`).
- **Custom** — explicit overrides on a base profile.
- **Guided** — step-by-step sessions that only ask the key differences.

User profiles persist under `FORGE_HOME/profiles/` (default
`~/.forge-docx-mcp/`). Dynamic templates persist under `FORGE_HOME/templates/`.

See `docs/FORMAT_PROFILES.md` and `docs/ARCHITECTURE.md` for details.

---

## Examples

Synthetic examples live in `examples/`:

- `example_reformat.py`
- `example_assemble.py`
- `example_reference_profile.py`
- `example_edit.py`

All examples use generic synthetic data only.

---

## Documentation

- `docs/MCP_TOOLS.md` — every MCP tool, parameters, results, common errors.
- `docs/FORMAT_PROFILES.md` — format model, profiles, FORGE_HOME.
- `docs/ARCHITECTURE.md` — pipeline diagrams and subsystem responsibilities.
- `CHANGELOG.md` — release notes, including v1.2.0.

---

## Release

- Current release: **v1.2.0**
- Status: **stable**

Known limitations are documented in `CHANGELOG.md` and `PROJECT_SUMMARY.md`.
No performance guarantees are implied by the development-machine benchmarks.
