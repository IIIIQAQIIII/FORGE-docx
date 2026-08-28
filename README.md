# FORGE

**Format-Oriented Rendering & Generation Engine**

[![Version](https://img.shields.io/badge/version-v1.2.0-111111?style=flat-square)](VERSION)
[![Python](https://img.shields.io/badge/Python-3.11-3776AB?style=flat-square&logo=python&logoColor=white)](requirements.txt)
[![MCP](https://img.shields.io/badge/MCP-stdio-6F42C1?style=flat-square)](README.md)
[![DOCX](https://img.shields.io/badge/DOCX-content--preserving-2B579A?style=flat-square)](templates)
[![GitHub Release](https://img.shields.io/github/v/release/IIIIQAQIIII/forge-word-docx-mcp?style=flat-square&label=release)](https://github.com/IIIIQAQIIII/forge-word-docx-mcp/releases/latest)

> **Forge documents that don’t break.**
>
> **AI writes. FORGE holds the format.**
>
> **Preserve the content. Reforge the format.**

Version / 当前版本：**v1.2.0**

FORGE is a local stdio MCP engine for generating, inspecting, reformatting,
assembling, and safely editing editable Word DOCX files. It uses schemas,
format profiles, source-preserving rendering, and validation to keep document
content stable while transforming its presentation.

FORGE 是一个本地运行的 stdio MCP 文档引擎，可生成、检查、重排、汇编和安全定点编辑
可继续修改的 Word DOCX 文件。它通过 Schema、格式方案、源文件保真渲染和内容校验，
在转换文档呈现形式的同时保护用户原文。

核心思路：**AI 负责理解需求和生成内容，FORGE 负责守住内容、结构与格式。**

## Why FORGE / 为什么选择 FORGE

AI 很擅长写内容，但复杂 Word 文档最容易在标题、段距、页边距、页码、表格、图片、
多文档汇编和模板结构上发生格式漂移。FORGE 让格式方案与模板成为版式依据，
并用内容指纹和保真校验阻止意外丢字、改字或结构损坏。

```text
AI instruction / Source DOCX
             ↓
     Inspect + Semantic Roles
             ↓
  Format Profile / DOCX Template
             ↓
 Generate / Reformat / Assemble / Edit
             ↓
   Preservation Validation
             ↓
       Editable DOCX
```

> 本公开版本中的单位名称、人员名称和示例内容均为虚构或脱敏信息，
> 不代表任何真实机构或个人。

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
