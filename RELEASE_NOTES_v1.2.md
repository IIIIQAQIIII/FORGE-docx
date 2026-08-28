# FORGE v1.2 Release Notes

**Preserve the content. Reforge the format.**

FORGE v1.2 turns a template generator into a schema-driven DOCX formatting and
transformation engine for AI agents.

FORGE v1.2 将原有模板生成器升级为面向 AI Agent 的 Schema 驱动 DOCX
格式化与文档转换引擎。它支持生成、检查、内容保真重排、批量汇编和安全定点编辑，
输出仍是可继续编辑的 Word 文档；除非执行明确的编辑指令，否则不会改写用户原文。

## Major Capabilities

- **Content-preserving DOCX reformatting** — source-preserving renderer plus hard preservation validation (text, structure, media, sequence, tables, media relationships).
- **Batch assembly** — normalize + package-aware merge for many DOCX files, with per-item preservation and assembly payload fingerprints.
- **Reference / Custom / Guided formatting profiles** — learn formats from a sample, define overrides, or walk through a guided session; all persisted in `FORGE_HOME`.
- **Dynamic templates** — register user DOCX templates without touching core code.
- **Source-preserving targeted editing** — explicit edits protected by source-version checks and an Edit Contract.
- **Checkpoint/resume for large assemblies** — interrupted batches resume from completed items only after config/source verification.

## Honest Known Limitations

- TOC, nested table editing, tracked-changes editing, and complex field editing are not implemented.
- Replacements cannot cross hyperlink / field / image boundaries.
- Page numbers do not support PAGE/NUMPAGES, roman numerals, or complex odd/even layouts.
- Large DOCX processing still loads the package XML in memory.
- Benchmarks are development-machine reference values, not guarantees.

## Compatibility

- Legacy `reformat_docx` and built-in generation behavior are unchanged.
- Public v1.1 code and tag remain untouched.
