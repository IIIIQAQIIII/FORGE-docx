# Changelog

## v1.2.1

### Fixed

- Restored MIT `LICENSE` and the current public `INSTALL.md`.
- Repaired the end-user runtime ZIP allowlist and added clean-venv smoke validation.
- Restored public GitHub Actions CI and corrected README references.
- Corrected the DeepSeek Harness snippet filename in both installers.
- Preserved `LICENSE`, `INSTALL.md`, and CI in the public sanitizer allowlist.

### Changed

- Aligned Python support wording with the verified 3.10 / 3.11 / 3.12 CI matrix.
- Updated repository documentation and metadata for FORGE v1.2 capabilities.

**No product behavior changes.**

## v1.2.0

### Added

- **Open Format System** — Preset + Reference + Custom + Guided profiles.
  - `FORGE_HOME` user asset directory (profiles / templates / drafts / jobs).
  - Versioned, atomic user-profile persistence with schema validation.
  - Reference Format Analyzer with effective-formatting evidence and consensus/conflict handling.
  - Guided Format Builder sessions (≤5 questions per round).
  - Dynamic Template Registry with `register_document_template` and safe package validation.
- **Inspect MCP tool** — paginated structural inspection (`inspect_document`).
- **Edit Engine** — source-preserving, contract-validated targeted edits (`edit_document`).
- **Nested SourceLocator** — deterministic `body/<i>/table/row/<r>/cell/<c>/p/<p>` locators.
- **Assembly checkpoint/resume** — atomic manifest in `FORGE_HOME/jobs/<job_id>/`, config-change and source-change detection.
- **Centralized error model** — `forge_errors.py` stable v1.2 error codes.
- **Centralized version** — `forge_version.py` (`1.2.0`).
- Documentation: `docs/MCP_TOOLS.md`, `docs/FORMAT_PROFILES.md`, `docs/ARCHITECTURE.md`, synthetic `examples/`.

### Changed

- `generate_by_type` now resolves dynamic user templates before built-in legacy types.
- `list_format_profiles` reports `origins` (builtin | user).
- `assemble_documents` gained `checkpoint` / `job_id` / `resume` parameters (old calls unchanged).
- README rewritten as the FORGE product README.

### Compatibility

- Legacy `reformat_docx` behavior unchanged.
- Built-in `generate_docx` / `generate_by_type` / `generate_document_set` behavior unchanged for built-in types.
- `test_server.py` self-test remains 25/25.
- Public v1.1 tag and code remain untouched.

### Known Limitations

- TOC not implemented; cover support is simple text only.
- Nested table editing, table row/column structural editing, Opaque content editing, tracked changes / comments editing not implemented.
- `replace_text` rejects matches crossing hyperlink / field / image boundaries (`EDIT_COMPLEX_BOUNDARY_UNSUPPORTED`).
- `delete_paragraph` only supports text-only paragraphs.
- Large DOCX processing still loads the full package XML in memory (no streaming parser).
- Page numbers do not support PAGE/NUMPAGES combinations, roman numerals, or complex odd/even layouts.
- Assembly does not import source headers/footers; multi-section source layout is unified into assembly-level chrome.
- Performance benchmarks are development-machine reference values only, not guarantees.
