# FORGE v1.2.1 Maintenance Release

This is a maintenance release. **No product behavior changes.**

## Fixed

- Restored the MIT `LICENSE` file.
- Restored and updated `INSTALL.md` for the current public repository and installers.
- Repaired the runtime ZIP allowlist so all v1.2 runtime modules and templates are included.
- Added clean-venv runtime ZIP validation: install dependencies, import `server`, then run the MCP smoke test.
- Restored GitHub Actions CI for the public repository.
- Repaired README links and removed the public reference to private-only project material.
- Corrected the DeepSeek Harness configuration snippet filename.
- Updated the public sanitizer allowlist to retain `LICENSE`, `INSTALL.md`, and public CI.

## Changed

- Updated repository documentation and metadata for FORGE v1.2 capabilities.
- Aligned the supported Python statement with the verified CI matrix: Python 3.10, 3.11, and 3.12.

## Compatibility

- Existing MCP tools and product behavior are unchanged.
- Published `v1.1` and `v1.2` tags and release assets remain untouched.
