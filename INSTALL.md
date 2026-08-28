# FORGE Installation

FORGE is a local stdio MCP server. This guide covers the recommended installer
flow and the generic stdio setup.

Python support: **Python 3.10 / 3.11 / 3.12** (3.13+ is not yet verified).

## Get the Project

```bash
git clone https://github.com/IIIIQAQIIII/forge-word-docx-mcp.git
cd forge-word-docx-mcp
```

## macOS / Linux

```bash
bash install_mcp.sh
```

## Windows PowerShell

```powershell
powershell -ExecutionPolicy Bypass -File .\install_mcp.ps1
```

The installer creates a project-local `.venv`, installs runtime dependencies,
creates `outputs/`, performs a lightweight server self-check, and prints the
generic stdio MCP settings below.

## Codex

If the `codex` CLI is detected, the installer automatically registers the MCP
server under the server id `word-docx`:

```text
codex mcp add word-docx -- <venv-python> <project>/server.py
```

## DeepSeek Harness

If the `dsh` CLI is detected, the installer writes a ready-to-use plugin row to:

```text
deepseek-harness-forge-word-docx-mcp.yml
```

Copy that row into the active profile's `cordis.yml` plugins list.

## Generic stdio MCP

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

## FORGE_HOME

User assets are stored under `~/.forge-docx-mcp/` by default:

```text
profiles/    persisted format profiles
templates/   registered user templates
drafts/      guided profile sessions
jobs/        assembly checkpoint/resume workspaces
```

Set the `FORGE_HOME` environment variable to override this location.

## Outputs

Generated and transformed documents are written to `outputs/` unless an
absolute output path is provided.

## Development

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/pip install -r requirements-dev.txt
.venv/bin/python -m pytest -q
.venv/bin/python test_server.py
```
