#!/usr/bin/env bash
set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="$PROJECT_DIR/.venv"
PYTHON_BIN="$VENV_DIR/bin/python"
HARNESS_SNIPPET="$PROJECT_DIR/deepseek-harness-forge-forge-word-docx-mcp.yml"

printf '\nFORGE universal installer\n'
printf 'Project: %s\n\n' "$PROJECT_DIR"

if command -v python3 >/dev/null 2>&1; then
  SYSTEM_PYTHON="$(command -v python3)"
elif command -v python >/dev/null 2>&1; then
  SYSTEM_PYTHON="$(command -v python)"
else
  echo "Error: Python 3 was not found. Please install Python 3.10 or newer first."
  exit 1
fi

"$SYSTEM_PYTHON" - <<'PY'
import sys
if sys.version_info < (3, 10):
    raise SystemExit(f"Python 3.10+ is required; current version is {sys.version.split()[0]}")
print(f"Python: {sys.version.split()[0]}")
PY

if [ ! -f "$PROJECT_DIR/server.py" ] || [ ! -f "$PROJECT_DIR/requirements.txt" ] || [ ! -d "$PROJECT_DIR/templates" ]; then
  echo "Error: release files are incomplete. server.py, requirements.txt and templates/ are required."
  exit 1
fi

mkdir -p "$PROJECT_DIR/outputs"

if [ ! -x "$PYTHON_BIN" ]; then
  echo "Creating virtual environment..."
  "$SYSTEM_PYTHON" -m venv "$VENV_DIR"
else
  echo "Reusing existing virtual environment..."
fi

"$PYTHON_BIN" -m pip install --upgrade pip
"$PYTHON_BIN" -m pip install -r "$PROJECT_DIR/requirements.txt"

"$PYTHON_BIN" - <<PY
import sys
sys.path.insert(0, r"$PROJECT_DIR")
import server
assert server.TEMPLATES_DIR.is_dir()
print(f"Loaded MCP server; templates: {len(list(server.TEMPLATES_DIR.glob('*.docx')))}")
PY

configured=0

if command -v codex >/dev/null 2>&1; then
  echo "Configuring Codex MCP..."
  if codex mcp get word-docx >/dev/null 2>&1; then
    codex mcp remove word-docx >/dev/null
  fi
  codex mcp add word-docx -- "$PYTHON_BIN" "$PROJECT_DIR/server.py"
  echo "Configured Codex: word-docx"
  configured=1
fi

cat > "$HARNESS_SNIPPET" <<EOF
# Add this plugin row to the plugins list of the DeepSeek Harness profile's cordis.yml.
# DeepSeek Harness uses @deepseek-ai/dsh-mcp-client to bridge local stdio MCP servers.
- id: mcp-word-docx
  name: '@deepseek-ai/dsh-mcp-client'
  config:
    serverName: word-docx
    transport: stdio
    command: '$PYTHON_BIN'
    args:
      - '$PROJECT_DIR/server.py'
    cwd: '$PROJECT_DIR'
    toolCallTimeoutMs: 60000
    failOnStartupError: false
EOF

if command -v dsh >/dev/null 2>&1; then
  echo "DeepSeek Harness detected."
  echo "Harness config snippet generated: $HARNESS_SNIPPET"
  echo "Copy that plugin row into the active profile's cordis.yml, then reload/restart DSH."
  configured=1
fi

echo
echo "Runtime installation complete."
echo "Generated documents: $PROJECT_DIR/outputs"
echo "Generic stdio MCP configuration:"
echo "  command: $PYTHON_BIN"
echo "  args:    $PROJECT_DIR/server.py"
if [ "$configured" -eq 0 ]; then
  echo "No supported client CLI was detected; use the generic stdio settings above."
fi
