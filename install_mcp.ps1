$ErrorActionPreference = 'Stop'

$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvDir = Join-Path $ProjectDir '.venv'
$PythonBin = Join-Path $VenvDir 'Scripts\python.exe'
$HarnessSnippet = Join-Path $ProjectDir 'deepseek-harness-forge-word-docx-mcp.yml'

Write-Host "`nFORGE universal installer"
Write-Host "Project: $ProjectDir`n"

$SystemPython = $null
foreach ($candidate in @('python', 'py')) {
    try {
        $cmd = Get-Command $candidate -ErrorAction Stop
        $SystemPython = $candidate
        break
    } catch {}
}
if (-not $SystemPython) {
    throw 'Python 3.10, 3.11 or 3.12 was not found.'
}

if ($SystemPython -eq 'py') {
    & py -3 -c "import sys; assert (3,10) <= sys.version_info < (3,13), sys.version; print('Python:', sys.version.split()[0])"
} else {
    & python -c "import sys; assert (3,10) <= sys.version_info < (3,13), sys.version; print('Python:', sys.version.split()[0])"
}

if (-not (Test-Path (Join-Path $ProjectDir 'server.py')) -or
    -not (Test-Path (Join-Path $ProjectDir 'requirements.txt')) -or
    -not (Test-Path (Join-Path $ProjectDir 'templates'))) {
    throw 'Release files are incomplete. server.py, requirements.txt and templates\ are required.'
}

New-Item -ItemType Directory -Force -Path (Join-Path $ProjectDir 'outputs') | Out-Null

if (-not (Test-Path $PythonBin)) {
    Write-Host 'Creating virtual environment...'
    if ($SystemPython -eq 'py') { & py -3 -m venv $VenvDir } else { & python -m venv $VenvDir }
} else {
    Write-Host 'Reusing existing virtual environment...'
}

& $PythonBin -m pip install --upgrade pip
& $PythonBin -m pip install -r (Join-Path $ProjectDir 'requirements.txt')
& $PythonBin -c "import sys; sys.path.insert(0, r'$ProjectDir'); import server; assert server.TEMPLATES_DIR.is_dir(); print('Loaded MCP server; templates:', len(list(server.TEMPLATES_DIR.glob('*.docx'))))"

$configured = $false
if (Get-Command codex -ErrorAction SilentlyContinue) {
    Write-Host 'Configuring Codex MCP...'
    & codex mcp get word-docx *> $null
    if ($LASTEXITCODE -eq 0) { & codex mcp remove word-docx *> $null }
    & codex mcp add word-docx -- $PythonBin (Join-Path $ProjectDir 'server.py')
    Write-Host 'Configured Codex: word-docx'
    $configured = $true
}

$serverPath = (Join-Path $ProjectDir 'server.py').Replace("'", "''")
$pythonPath = $PythonBin.Replace("'", "''")
$cwdPath = $ProjectDir.Replace("'", "''")
$snippet = @"
# Add this plugin row to the plugins list of the DeepSeek Harness profile's cordis.yml.
# DeepSeek Harness uses @deepseek-ai/dsh-mcp-client to bridge local stdio MCP servers.
- id: mcp-word-docx
  name: '@deepseek-ai/dsh-mcp-client'
  config:
    serverName: word-docx
    transport: stdio
    command: '$pythonPath'
    args:
      - '$serverPath'
    cwd: '$cwdPath'
    toolCallTimeoutMs: 60000
    failOnStartupError: false
"@
Set-Content -Path $HarnessSnippet -Value $snippet -Encoding UTF8

if (Get-Command dsh -ErrorAction SilentlyContinue) {
    Write-Host 'DeepSeek Harness detected.'
    Write-Host "Harness config snippet generated: $HarnessSnippet"
    Write-Host "Copy that plugin row into the active profile's cordis.yml, then reload/restart DSH."
    $configured = $true
}

Write-Host "`nRuntime installation complete."
Write-Host "Generated documents: $(Join-Path $ProjectDir 'outputs')"
Write-Host 'Generic stdio MCP configuration:'
Write-Host "  command: $PythonBin"
Write-Host "  args:    $(Join-Path $ProjectDir 'server.py')"
if (-not $configured) {
    Write-Host 'No supported client CLI was detected; use the generic stdio settings above.'
}
