$ErrorActionPreference = "Stop"

$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvDir = Join-Path $ProjectDir ".venv"
$PythonBin = Join-Path $VenvDir "Scripts\python.exe"
$ServerPath = Join-Path $ProjectDir "server.py"
$RequirementsPath = Join-Path $ProjectDir "requirements.txt"
$TemplatesDir = Join-Path $ProjectDir "templates"
$OutputsDir = Join-Path $ProjectDir "outputs"

Write-Host ""
Write-Host "FORGE installer"
Write-Host "Project: $ProjectDir"
Write-Host ""

if (-not (Test-Path $ServerPath) -or -not (Test-Path $RequirementsPath) -or -not (Test-Path $TemplatesDir)) {
    throw "Release files are incomplete. server.py, requirements.txt and templates\ are required."
}

$PythonCommand = $null
$PythonArgs = @()
if (Get-Command py -ErrorAction SilentlyContinue) {
    $PythonCommand = "py"
    $PythonArgs = @("-3")
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
    $PythonCommand = "python"
} else {
    throw "Python 3 was not found. Please install Python 3.10 or newer first."
}

& $PythonCommand @PythonArgs -c "import sys; assert sys.version_info >= (3,10), f'Python 3.10+ is required; current version is {sys.version.split()[0]}'; print('Python:', sys.version.split()[0])"
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

New-Item -ItemType Directory -Force -Path $OutputsDir | Out-Null

if (-not (Test-Path $PythonBin)) {
    Write-Host "Creating virtual environment..."
    & $PythonCommand @PythonArgs -m venv $VenvDir
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
} else {
    Write-Host "Reusing existing virtual environment..."
}

& $PythonBin -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
& $PythonBin -m pip install -r $RequirementsPath
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

& $PythonBin -c "import sys; sys.path.insert(0, r'$ProjectDir'); import server; assert server.TEMPLATES_DIR.is_dir(); print('Loaded MCP server; templates:', len(list(server.TEMPLATES_DIR.glob('*.docx'))))"
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

if (Get-Command codex -ErrorAction SilentlyContinue) {
    Write-Host "Configuring Codex MCP..."
    & codex mcp get word-docx *> $null
    if ($LASTEXITCODE -eq 0) {
        & codex mcp remove word-docx *> $null
    }
    & codex mcp add word-docx -- $PythonBin $ServerPath
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
    Write-Host ""
    Write-Host "Installed: word-docx"
    Write-Host "Restart Codex to load the MCP server."
} else {
    Write-Host ""
    Write-Host "Dependencies installed successfully. Codex CLI was not found."
    Write-Host "For another MCP client, use:"
    Write-Host "  command: $PythonBin"
    Write-Host "  args:    $ServerPath"
}

Write-Host "Generated documents will be saved in: $OutputsDir"
