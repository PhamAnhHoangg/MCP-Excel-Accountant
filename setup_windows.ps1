# One-time setup on Windows. Run in PowerShell inside the repo folder:
#   powershell -ExecutionPolicy Bypass -File .\setup_windows.ps1
$ErrorActionPreference = "Stop"
$repo = Split-Path -Parent $MyInvocation.MyCommand.Path
$work = Join-Path $env:USERPROFILE "excel-workspace"

Write-Host "1/4 Checking Python..." -ForegroundColor Cyan
python --version
pip install -r (Join-Path $repo "requirements.txt")

Write-Host "2/4 Creating workspace folder: $work" -ForegroundColor Cyan
New-Item -ItemType Directory -Force -Path $work | Out-Null
$env:EXCEL_MCP_ROOT = $work

Write-Host "3/4 Running tests..." -ForegroundColor Cyan
python (Join-Path $repo "tests\test_server.py")

Write-Host "4/4 Done. Three ways to use it:" -ForegroundColor Green
$py = (Get-Command python).Source
$srv = Join-Path $repo "server.py"
Write-Host "`nA) No Claude app needed (command line):"
Write-Host "   cd $repo"
Write-Host "   python cli.py new-full SH-Window-model.xlsx --open"
Write-Host "`nB) Claude Code in the terminal:"
Write-Host "   claude mcp add excel-accounting --env EXCEL_MCP_ROOT=$work -- `"$py`" `"$srv`""
Write-Host "`nC) Claude Desktop: paste into %APPDATA%\Claude\claude_desktop_config.json"
$cfg = @{ mcpServers = @{ "excel-accounting" = @{ command = $py; args = @($srv); env = @{ EXCEL_MCP_ROOT = $work } } } } | ConvertTo-Json -Depth 5
Write-Host $cfg
