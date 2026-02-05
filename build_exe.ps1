$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

$venv = Join-Path $PSScriptRoot '.venv'
$py = 'python'

if (-not (Test-Path $venv)) {
  & $py -m venv $venv
}

$python = Join-Path $venv 'Scripts\python.exe'

& $python -m pip install -U pip wheel
& $python -m pip install -r requirements.txt
& $python -m pip install pyinstaller

# Clean previous builds
if (Test-Path 'build') { Remove-Item -Recurse -Force 'build' }
if (Test-Path 'dist') { Remove-Item -Recurse -Force 'dist' }
if (Test-Path '*.spec') { Remove-Item -Force '*.spec' }

& $python -m PyInstaller --noconsole --onefile --name sow_merge_tool sow_merge_tool.py

Write-Host "Build complete: $(Join-Path $PSScriptRoot 'dist\sow_merge_tool.exe')" -ForegroundColor Green
