$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

$venv = Join-Path $PSScriptRoot '.venv'
if (-not (Test-Path $venv)) {
  python -m venv $venv
}

$python = Join-Path $venv 'Scripts\python.exe'

& $python -m pip install -U pip
& $python -m pip install -r requirements.txt
& $python -m pip install -r requirements-dev.txt

Write-Host "Running UI tests..." -ForegroundColor Cyan
& $python ui_test_scenarios.py
