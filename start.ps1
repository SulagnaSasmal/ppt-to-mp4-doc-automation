$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

$pythonExe = Join-Path $projectRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $pythonExe)) {
    Write-Error "Virtual environment not found at .venv. Create it first (python -m venv .venv)."
}

& $pythonExe scripts\check_environment.py
if ($LASTEXITCODE -ne 0) {
    Write-Error "Environment checks failed. Fix issues above and retry."
}

& $pythonExe -m uvicorn app:app --host 127.0.0.1 --port 8000
