<#
.SYNOPSIS
    Sets up the environment and trigger the Incatation.py script
#>
param (
    [switch]$SkipDownload
)

$ErrorActionPreference = "Stop"

trap {
    Write-Host "An error occurred:" -ForegroundColor Red
    Write-Error $_
    Read-Host "Press Enter to exit..."
    Exit 1
}

Write-Host "~~~ PREPARING THE INCANTATION ~~~" -ForegroundColor Magenta

# 1. Elevate (Admin Rights)
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Focusing the energy..." -ForegroundColor Yellow
    # Relaunch as Admin in the current directory
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -SkipDownload" -Verb RunAs
    Exit
}

# 2. Check/Install Python
if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Host "Python is absent. Conjuring it via Winget..." -ForegroundColor Green
    winget install -e --id Python.Python.3.12 --scope machine --accept-source-agreements --accept-package-agreements
    # Refresh env vars for this session
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
    Start-Sleep -Seconds 2
}

# 3. Install/Update uv
Write-Host "Gathering the tools (uv)..." -ForegroundColor Green
if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
    powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
    # Manual Path Refresh for uv
    $UvBin = Join-Path $env:USERPROFILE ".local\bin"
    if (Test-Path $UvBin) {
        $env:Path = "$UvBin;" + $env:Path
    }
} else {
    uv self update
}
Start-Sleep -Seconds 1

# 4. Create Virtual Environment
$VenvPath = Join-Path $PSScriptRoot ".venv"
if (-not (Test-Path $VenvPath)) {
    Write-Host "Weaving the containment field (.venv)..." -ForegroundColor Green
    uv venv $VenvPath
}
Start-Sleep -Seconds 1

# 5. Install Dependencies (Rich)
Write-Host "Infusing reagents (Rich)..." -ForegroundColor Green
# Using uv pip to install directly into the venv
uv pip install rich requests --python "$VenvPath\Scripts\python.exe"
Start-Sleep -Seconds 1

# 6. Execute the Grimoire (Python)
Write-Host "Everything is ready." -ForegroundColor Green
Start-Sleep -Seconds 1
Read-Host "Press Enter to start the incantation..."
Clear-Host

# Ensure Incantation.py exists
if (-not (Test-Path "Incantation.py")) {
    Write-Error "The incantation failed"
    Read-Host "Press Enter to exit..."
    Exit
}

& "$VenvPath\Scripts\python.exe" "Incantation.py"

Write-Host "The incantation has ended."
Read-Host "Press Enter to vanish..."