<#
.SYNOPSIS
    The Summoning Circle. Prepares the environment for the Grimoire.
#>
$ErrorActionPreference = "Stop"

# 1. Elevate immediately (The ritual requires high privileges)
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Requesting higher powers (Elevation)..." -ForegroundColor Yellow
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    Exit
}

Write-Host "~~~ THE SUMMONING CIRCLE ~~~" -ForegroundColor Magenta

# 2. Check/Install Python
if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Host "Python is missing. Conjuring it via Winget..." -ForegroundColor Cyan
    winget install -e --id Python.Python.3.12 --scope machine --accept-source-agreements --accept-package-agreements
    # Refresh env vars for this session
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
}

# 3. Install/Update uv (The Astral Package Manager)
Write-Host "Aligning the Astral Package Manager (uv)..." -ForegroundColor Cyan

# Install if missing
if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
    powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
    
    # --- FIX STARTS HERE ---
    # Manually add the new .local/bin to the CURRENT session's path so we can use it immediately
    $UvBin = Join-Path $env:USERPROFILE ".local\bin"
    if (Test-Path $UvBin) {
        Write-Host "      Refresh: Adding $UvBin to current session..." -ForegroundColor Gray
        $env:Path = "$UvBin;" + $env:Path
    }
    # --- FIX ENDS HERE ---
} else {
    uv self update
}

# Double check it works now
if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
    Write-Error "uv installation failed or path not updated. Please restart your terminal and try again."
    Exit 1
}

# 4. Create Virtual Environment
$VenvPath = Join-Path $PSScriptRoot ".venv"
if (-not (Test-Path $VenvPath)) {
    Write-Host "Weaving the containment field (.venv)..." -ForegroundColor Cyan
    uv venv $VenvPath
}

# 5. Install Dependencies (Rich)
Write-Host "Infusing reagents (Rich)..." -ForegroundColor Cyan

# We use uv to install packages because the venv was created without internal pip
# This is much faster and bypasses the 'No module named pip' error
uv pip install rich requests --python "$VenvPath\Scripts\python.exe"

# 6. Execute the Grimoire
Write-Host "The Circle is complete. Awakening the Grimoire..." -ForegroundColor Green
Start-Sleep -Seconds 1
Clear-Host
& "$VenvPath\Scripts\python.exe" "Incantation.py"

Write-Host "The ritual has ended. Press Enter to vanish."
Read-Host