<#
.SYNOPSIS
    The Unified Summoning Circle.
    Handles both the initial download (Kickstart) and the environment setup.
#>
param (
    [switch]$SkipDownload
)

# --- CONFIGURATION ---
$GitHubUser = "GDelforge"
$RepoName   = "Resonance"
$Branch     = "main"
$TargetDir  = "C:\Temp\$RepoName"
# ---------------------

$ErrorActionPreference = "Stop"

# === STAGE 1: LOCALIZATION & DOWNLOAD ===
# We check if we are already inside the "Sanctum" (Target Directory)
# If not, we assume we are running from the web or a temp folder, so we download the repo.

$CurrentDir = $PSScriptRoot
if ($CurrentDir -ne $TargetDir -and -not $SkipDownload) {
    Write-Host "~~~ PHASE 1: MATERIALIZATION ~~~" -ForegroundColor Magenta
    
    # 1. Clean/Create Target
    if (Test-Path $TargetDir) {
        Write-Host "The Sanctum already exists at $TargetDir." -ForegroundColor Yellow
        # Optional: Remove-Item $TargetDir -Recurse -Force (If you want a fresh wipe every time)
    } else {
        New-Item -Path $TargetDir -ItemType Directory -Force | Out-Null
    }

    # 2. Download Repository
    $ZipPath = "$env:TEMP\Grimoire_Setup.zip"
    $Url = "https://github.com/$GitHubUser/$RepoName/archive/refs/heads/$Branch.zip"
    
    Write-Host "Summoning artifacts from $Url..." -ForegroundColor Cyan
    try {
        Invoke-WebRequest -Uri $Url -OutFile $ZipPath
    } catch {
        Write-Error "Failed to download the Grimoire. Is the repository Public?"
        Exit
    }

    # 3. Extract
    Write-Host "Unpacking artifacts..." -ForegroundColor Cyan
    # GitHub zips are usually nested (Repo-main). We extract to temp, then move contents.
    $TempExtract = "$env:TEMP\Grimoire_Extract"
    if (Test-Path $TempExtract) { Remove-Item $TempExtract -Recurse -Force }
    Expand-Archive -Path $ZipPath -DestinationPath $TempExtract -Force
    
    # Move files from subfolder to TargetDir
    $SubFolder = Get-ChildItem -Path $TempExtract -Directory | Select-Object -First 1
    Copy-Item -Path "$($SubFolder.FullName)\*" -Destination $TargetDir -Recurse -Force

    # 4. Handoff to Local Script
    Write-Host "Transferring control to the local Sanctum..." -ForegroundColor Green
    Set-Location $TargetDir
    
    # We re-run this same script, but from the C:\Data location
    # We use -ExecutionPolicy Bypass to ensure it runs
    & PowerShell -ExecutionPolicy Bypass -File "$TargetDir\Summon.ps1" -SkipDownload
    
    # Exit this temporary instance
    Exit
}

# === STAGE 2: ELEVATION & SETUP ===
# If we reached here, we are running locally inside C:\Data\Grimoire.

Write-Host "~~~ PHASE 2: THE SUMMONING CIRCLE ~~~" -ForegroundColor Magenta

# 1. Elevate (Admin Rights)
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Requesting higher powers (Elevation)..." -ForegroundColor Yellow
    # Relaunch as Admin in the current directory
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -SkipDownload" -Verb RunAs
    Exit
}

# 2. Check/Install Python
if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Host "Python is missing. Conjuring it via Winget..." -ForegroundColor Cyan
    winget install -e --id Python.Python.3.12 --scope machine --accept-source-agreements --accept-package-agreements
    # Refresh env vars for this session
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
}

# 3. Install/Update uv
Write-Host "Aligning the Astral Package Manager (uv)..." -ForegroundColor Cyan
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

# 4. Create Virtual Environment
$VenvPath = Join-Path $PSScriptRoot ".venv"
if (-not (Test-Path $VenvPath)) {
    Write-Host "Weaving the containment field (.venv)..." -ForegroundColor Cyan
    uv venv $VenvPath
}

# 5. Install Dependencies (Rich)
Write-Host "Infusing reagents (Rich)..." -ForegroundColor Cyan
# Using uv pip to install directly into the venv
uv pip install rich requests --python "$VenvPath\Scripts\python.exe"

# 6. Execute the Grimoire (Python)
Write-Host "The Circle is complete. Awakening the Grimoire..." -ForegroundColor Green
Start-Sleep -Seconds 1
Clear-Host

# Ensure Incantation.py exists
if (-not (Test-Path "Incantation.py")) {
    Write-Error "The Scroll (Incantation.py) is missing from the Sanctum!"
    Read-Host "Press Enter to exit..."
    Exit
}

& "$VenvPath\Scripts\python.exe" "Incantation.py"

Write-Host "The ritual has ended."
Read-Host "Press Enter to vanish..."