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

trap {
    Write-Host "An error occurred:" -ForegroundColor Red
    Write-Error $_
    Read-Host "Press Enter to exit..."
    Exit 1
}

# === STAGE 1: LOCALIZATION & DOWNLOAD ===
# We check if we are already inside the "Sanctum" (Target Directory)
# If not, we assume we are running from the web or a temp folder, so we download the repo.

$CurrentDir = $PSScriptRoot
if ($CurrentDir -ne $TargetDir -and -not $SkipDownload) {
    Write-Host "~~~ PHASE 1: MATERIALIZATION ~~~" -ForegroundColor Magenta
    
    # 1. Clean/Create Target
    if (Test-Path $TargetDir) {
        Write-Host "Gathering the material in $TargetDir." -ForegroundColor Yellow
        # Optional: Remove-Item $TargetDir -Recurse -Force (If you want a fresh wipe every time)
    } else {
        New-Item -Path $TargetDir -ItemType Directory -Force | Out-Null
    }

    # 2. Download Repository
    $ZipPath = "$env:TEMP\Grimoire_Setup.zip"
    $Url = "https://github.com/$GitHubUser/$RepoName/archive/refs/heads/$Branch.zip"
    
    Write-Host "Reading knowledge from $Url..." -ForegroundColor Cyan
    try {
        Invoke-WebRequest -Uri $Url -OutFile $ZipPath
    } catch {
        Write-Error "Failed to acces the knowledge."
        Read-Host "Press Enter to exit..."
        Exit
    }
    Start-Sleep -Seconds 1

    # 3. Extract
    Write-Host "Unpacking knowledge..." -ForegroundColor Cyan
    # GitHub zips are usually nested (Repo-main). We extract to temp, then move contents.
    $TempExtract = "$env:TEMP\Extract"
    if (Test-Path $TempExtract) { Remove-Item $TempExtract -Recurse -Force }
    Expand-Archive -Path $ZipPath -DestinationPath $TempExtract -Force
    
    # Move files from subfolder to TargetDir
    $SubFolder = Get-ChildItem -Path $TempExtract -Directory | Select-Object -First 1
    Copy-Item -Path "$($SubFolder.FullName)\*" -Destination $TargetDir -Recurse -Force
    Start-Sleep -Seconds 1

    # 4. Handoff to Local Script
    Write-Host "Reading the scroll of knowledge..." -ForegroundColor Green
    $LocalScript = Join-Path $TargetDir "Incantation\Summon.ps1"
    
    if (-not (Test-Path $LocalScript)) {
        Write-Error "CRITICAL: Could not find incantation scroll at $LocalScript"
        Read-Host "Press Enter to inspect the damage..."
        Exit
    }

    Set-Location (Split-Path $LocalScript) # Move into the Incantation folder
    Start-Sleep -Seconds 2
    & PowerShell -ExecutionPolicy Bypass -File $LocalScript -SkipDownload
    
    Exit
}

# === STAGE 2: ELEVATION & SETUP ===
# If we reached here, we are running locally inside C:\Data\Grimoire.

Write-Host "~~~ PHASE 2: PREPARING THE INCANTATION ~~~" -ForegroundColor Magenta

# 1. Elevate (Admin Rights)
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Focusing the energy..." -ForegroundColor Yellow
    # Relaunch as Admin in the current directory
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -SkipDownload" -Verb RunAs
    Exit
}

# 2. Check/Install Python
if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Host "Python is absent. Conjuring it via Winget..." -ForegroundColor Cyan
    winget install -e --id Python.Python.3.12 --scope machine --accept-source-agreements --accept-package-agreements
    # Refresh env vars for this session
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
    Start-Sleep -Seconds 2
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
Start-Sleep -Seconds 1

# 4. Create Virtual Environment
$VenvPath = Join-Path $PSScriptRoot ".venv"
if (-not (Test-Path $VenvPath)) {
    Write-Host "Weaving the containment field (.venv)..." -ForegroundColor Cyan
    uv venv $VenvPath
}
Start-Sleep -Seconds 1

# 5. Install Dependencies (Rich)
Write-Host "Infusing reagents (Rich)..." -ForegroundColor Cyan
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