@echo off
setlocal

:: -----------------------------------------------------------------------------
:: Summoner's Sigil
:: Launches the Summon.ps1 incantation inside a Windows Terminal (Admin)
:: -----------------------------------------------------------------------------

:: 1. Determine Locations
set "WORK_DIR=%~dp0"
:: Strip trailing backslash for safe path concatenation
set "WORK_DIR=%WORK_DIR:~0,-1%"
set "PS_SCRIPT=%WORK_DIR%\Summon.ps1"

:: 2. Execute
:: We use PowerShell to invoke 'Start-Process -Verb RunAs' to gain Admin privileges.
:: Then we launch 'wt.exe' with the following arguments:
::   -p "Windows PowerShell"  -> Forces the specific profile
::   -d "%WORK_DIR%"          -> Sets the working directory to 'Incantation' (Required for dependencies)
::   powershell ...           -> The command to run inside the terminal tab
echo Summoning the Terminal...
powershell -Command "Start-Process wt.exe -ArgumentList '-p \"Windows PowerShell\" -d \"%WORK_DIR%\" powershell -NoExit -ExecutionPolicy Bypass -File \"%PS_SCRIPT%\" -SkipDownload' -Verb RunAs"

endlocal
