@echo off
set "SOURCE=%~dp0"
set "DEST=G:\My Drive\Data\Resonance"

echo Deploying from "%SOURCE%" to "%DEST%"...

:: Copy all content recursively (/E), excluding git/venv folders (/XD) and this script (/XF)
robocopy "%SOURCE%." "%DEST%" /E /XD .* __pycache__ /XF deploy.bat .*

echo Done.
pause
