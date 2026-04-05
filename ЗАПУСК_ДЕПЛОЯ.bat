@echo off
cd /d "%~dp0"
echo CorelDRAW must be running with a document open.
echo Press Alt+F11 once, then close VBA, then press any key...
pause >nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0deploy.ps1"
pause
