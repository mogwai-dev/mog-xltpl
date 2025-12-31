@echo off
setlocal
rem Recursively strip Excel metadata under the given folder (default: this repo root)

uv run python "scripts\strip_excel_meta_all.py"
PAUSE
set rc=%ERRORLEVEL%
endlocal & exit /b %rc%
