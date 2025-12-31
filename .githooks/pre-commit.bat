@echo off
setlocal

rem Windows wrapper to run the Python pre-commit hook
set HOOK_PY=%~dp0pre-commit.py

python "%HOOK_PY%"
set rc=%ERRORLEVEL%
endlocal & exit /b %rc%
