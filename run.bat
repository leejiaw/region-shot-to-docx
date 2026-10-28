@echo off
setlocal enableextensions
cd /d "%~dp0"

set "SCRIPT=%~dp0region_shot_to_docx.pyw"

rem 1) Python launcher (py.exe / pyw.exe)
where pyw.exe >nul 2>nul
if not errorlevel 1 (
    start "" pyw.exe -3 "%SCRIPT%"
    exit /b 0
)

rem 2) pythonw on PATH
where pythonw.exe >nul 2>nul
if not errorlevel 1 (
    start "" pythonw.exe "%SCRIPT%"
    exit /b 0
)

rem 3) python on PATH (shows a console window, but works everywhere)
where python.exe >nul 2>nul
if not errorlevel 1 (
    start "" python.exe "%SCRIPT%"
    exit /b 0
)

echo.
echo   Python 3 not found.
echo   Install Python 3.9+ from https://www.python.org/downloads/
echo   then run:  pip install -r requirements.txt
echo.
pause
