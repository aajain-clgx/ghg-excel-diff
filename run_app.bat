@echo off
REM ═══════════════════════════════════════════════════════════════════════
REM  GHG Excel QA Tool — Windows Launcher
REM  Double-click this file to start the tool.
REM  Your browser will open automatically at http://localhost:8510
REM
REM  FIRST TIME? Run setup_windows.bat before using this.
REM ═══════════════════════════════════════════════════════════════════════

setlocal
set TOOL_DIR=%~dp0

REM ── Detect Python (embedded bundle first, then system) ─────────────────
set PYTHON=

if exist "%TOOL_DIR%python\python.exe" (
    set PYTHON=%TOOL_DIR%python\python.exe
    goto :launch
)

where python >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set PYTHON=python
    goto :launch
)

where python3 >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set PYTHON=python3
    goto :launch
)

echo.
echo ERROR: Python not found. Please run setup_windows.bat first.
echo.
pause
exit /b 1

:launch
echo ===================================================
echo  GHG Excel QA Tool
echo ===================================================
echo.
echo  Starting... please wait.
echo  Your browser will open automatically.
echo  URL: http://localhost:8510
echo.
echo  To stop the tool, close this window.
echo ===================================================
echo.

%PYTHON% -m streamlit run "%TOOL_DIR%app\streamlit_app.py" ^
    --server.headless true ^
    --server.port 8510 ^
    --browser.serverAddress localhost ^
    --browser.gatherUsageStats false

REM ── Handle exit ────────────────────────────────────────────────────────
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo The tool stopped with an error.
    echo.
    echo If you see "ModuleNotFoundError", run setup_windows.bat first.
    echo If the problem persists, contact your IT support desk.
    echo.
    pause
)
