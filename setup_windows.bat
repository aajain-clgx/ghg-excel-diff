@echo off
REM ═══════════════════════════════════════════════════════════════════════
REM  GHG Excel QA Tool — First-Time Setup (Windows)
REM  Run this ONCE before using run_app.bat for the first time.
REM  Double-click this file or run it from the command prompt.
REM ═══════════════════════════════════════════════════════════════════════

setlocal
set TOOL_DIR=%~dp0

echo ===================================================
echo  GHG Excel QA Tool — First-Time Setup
echo ===================================================
echo.

REM ── Detect Python ──────────────────────────────────────────────────────
set PYTHON=

REM Check for embedded Python bundle first (Option A — no admin needed)
if exist "%TOOL_DIR%python\python.exe" (
    set PYTHON=%TOOL_DIR%python\python.exe
    echo Found embedded Python: %TOOL_DIR%python\python.exe
    goto :install
)

REM Fall back to system Python (Option B — requires IT-installed Python 3.9+)
where python >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set PYTHON=python
    for /f "tokens=*" %%i in ('python --version 2^>^&1') do echo Found system Python: %%i
    goto :install
)

where python3 >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set PYTHON=python3
    for /f "tokens=*" %%i in ('python3 --version 2^>^&1') do echo Found system Python: %%i
    goto :install
)

echo.
echo ERROR: Python not found.
echo.
echo Option 1 (Recommended — no admin needed):
echo   Download WinPython from https://winpython.github.io/
echo   Extract it and copy the inner python-X.X.X folder to:
echo   %TOOL_DIR%python\
echo   Then run this setup script again.
echo.
echo Option 2 (Requires IT):
echo   Ask IT to install Python 3.9 or later from python.org
echo   Then run this setup script again.
echo.
pause
exit /b 1

:install
echo.
echo Installing required packages...
echo (This may take 1-2 minutes on first run)
echo.

%PYTHON% -m pip install --upgrade pip --quiet
if %ERRORLEVEL% NEQ 0 (
    echo WARNING: Could not upgrade pip, continuing anyway...
)

%PYTHON% -m pip install -r "%TOOL_DIR%requirements.txt"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: Package installation failed.
    echo.
    echo If your machine blocks internet access, ask IT to run:
    echo   pip install streamlit pandas openpyxl
    echo from a machine with internet access and copy the result.
    pause
    exit /b 1
)

echo.
echo ===================================================
echo  Setup complete!
echo  You can now double-click run_app.bat to start.
echo ===================================================
echo.
pause
