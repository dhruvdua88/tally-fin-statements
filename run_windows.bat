@echo off
setlocal EnableDelayedExpansion
title Tally Financial Statements Generator

cls
echo ============================================================
echo   Tally Financial Statements Generator
echo   Schedule III Balance Sheet, P^&L, 3-Year Projections
echo ============================================================
echo.

REM --- 1. Check Python --------------------------------------------------------
python --version >nul 2>&1
if errorlevel 1 (
    echo [X]  Python 3.11 or later is not installed.
    echo.
    echo      Download it free from:
    echo      https://www.python.org/downloads/windows/
    echo.
    echo      IMPORTANT: tick "Add Python to PATH" during install.
    echo      Then double-click this file again.
    echo.
    start https://www.python.org/downloads/windows/
    pause
    exit /b 1
)

for /f "tokens=*" %%v in ('python --version 2^>^&1') do set PY_VER=%%v
echo [OK] %PY_VER% found.
echo.

REM --- 2. Create / reuse virtual environment ---------------------------------
if not exist ".venv\" (
    echo [..] Creating virtual environment (first-time, ~10 seconds)...
    python -m venv .venv
    if errorlevel 1 (
        echo [X]  Could not create virtual environment.
        pause
        exit /b 1
    )
    echo [OK] Virtual environment created.
) else (
    echo [OK] Virtual environment ready.
)
echo.

REM --- 3. Install / upgrade dependencies -------------------------------------
echo [..] Installing packages (openpyxl, PySide6)...
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip --quiet
python -m pip install --upgrade -r requirements.txt --quiet
if errorlevel 1 (
    echo [X]  Package install failed. Check your internet connection.
    pause
    exit /b 1
)
echo [OK] Packages ready.
echo.

REM --- 4. Launch -------------------------------------------------------------
echo [^>^>] Starting app - this window can be closed once it opens.
echo.
python app.py

if errorlevel 1 (
    echo.
    echo [X]  App exited with an error (see above).
    pause
)
