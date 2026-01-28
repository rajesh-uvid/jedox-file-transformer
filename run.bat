@echo off
setlocal enabledelayedexpansion
title Jedox Transformer - Global Environment

echo =========================================
echo   Checking Global Python Environment...
echo =========================================

:: 1. Verify Python is in the System PATH
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found. Please install Python and add it to PATH.
    pause
    exit /b
)

:: 2. Define and install missing global libraries
set "LIBS=streamlit pandas openpyxl"
for %%L in (%LIBS%) do (
    echo Checking for %%L...
    python -c "import %%L" >nul 2>&1
    if !errorlevel! neq 0 (
        echo [INSTALL] %%L not found. Installing globally...
        python -m pip install %%L
    )
)

echo.
echo =========================================
echo   Launching Jedox Transformer...
echo =========================================

:: 3. Run Streamlit with production-ready flags
:: --server.headless true: Prevents unwanted popups on the host machine
:: --client.showErrorDetails false: Keeps the UI clean if a file error occurs
python -m streamlit run app.py --server.headless true --client.showErrorDetails false

pause