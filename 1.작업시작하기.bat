@echo off
cd /d "%~dp0"
title PPT Spell Checker AI Tool

echo ===================================================
echo    PPT Spell Checker and Narration AI Tool
echo ===================================================
echo.
echo  [WARNING] Do NOT close this window while using the app!
echo.

REM --- Check if Python is installed ---
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH.
    echo Please install Python from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)

REM --- Check if Streamlit is installed ---
python -m streamlit --version >nul 2>&1
if errorlevel 1 (
    echo [SETUP] Required packages not found. Installing now...
    echo This only happens the first time. Please wait a few minutes.
    echo.
    python -m pip install --upgrade pip
    python -m pip install streamlit python-pptx pandas openai python-dotenv PyMuPDF
    if errorlevel 1 (
        echo.
        echo [ERROR] Failed to install required packages.
        echo Please check your internet connection and try again.
        echo.
        pause
        exit /b 1
    )
    echo.
    echo [SETUP] Installation complete!
    echo.
)

echo Starting server... Your browser will open automatically.
echo If it does not, open your browser and go to: http://localhost:8501
echo.
echo ===================================================
echo.

python -m streamlit run app.py

pause
