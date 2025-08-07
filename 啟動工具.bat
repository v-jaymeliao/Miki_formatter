@echo off
title Miki Word Document Formatter
echo.
echo ========================================
echo    Miki Word Document Formatter
echo ========================================
echo.
echo Starting GUI interface...
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python not found!
    echo Please make sure Python is installed and added to system PATH.
    echo.
    pause
    exit /b 1
)

REM Check if dependencies are installed
python -c "import docx" >nul 2>&1
if errorlevel 1 (
    echo Installing required packages...
    pip install python-docx
    if errorlevel 1 (
        echo Package installation failed! Please contact technical support.
        pause
        exit /b 1
    )
)

REM Start GUI
python gui_formatter.py

REM If program exits with error, show error message
if errorlevel 1 (
    echo.
    echo Program encountered an error during execution!
    echo Please contact technical support or check error messages.
    echo.
    pause
)
