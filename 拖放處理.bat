@echo off
title Miki Word Document Formatter - Drag & Drop Handler
echo.
echo ========================================
echo    Miki Word Document Formatter
echo ========================================
echo.

if "%~1"=="" (
    echo Usage:
    echo 1. Drag Word files or folders containing Word files onto this batch file
    echo 2. Or double-click this file and enter the path
    echo.
    set /p "input_path=Please enter the file or folder path to process: "
) else (
    set "input_path=%~1"
    echo Processing dragged path: %input_path%
)

echo.
echo Processing...
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python not found!
    pause
    exit /b 1
)

REM Check dependencies
python -c "import docx" >nul 2>&1
if errorlevel 1 (
    echo Installing required packages...
    pip install python-docx
)

REM Process files
python main.py "%input_path%"

echo.
echo Processing complete! Press any key to close window.
pause >nul
