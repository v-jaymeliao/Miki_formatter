@echo off
echo =================================================
echo    Miki Word Document Formatter - Build Script
echo =================================================
echo.

echo Checking Python environment...
python --version
if errorlevel 1 (
    echo Error: Python not found!
    echo Please install Python 3.6+ from https://python.org
    pause
    exit /b 1
)
echo.

echo Installing build tools...
pip install pyinstaller python-docx docx2pdf
echo.

echo Cleaning old files...
if exist "dist" rmdir /s /q dist
if exist "build" rmdir /s /q build
echo.

echo Starting build process...
pyinstaller MikiFormatter.spec

echo.
if exist "dist\MikiFormatter.exe" (
    echo ^> Success! Executable file created: dist\MikiFormatter.exe
    echo.
    echo File size: 
    for %%I in ("dist\MikiFormatter.exe") do echo %%~zI bytes
    echo.
    echo Testing execution...
    start "Test" "dist\MikiFormatter.exe"
) else (
    echo ^> Build failed!
)
echo.
pause
