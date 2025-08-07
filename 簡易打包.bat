@echo off
echo =================================================
echo    Miki Word Document Formatter - Enhanced Build
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

echo Installing/Updating all dependencies...
pip install --upgrade pyinstaller python-docx docx2pdf lxml
pip install --upgrade pywin32 tqdm colorama
echo.

echo Cleaning old files...
if exist "dist\MikiFormatter.exe" (
    echo Attempting to remove old executable...
    del /f "dist\MikiFormatter.exe" 2>nul
    timeout /t 2 >nul
)
if exist "dist" rmdir /s /q dist 2>nul
if exist "build" rmdir /s /q build 2>nul
echo.

echo Generating pywin32 cache...
python -c "import win32com.client; win32com.client.gencache.EnsureDispatch('Word.Application')" 2>nul
echo.

echo Starting enhanced build process...
pyinstaller --clean MikiFormatter.spec

echo.
if exist "dist\MikiFormatter.exe" (
    echo ^> Success! Enhanced executable file created: dist\MikiFormatter.exe
    echo.
    echo File size: 
    for %%I in ("dist\MikiFormatter.exe") do echo %%~zI bytes
    echo.
    echo Testing execution...
    echo Starting test with console enabled for debugging...
    start "Test MikiFormatter" "dist\MikiFormatter.exe"
) else (
    echo ^> Enhanced build failed!
    echo Checking build logs...
    if exist "build\MikiFormatter\warn-MikiFormatter.txt" (
        echo.
        echo Build warnings:
        type "build\MikiFormatter\warn-MikiFormatter.txt"
    )
)
echo.
pause
