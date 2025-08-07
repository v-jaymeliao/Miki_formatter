@echo off
title Miki Word 文件格式化工具 - 拖放處理器
echo.
echo ========================================
echo    Miki Word 文件格式化工具
echo ========================================
echo.

if "%~1"=="" (
    echo 使用方法:
    echo 1. 將 Word 文件或包含 Word 文件的資料夾拖放到此批處理文件上
    echo 2. 或者雙擊此文件然後輸入路徑
    echo.
    set /p "input_path=請輸入要處理的文件或資料夾路徑: "
) else (
    set "input_path=%~1"
    echo 處理拖放的路徑: %input_path%
)

echo.
echo 正在處理...
echo.

REM 檢查 Python
python --version >nul 2>&1
if errorlevel 1 (
    echo 錯誤: 找不到 Python！
    pause
    exit /b 1
)

REM 檢查依賴
python -c "import docx" >nul 2>&1
if errorlevel 1 (
    echo 正在安裝必要套件...
    pip install python-docx
)

REM 處理文件
python main.py "%input_path%"

echo.
echo 處理完成！按任意鍵關閉視窗。
pause >nul
