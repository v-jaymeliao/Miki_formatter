@echo off
title Miki Word 文件格式化工具
echo.
echo ========================================
echo    Miki Word 文件格式化工具
echo ========================================
echo.
echo 正在啟動圖形界面...
echo.

REM 檢查 Python 是否安裝
python --version >nul 2>&1
if errorlevel 1 (
    echo 錯誤: 找不到 Python 程式！
    echo 請確保已安裝 Python 並加入系統路徑。
    echo.
    pause
    exit /b 1
)

REM 檢查依賴是否安裝
python -c "import docx" >nul 2>&1
if errorlevel 1 (
    echo 正在安裝必要套件...
    pip install python-docx
    if errorlevel 1 (
        echo 套件安裝失敗！請聯繫技術支援。
        pause
        exit /b 1
    )
)

REM 啟動 GUI
python gui_formatter.py

REM 如果程式異常結束，顯示錯誤訊息
if errorlevel 1 (
    echo.
    echo 程式執行時發生錯誤！
    echo 請聯繫技術支援或查看錯誤訊息。
    echo.
    pause
)
