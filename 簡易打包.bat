@echo off
echo =================================================
echo    Miki Word 文件格式化工具 - 簡易打包腳本
echo =================================================
echo.

echo 檢查 Python 環境...
python --version
if errorlevel 1 (
    echo 錯誤: 找不到 Python！
    echo 請先安裝 Python 3.6+ 從 https://python.org
    pause
    exit /b 1
)
echo.

echo 安裝打包工具...
pip install pyinstaller python-docx
echo.

echo 清理舊檔案...
if exist "dist" rmdir /s /q dist
if exist "build" rmdir /s /q build
echo.

echo 開始打包...
pyinstaller --onefile --windowed ^
    --name "MikiFormatter" ^
    --add-data "README.md;." ^
    --hidden-import "tkinter" ^
    --hidden-import "tkinter.ttk" ^
    --hidden-import "tkinter.filedialog" ^
    --hidden-import "tkinter.messagebox" ^
    --hidden-import "docx" ^
    --hidden-import "docx.oxml" ^
    --hidden-import "docx.oxml.ns" ^
    --hidden-import "threading" ^
    gui_formatter.py

echo.
if exist "dist\MikiFormatter.exe" (
    echo ✓ 成功！可執行檔案在: dist\MikiFormatter.exe
    echo.
    echo 檔案大小: 
    for %%I in ("dist\MikiFormatter.exe") do echo %%~zI bytes
    echo.
    echo 測試執行...
    start "Test" "dist\MikiFormatter.exe"
) else (
    echo ✗ 打包失敗！
)
echo.
pause
