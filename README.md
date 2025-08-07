# 📄 Miki Word 文件格式化工具

自動為 Word 文件中的表格添加總計行，支援批量處理。

## 🚀 快速使用（最終用戶）

### 方法一：圖形界面（推薦）
1. 雙擊 `MikiFormatter.exe`（或 `啟動工具.bat`）
2. 選擇要處理的文件或資料夾
3. 點擊「開始處理」

### 方法二：拖放處理
直接將文件或資料夾拖放到 `拖放處理.bat` 上

## 💻 開發者指南

### 環境設置
```bash
# 1. 安裝依賴
pip install -r requirements.txt

# 2. 測試運行
python gui_formatter.py
```

### 專案結構
```
miki_formatter/
├── main.py              # 核心處理邏輯
├── gui_formatter.py     # GUI 界面
├── 簡易打包.bat         # 打包腳本
└── dist/               # 打包輸出
    └── MikiFormatter.exe
```

### 關鍵修改點

#### 修改輸出檔案命名（main.py 第247行）
```python
outname = os.path.join(success_dir, f"Formatted_{name}{ext}")
# 修改 "Formatted_" 部分
```

#### 修改輸出目錄（main.py 第240行）
```python
success_dir = os.path.join(dir_path, "success")
# 修改 "success" 部分  
```

#### 修改版本號（gui_formatter.py 第8行）
```python
VERSION = "1.0.0"  # 更新版本號
```

#### 修改界面文字（gui_formatter.py）
搜尋並修改：`"📄 選擇文件"`、`"🚀 開始處理"` 等

### 修改代碼流程
1. **修改代碼** → 2. **本地測試**：`python gui_formatter.py` → 3. **重新打包**：雙擊 `簡易打包.bat` → 4. **測試執行檔**

### ⚠️ 重要提醒
- **不要修改** 表格欄位定義：`["Package", "Service", "Type", "Purchased", "Used", "Remaining", "Trend"]`
- **修改前** 先備份 `dist\MikiFormatter.exe`
- **每次修改** 重要功能後都要更新版本號

### 故障排除
| 問題 | 解決方案 |
|------|----------|
| GUI 無法啟動 | `python --version` 檢查環境 |
| 找不到模組 | `pip install -r requirements.txt` |
| 打包失敗 | `pip install pyinstaller` |
| 執行檔無法運行 | 使用啟動程式.bat |

## 🛠️ 命令行版本（開發者）

### 基本用法
```bash
# GUI 模式
python gui_formatter.py

# 命令行模式
python main.py <文件或目錄路徑>

# 處理單個文件
python main.py "report.docx"

# 處理目錄（含子目錄）
python main.py "C:\Reports"

# 不含子目錄
python main.py "C:\Reports" --no-recursive
```

### 交互式模式
```bash
python main.py
```

## 📋 功能特點

✅ **批量處理** - 單個文件或整個目錄  
✅ **智能過濾** - 跳過已處理文件（Formatted_）  
✅ **錯誤處理** - 失敗不影響其他文件  
✅ **格式保持** - 保持原始表格格式  
✅ **GUI 界面** - 友善的圖形界面  

## 📦 系統要求

- **最終用戶**：Windows 10/11（不需要 Python）
- **開發者**：Python 3.6+, python-docx, tkinter

---

**版本**: 1.0.0 | **最後更新**: 2025年8月7日
