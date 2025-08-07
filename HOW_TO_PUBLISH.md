# 🚀 如何發布專業版本

## 步驟一：準備發布檔案

### 1. 完成打包
```bash
# 使用現有的打包腳本
./簡易打包.bat

# 或手動使用 PyInstaller
pip install -r build_requirements.txt
pyinstaller MikiFormatter.spec

# 檢查輸出
# 可執行文件會在 dist/MikiFormatter.exe
# 建構日誌會在 build/ 資料夾
```

**確認檔案結構：**
```
dist/
└── MikiFormatter.exe          # 打包完成的執行文件

根目錄檔案：
├── 啟動工具.bat               # GUI 啟動器
├── 拖放處理.bat               # 拖放處理器
├── 使用說明.md               # 中文使用說明
├── HOW_TO_USE.txt            # 英文使用說明  
├── RELEASE_NOTES.md          # 版本說明
└── LICENSE                   # 授權文件
```

### 2. 準備發布資料夾
```
MikiFormatter_v1.0.0/
├── MikiFormatter.exe           # 主程式 (from dist/)
├── 啟動工具.bat                # 啟動工具（GUI介面）
├── 拖放處理.bat                # 拖放處理器（命令行）
├── 使用說明.md                 # 中文說明
├── HOW_TO_USE.txt             # 英文說明
├── RELEASE_NOTES.md           # 版本說明
└── LICENSE                    # 授權文件
```

## 步驟二：GitHub Release 發布

### 1. 推送最新代碼
```bash
git add .
git commit -m "v1.0.0 - Initial professional release"
git push origin main
```

### 2. 創建 Release
1. 前往你的 GitHub 倉庫：https://github.com/v-jaymeliao/Miki_formatter
2. 點擊右側的 "Releases"
3. 點擊 "Create a new release"
4. 填寫 Release 資訊：

**Tag version:** `v1.0.0`
**Release title:** `Miki Word Document Formatter v1.0.0`
**Description:**
```markdown
## 🎉 First Professional Release!

### ✨ Features
- 🖥️ User-friendly GUI interface  
- 🔄 Batch processing for multiple documents
- 📊 Automatic table calculations with total rows
- 💾 Safe processing (originals preserved)
- 📄 Dual output: Word (.docx) + PDF (.pdf)
- 🎨 Yellow highlighting for total rows
- 🌍 Multi-language support (English/中文)

### 📥 Download
- **MikiFormatter.exe** - Ready-to-use executable (No installation required)
- **Source code** - For developers

### 📋 System Requirements
- Windows 7/8/10/11
- Microsoft Word (required for PDF conversion)
- 100MB free space (for both Word and PDF outputs)
- Sufficient RAM for large batch processing

### 🚀 Quick Start
1. Download and extract the package
2. Double-click MikiFormatter.exe (GUI) or 啟動工具.bat  
3. Select your Word files or use drag-drop with 拖放處理.bat
4. Click "Start Processing"
5. Find results in success_docx/ and success_pdf/ folders

**Full documentation:** See README.md
```

### 3. 上傳檔案
- 複製 `dist/MikiFormatter.exe` 到發布資料夾
- 可選：重命名為 `MikiFormatter_v1.0.0.exe`
- 創建 ZIP 檔案包含完整套件：(會自動幫你完成, 記得檢查)
  - MikiFormatter.exe
  - 啟動工具.bat
  - 拖放處理.bat  
  - 使用說明.md
  - HOW_TO_USE.txt
  - RELEASE_NOTES.md
  - LICENSE
- 拖放 ZIP 檔案到 "Attach binaries" 區域

### 4. 發布
- 勾選 "Set as the latest release"
- 點擊 "Publish release"

## 步驟三：其他發布平台

### 1. Microsoft Store (進階)
- 需要開發者帳號 ($19 USD)
- 需要簽名認證
- 專業分發管道

### 2. SourceForge (免費替代)
- 免費託管
- 下載統計
- 鏡像分發

### 3. 自建網站
```html
<!-- 簡單下載頁面 -->
<div class="download-section">
  <h2>Miki Word Document Formatter</h2>
  <a href="MikiFormatter_v1.0.0.exe" class="download-btn">
    📥 Download v1.0.0 (Windows)
  </a>
  <p>No installation required • Windows 7+ • Free</p>
</div>
```

## 步驟四：推廣

### 1. 社群分享
- LinkedIn 專業網路
- Reddit (r/businesstools, r/productivity)
- Facebook 群組

### 2. 文件範例
```markdown
## 🔗 下載連結

**最新版本：v1.0.0**

### 選項一：完整套件 (推薦)
[� 下載完整套件 MikiFormatter_v1.0.0.zip](https://github.com/v-jaymeliao/Miki_formatter/releases/latest/download/MikiFormatter_v1.0.0.zip)
- 包含：執行文件 + 啟動腳本 + 完整文件
- 檔案大小：約 30MB

### 選項二：僅執行文件
[📥 下載 MikiFormatter.exe](https://github.com/v-jaymeliao/Miki_formatter/releases/latest/download/MikiFormatter.exe)  
- 僅包含主程式
- 檔案大小：約 25MB

*系統需求：Windows 7+ + Microsoft Word | 無需安裝其他軟體*
```

## 步驟五：維護更新

### 版本號管理
```python
# gui_formatter.py 或 main.py
VERSION = "1.0.1"  # 修正版本
VERSION = "1.1.0"  # 功能版本  
VERSION = "2.0.0"  # 主要版本

# 同時更新 MikiFormatter.spec 中的版本資訊
```

### 發布前檢查清單
- [ ] 測試 MikiFormatter.exe 正常運作
- [ ] 測試 啟動工具.bat 可以啟動 GUI  
- [ ] 測試 拖放處理.bat 可以處理文件
- [ ] 更新 RELEASE_NOTES.md
- [ ] 更新 使用說明.md 和 HOW_TO_USE.txt
- [ ] 檢查 LICENSE 文件正確
- [ ] 執行完整的打包流程

### 自動更新通知（進階）
```python
def check_for_updates():
    # 檢查 GitHub API 獲取最新版本
    pass
```

## 🎯 預期效果

發布後你將獲得：
- ✅ 專業的 GitHub Release 頁面
- ✅ 直接下載連結
- ✅ 版本追蹤和更新記錄
- ✅ 使用統計（下載次數）
- ✅ 用戶回饋管道（Issues）
- ✅ 搜尋引擎可見性

**範例連結：**
```
完整套件：https://github.com/v-jaymeliao/Miki_formatter/releases/latest/download/MikiFormatter_v1.0.0.zip

直接下載執行文件：https://github.com/v-jaymeliao/Miki_formatter/releases/latest/download/MikiFormatter.exe

專案首頁：https://github.com/v-jaymeliao/Miki_formatter

中文說明：https://github.com/v-jaymeliao/Miki_formatter/blob/main/使用說明.md
```
