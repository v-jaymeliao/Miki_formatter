# Miki Word 文件格式化工具

這個工具可以批量處理 Word 文件，為表格添加總計行並格式化。

## 使用方法

### 1. 命令行模式

#### 基本用法：
```bash
python main.py <文件或目錄路徑>
```

#### 選項：
- `--no-recursive`: 不遞歸搜索子目錄
- `--pattern <模式>`: 指定文件過濾模式（默認: *.docx）

#### 範例：
```bash
# 處理單個文件
python main.py "report.docx"

# 處理整個目錄（包含子目錄）
python main.py "C:\Reports"

# 處理目錄但不包含子目錄
python main.py "C:\Reports" --no-recursive

# 處理特定模式的文件
python main.py "C:\Reports" --pattern "*report*.docx"
```

### 2. 交互式模式

直接運行程序，不提供參數：
```bash
python main.py
```

程序會提示你輸入文件或目錄路徑，並詢問是否遞歸搜索。

### 3. GUI 模式

運行圖形界面版本：
```bash
python gui_formatter.py
```

GUI 提供以下功能：
- 選擇單個文件或目錄
- 設置是否遞歸搜索
- 自定義文件過濾模式
- 實時查看處理日誌
- 進度指示

## 功能特點

1. **批量處理**: 可以處理單個文件或整個目錄樹
2. **智能過濾**: 自動跳過已經處理過的文件（以 "Formatted_" 開頭）
3. **錯誤處理**: 即使某個文件處理失敗，也會繼續處理其他文件
4. **詳細日誌**: 顯示處理進度和結果摘要
5. **格式保持**: 保持原始表格的格式和對齊方式

## 輸出

- 處理後的文件會以 "Formatted_" 前綴保存
- 原始文件保持不變
- 處理結果會顯示成功和失敗的文件清單

## 系統要求

- Python 3.6+
- python-docx 庫
- tkinter（GUI 版本需要，通常 Python 自帶）

## 安裝依賴

```bash
pip install python-docx
```

## 注意事項

1. 確保 Word 文件沒有被其他程序打開
2. 處理大量文件時可能需要一些時間
3. 建議先在少量文件上測試
4. 程序會自動跳過非 .docx 文件
