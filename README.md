# 📊 Miki Word Document Formatter

> **Professional Word document table formatter with GUI interface**

## ✨ Features

- 🖥️ **User-friendly GUI** - No technical knowledge required
- 📁 **Batch Processing** - Handle multiple files or entire folders
- 🎯 **Smart Table Detection** - Automatically finds and formats target tables
- 💾 **Safe Processing** - Original files remain untouched
- 🚀 **One-Click Operation** - Just select files and click process

## 🎯 What It Does

Automatically adds **Total** rows to Word tables containing these columns:
- Package, Service, Type, Purchased, Used, Remaining, Trend

**Output formats:**
- **Word files** (.docx) - Saved in `success_docx` folder
- **PDF files** (.pdf) - Saved in `success_pdf` folder
- Total rows highlighted with **yellow background** for better visibility

## 📥 Download

**[Download MikiFormatter.exe](https://github.com/v-jaymeliao/Miki_formatter/releases/latest)**

## 🚀 Quick Start

### Method 1: GUI Interface (Recommended)
1. Double-click `MikiFormatter.exe` to launch the APP with GUI
2. Click "📄 Select Single Word File" or "📁 Select Folder"
3. Choose your files or folders
4. Click "🚀 Start Processing"
5. Find results in the "success_docx" folder (Word files) and "success_pdf" folder (PDF files)

### Method 2: Drag & Drop (Console App without GUI)
1. Drag Word files or folders directly onto `拖放處理.bat`
2. Processing starts automatically
3. Wait for completion

## 💻 System Requirements

- Windows 7/8/10/11
- Microsoft Word (required for PDF conversion)
- Sufficient disk space (stores both Word and PDF formats)

## 📋 Usage Examples

### Single File
```
Original: report.docx
Results: 
├── success_docx/report.docx  (Word format)
└── success_pdf/report.pdf    (PDF format)
```

### Batch Processing
```
Input Folder: C:\Reports\
Output: 
├── C:\Reports\success_docx\
│   ├── report1.docx
│   ├── report2.docx
│   └── report3.docx
└── C:\Reports\success_pdf\
    ├── report1.pdf
    ├── report2.pdf
    └── report3.pdf
```

## ⚠️ Important Notes

- **Original files remain untouched** - All processing creates new files
- **Both formats generated** - Word (.docx) and PDF (.pdf) files are created
- **Total rows highlighted** - Yellow background makes totals stand out
- **Ensure Word files are closed** before processing
- **Large batches** - Tool processes in batches of 10 files to prevent memory issues
- **File format** - Only .docx files are supported

## 🆘 Support

- 📖 [User Guide (English)](HOW_TO_USE.txt)
- 📖 [使用說明 (中文)](使用說明.md) - Detailed Chinese guide
- 🐛 [Report Issues](https://github.com/v-jaymeliao/Miki_formatter/issues)
- 📧 Contact: Technical Support Team

## 📄 License

MIT License - Free to use, modify, and distribute.

---

⭐ **Found this helpful? Please star the repository!**