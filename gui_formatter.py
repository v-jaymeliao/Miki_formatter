import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import threading
from main import batch_process_documents

class WordFormatterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Miki Word 文件格式化工具")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # 創建主框架
        self.create_widgets()
        
        # 處理狀態
        self.processing = False
        
    def create_widgets(self):
        # 標題
        title_label = tk.Label(self.root, text="Miki Word 文件格式化工具", 
                              font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        # 輸入路徑框架
        input_frame = tk.Frame(self.root)
        input_frame.pack(pady=10, padx=20, fill='x')
        
        tk.Label(input_frame, text="選擇文件或目錄:").pack(anchor='w')
        
        path_frame = tk.Frame(input_frame)
        path_frame.pack(fill='x', pady=5)
        
        self.path_var = tk.StringVar()
        self.path_entry = tk.Entry(path_frame, textvariable=self.path_var, font=("Arial", 10))
        self.path_entry.pack(side='left', fill='x', expand=True)
        
        tk.Button(path_frame, text="選擇文件", command=self.select_file).pack(side='right', padx=(5,0))
        tk.Button(path_frame, text="選擇目錄", command=self.select_directory).pack(side='right')
        
        # 選項框架
        options_frame = tk.LabelFrame(self.root, text="處理選項", padx=10, pady=10)
        options_frame.pack(pady=10, padx=20, fill='x')
        
        self.recursive_var = tk.BooleanVar(value=True)
        tk.Checkbutton(options_frame, text="遞歸搜索子目錄", 
                      variable=self.recursive_var).pack(anchor='w')
        
        pattern_frame = tk.Frame(options_frame)
        pattern_frame.pack(fill='x', pady=5)
        tk.Label(pattern_frame, text="文件模式:").pack(side='left')
        self.pattern_var = tk.StringVar(value="*.docx")
        tk.Entry(pattern_frame, textvariable=self.pattern_var, width=15).pack(side='left', padx=(5,0))
        
        # 按鈕框架
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=20)
        
        self.process_button = tk.Button(button_frame, text="開始處理", 
                                       command=self.start_processing,
                                       bg="#4CAF50", fg="white", 
                                       font=("Arial", 12, "bold"),
                                       padx=20, pady=10)
        self.process_button.pack(side='left', padx=10)
        
        tk.Button(button_frame, text="清空", command=self.clear_log,
                 bg="#f44336", fg="white", font=("Arial", 12),
                 padx=20, pady=10).pack(side='left', padx=10)
        
        # 進度條
        self.progress = ttk.Progressbar(self.root, mode='indeterminate')
        self.progress.pack(pady=10, padx=20, fill='x')
        
        # 日誌框架
        log_frame = tk.LabelFrame(self.root, text="處理日誌", padx=10, pady=10)
        log_frame.pack(pady=10, padx=20, fill='both', expand=True)
        
        # 創建文本框和滾動條
        text_frame = tk.Frame(log_frame)
        text_frame.pack(fill='both', expand=True)
        
        self.log_text = tk.Text(text_frame, wrap=tk.WORD, font=("Consolas", 9))
        scrollbar = tk.Scrollbar(text_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
    def select_file(self):
        filename = filedialog.askopenfilename(
            title="選擇 Word 文件",
            filetypes=[("Word 文件", "*.docx"), ("所有文件", "*.*")]
        )
        if filename:
            self.path_var.set(filename)
            
    def select_directory(self):
        dirname = filedialog.askdirectory(title="選擇目錄")
        if dirname:
            self.path_var.set(dirname)
            
    def log_message(self, message):
        """線程安全的日誌輸出"""
        self.root.after(0, self._log_message, message)
        
    def _log_message(self, message):
        """在主線程中輸出日誌"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def clear_log(self):
        self.log_text.delete(1.0, tk.END)
        
    def start_processing(self):
        if self.processing:
            return
            
        input_path = self.path_var.get().strip()
        if not input_path:
            messagebox.showerror("錯誤", "請選擇文件或目錄")
            return
            
        if not os.path.exists(input_path):
            messagebox.showerror("錯誤", f"路徑不存在: {input_path}")
            return
            
        self.processing = True
        self.process_button.config(state='disabled', text="處理中...")
        self.progress.start()
        self.clear_log()
        
        # 在後台線程中處理
        thread = threading.Thread(target=self.process_files)
        thread.daemon = True
        thread.start()
        
    def process_files(self):
        try:
            input_path = self.path_var.get().strip()
            recursive = self.recursive_var.get()
            pattern = self.pattern_var.get().strip() or "*.docx"
            
            self.log_message("開始處理...")
            self.log_message(f"輸入路徑: {input_path}")
            self.log_message(f"遞歸搜索: {'是' if recursive else '否'}")
            self.log_message(f"文件模式: {pattern}")
            self.log_message("=" * 50)
            
            # 重定向 print 輸出到 GUI
            import sys
            from io import StringIO
            
            old_stdout = sys.stdout
            sys.stdout = StringIO()
            
            try:
                batch_process_documents(input_path, recursive, pattern)
                output = sys.stdout.getvalue()
                for line in output.split('\n'):
                    if line.strip():
                        self.log_message(line)
            finally:
                sys.stdout = old_stdout
                
            self.log_message("\n處理完成！")
            
        except Exception as e:
            self.log_message(f"處理過程中發生錯誤: {str(e)}")
        finally:
            # 恢復 UI 狀態
            self.root.after(0, self.finish_processing)
            
    def finish_processing(self):
        self.progress.stop()
        self.process_button.config(state='normal', text="開始處理")
        self.processing = False

def main():
    root = tk.Tk()
    app = WordFormatterGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
