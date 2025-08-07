import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import threading
from main import batch_process_documents

class WordFormatterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Miki Word 文件格式化工具")
        self.root.geometry("800x700")  # 更合適的初始大小
        self.root.minsize(600, 500)    # 設置最小尺寸
        self.root.resizable(True, True)
        
        # 設置圖標和樣式
        try:
            self.root.iconbitmap(default='')  # 可以添加圖標
        except:
            pass
        
        # 創建主框架
        self.create_widgets()
        
        # 綁定快捷鍵
        self.root.bind('<Control-o>', lambda e: self.select_file())
        self.root.bind('<Control-d>', lambda e: self.select_directory())
        self.root.bind('<F5>', lambda e: self.start_processing())
        self.root.bind('<Control-l>', lambda e: self.clear_log())
        self.root.bind('<F1>', lambda e: self.show_welcome_message())
        
        # 處理狀態
        self.processing = False
        
        # 啟動時顯示使用說明
        self.show_welcome_message()
        
    def create_widgets(self):
        # 歡迎標題
        welcome_frame = tk.Frame(self.root, bg="#f0f0f0", relief="ridge", bd=2)
        welcome_frame.pack(pady=10, padx=20, fill='x')
        
        title_label = tk.Label(welcome_frame, text="📄 Miki Word 文件格式化工具", 
                              font=("Arial", 18, "bold"), bg="#f0f0f0", fg="#2c3e50")
        title_label.pack(pady=10)
        
        subtitle_label = tk.Label(welcome_frame, text="自動為 Word 文件中的表格添加總計行", 
                                 font=("Arial", 11), bg="#f0f0f0", fg="#34495e")
        subtitle_label.pack(pady=(0, 10))
        
        # 簡單說明 - 可折疊的框架
        self.instructions_frame = tk.LabelFrame(self.root, text="📋 使用說明 (點擊展開/收起)", 
                                         font=("Arial", 10, "bold"), padx=15, pady=5)
        self.instructions_frame.pack(pady=(5, 10), padx=20, fill='x')
        
        # 讓標題可以點擊來切換顯示/隱藏
        self.instructions_visible = False
        self.instructions_content = tk.Frame(self.instructions_frame)
        
        # 點擊事件
        self.instructions_frame.bind("<Button-1>", self.toggle_instructions)
        
        instructions = [
            "1️⃣ 點擊下方按鈕選擇要處理的 Word 文件或整個資料夾",
            "2️⃣ 如果選擇資料夾，可以選擇是否搜尋子資料夾", 
            "3️⃣ 點擊「開始處理」按鈕",
            "4️⃣ 處理完成的文件會自動儲存在原位置的 'success' 資料夾中"
        ]
        
        for instruction in instructions:
            label = tk.Label(self.instructions_content, text=instruction, 
                           font=("Arial", 9), anchor='w', justify='left')
            label.pack(anchor='w', pady=1)
        
        # 文件/目錄選擇區域 - 緊湊佈局
        selection_frame = tk.LabelFrame(self.root, text="📂 選擇要處理的文件或資料夾", 
                                      font=("Arial", 10, "bold"), padx=15, pady=8)
        selection_frame.pack(pady=5, padx=20, fill='x')
        
        # 大按鈕區域
        big_buttons_frame = tk.Frame(selection_frame)
        big_buttons_frame.pack(pady=5)
        
        # 選擇單個文件按鈕 - 縮小尺寸
        file_button = tk.Button(big_buttons_frame, text="📄 選擇文件", 
                               command=self.select_file,
                               bg="#3498db", fg="white", font=("Arial", 10, "bold"),
                               padx=15, pady=10, relief="raised", bd=2)
        file_button.pack(side='left', padx=5)
        
        # 選擇資料夾按鈕 - 縮小尺寸
        folder_button = tk.Button(big_buttons_frame, text="📁 選擇資料夾", 
                                 command=self.select_directory,
                                 bg="#2ecc71", fg="white", font=("Arial", 10, "bold"),
                                 padx=15, pady=10, relief="raised", bd=2)
        folder_button.pack(side='left', padx=5)
        
        # 顯示選中的路徑 - 更緊湊
        path_display_frame = tk.Frame(selection_frame)
        path_display_frame.pack(fill='x', pady=5)
        
        tk.Label(path_display_frame, text="📍 選中:", font=("Arial", 9, "bold")).pack(anchor='w')
        
        self.path_var = tk.StringVar(value="尚未選擇文件或資料夾...")
        path_label = tk.Label(path_display_frame, textvariable=self.path_var, 
                             font=("Arial", 8), fg="#7f8c8d", wraplength=500, 
                             justify='left', relief="sunken", bd=1, padx=8, pady=3)
        path_label.pack(fill='x', pady=2)
        
        # 選項區域 - 緊湊佈局
        options_frame = tk.LabelFrame(self.root, text="⚙️ 處理選項", 
                                    font=("Arial", 10, "bold"), padx=15, pady=5)
        options_frame.pack(pady=5, padx=20, fill='x')
        
        self.recursive_var = tk.BooleanVar(value=True)
        recursive_cb = tk.Checkbutton(options_frame, text="🔄 搜尋子資料夾", 
                                     variable=self.recursive_var, font=("Arial", 9))
        recursive_cb.pack(anchor='w', pady=3)
        
        # 處理按鈕區域 - 緊湊佈局
        action_frame = tk.Frame(self.root, bg="#ecf0f1", relief="ridge", bd=1)
        action_frame.pack(pady=5, padx=20, fill='x')
        
        button_container = tk.Frame(action_frame, bg="#ecf0f1")
        button_container.pack(pady=8)
        
        self.process_button = tk.Button(button_container, text="🚀 開始處理", 
                                       command=self.start_processing,
                                       bg="#e74c3c", fg="white", 
                                       font=("Arial", 12, "bold"),
                                       padx=25, pady=8, relief="raised", bd=3)
        self.process_button.pack(side='left', padx=5)
        
        clear_button = tk.Button(button_container, text="🗑️ 清空", 
                               command=self.clear_log,
                               bg="#95a5a6", fg="white", font=("Arial", 9, "bold"),
                               padx=15, pady=8, relief="raised", bd=2)
        clear_button.pack(side='left', padx=5)
        
        # 進度條 - 緊湊佈局
        progress_frame = tk.Frame(self.root)
        progress_frame.pack(pady=2, padx=20, fill='x')
        
        tk.Label(progress_frame, text="處理進度:", font=("Arial", 9)).pack(anchor='w')
        self.progress = ttk.Progressbar(progress_frame, mode='indeterminate', style="TProgressbar")
        self.progress.pack(fill='x', pady=2)
        
        # 處理日誌區域 - 給予更多空間
        log_frame = tk.LabelFrame(self.root, text="📊 處理日誌", 
                                font=("Arial", 10, "bold"), padx=5, pady=5)
        log_frame.pack(pady=5, padx=20, fill='both', expand=True)
        
        # 創建文本框和滾動條 - 優化佈局
        text_frame = tk.Frame(log_frame)
        text_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        # 垂直滾動條
        scrollbar_v = tk.Scrollbar(text_frame, orient="vertical")
        scrollbar_v.pack(side="right", fill="y")
        
        # 水平滾動條（可選）
        scrollbar_h = tk.Scrollbar(text_frame, orient="horizontal")
        scrollbar_h.pack(side="bottom", fill="x")
        
        # 文本框 - 設置最小高度確保可見性
        self.log_text = tk.Text(text_frame, 
                              wrap=tk.NONE,  # 改為不自動換行以支持水平滾動
                              font=("Consolas", 9), 
                              bg="#f8f9fa", 
                              fg="#2c3e50",
                              height=10,  # 設置最小高度
                              yscrollcommand=scrollbar_v.set,
                              xscrollcommand=scrollbar_h.set)
        
        # 配置滾動條
        scrollbar_v.config(command=self.log_text.yview)
        scrollbar_h.config(command=self.log_text.xview)
        
        self.log_text.pack(side="left", fill="both", expand=True)
        
        # 預設顯示歡迎訊息
        self.log_text.insert(tk.END, "歡迎使用 Miki Word 文件格式化工具！\n")
        self.log_text.insert(tk.END, "請選擇要處理的文件或資料夾，然後點擊「開始處理」。\n")
        self.log_text.insert(tk.END, "=" * 50 + "\n\n")
        
    def toggle_instructions(self, event=None):
        """切換使用說明的顯示/隱藏"""
        if self.instructions_visible:
            self.instructions_content.pack_forget()
            self.instructions_frame.config(text="📋 使用說明 (點擊展開)")
            self.instructions_visible = False
        else:
            self.instructions_content.pack(fill='x', pady=5)
            self.instructions_frame.config(text="📋 使用說明 (點擊收起)")
            self.instructions_visible = True
            
    def select_file(self):
        filename = filedialog.askopenfilename(
            title="選擇 Word 文件",
            filetypes=[("Word 文件", "*.docx"), ("所有文件", "*.*")]
        )
        if filename:
            self.path_var.set(filename)
            
    def select_directory(self):
        dirname = filedialog.askdirectory(title="選擇資料夾")
        if dirname:
            self.path_var.set(dirname)
    
    def show_welcome_message(self):
        """顯示歡迎訊息"""
        welcome_msg = """
🎉 歡迎使用 Miki Word 文件格式化工具！

功能：
✅ 自動為 Word 表格添加總計行
✅ 批量處理多個文件
✅ 保持原始文件不變

快捷鍵：
• Ctrl+O: 選擇文件
• Ctrl+D: 選擇資料夾  
• F5: 開始處理
• Ctrl+L: 清空日誌
• F1: 顯示此說明

處理後的文件會儲存在 'success' 資料夾中。
        """
        messagebox.showinfo("使用說明", welcome_msg)
            
    def log_message(self, message):
        """線程安全的日誌輸出"""
        self.root.after(0, self._log_message, message)
        
    def _log_message(self, message):
        """在主線程中輸出日誌 - 改進的滾動控制"""
        self.log_text.insert(tk.END, message + "\n")
        
        # 自動滾動到底部
        self.log_text.see(tk.END)
        
        # 限制日誌行數以避免記憶體問題
        lines = self.log_text.get(1.0, tk.END).count('\n')
        if lines > 1000:  # 保留最後1000行
            self.log_text.delete(1.0, f"{lines-1000}.0")
        
        self.root.update_idletasks()
        
    def clear_log(self):
        """清空日誌並顯示歡迎訊息"""
        self.log_text.delete(1.0, tk.END)
        self.log_text.insert(tk.END, "歡迎使用 Miki Word 文件格式化工具！\n")
        self.log_text.insert(tk.END, "請選擇要處理的文件或資料夾，然後點擊「開始處理」。\n")
        self.log_text.insert(tk.END, "=" * 50 + "\n\n")
        
    def start_processing(self):
        if self.processing:
            return
            
        input_path = self.path_var.get().strip()
        if not input_path or input_path == "尚未選擇文件或資料夾...":
            messagebox.showerror("錯誤", "請先選擇要處理的文件或資料夾")
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
            pattern = "*.docx"  # 固定使用 docx 模式
            
            self.log_message("開始處理...")
            self.log_message(f"輸入路徑: {input_path}")
            self.log_message(f"搜尋子資料夾: {'是' if recursive else '否'}")
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
                
            self.log_message("\n🎉 處理完成！")
            self.log_message("處理後的文件已儲存在各自的 'success' 資料夾中。")
            
            # 顯示完成對話框
            self.root.after(0, lambda: messagebox.showinfo("完成", "文件處理完成！\n\n處理後的文件已儲存在 'success' 資料夾中。"))
            
        except Exception as e:
            self.log_message(f"處理過程中發生錯誤: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("錯誤", f"處理失敗: {str(e)}"))
        finally:
            # 恢復 UI 狀態
            self.root.after(0, self.finish_processing)
            
    def finish_processing(self):
        self.progress.stop()
        self.process_button.config(state='normal', text="🚀 開始處理")
        self.processing = False

def main():
    root = tk.Tk()
    app = WordFormatterGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
