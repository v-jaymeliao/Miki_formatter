"""
Miki Word Document Formatter GUI
圖形界面版本的Word文件格式化工具 - 簡化版
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
import threading
from io import StringIO
from main import batch_process_documents

class WordFormatterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Miki Word 文件格式化工具 v1.0")
        self.root.geometry("1000x700")
        self.root.minsize(800, 600)
        self.root.resizable(True, True)
        
        # 設置圖標和樣式
        try:
            self.root.iconbitmap(default='')
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
            "4️⃣ 處理完成的文件會自動儲存在對應的資料夾中"
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
        
        # 主內容區域 - 左右分欄佈局
        main_content_frame = tk.Frame(self.root)
        main_content_frame.pack(pady=5, padx=20, fill='both', expand=True)
        
        # 左側控制區域 - 固定寬度
        left_frame = tk.Frame(main_content_frame, width=300)
        left_frame.pack(side='left', fill='y', padx=(0, 10))
        left_frame.pack_propagate(False)  # 保持固定寬度
        
        # 右側日誌區域 - 自動擴展
        right_frame = tk.Frame(main_content_frame)
        right_frame.pack(side='right', fill='both', expand=True)
        
        # 左側內容：處理狀態和其他信息
        status_frame = tk.LabelFrame(left_frame, text="📋 處理狀態", 
                                   font=("Arial", 10, "bold"), padx=10, pady=10)
        status_frame.pack(fill='x', pady=(0, 10))
        
        # 添加一些狀態信息
        self.status_label = tk.Label(status_frame, text="等待開始...", 
                                   font=("Arial", 9), fg="#7f8c8d")
        self.status_label.pack(anchor='w')
        
        # 處理統計
        stats_frame = tk.LabelFrame(left_frame, text="📊 處理統計", 
                                  font=("Arial", 10, "bold"), padx=10, pady=10)
        stats_frame.pack(fill='x', pady=(0, 10))
        
        self.stats_text = tk.Text(stats_frame, height=6, font=("Consolas", 8),
                                bg="#f8f9fa", fg="#2c3e50", wrap=tk.WORD)
        self.stats_text.pack(fill='x')
        self.stats_text.insert(tk.END, "尚未開始處理...")
        
        # 文件處理結果列表
        result_frame = tk.LabelFrame(left_frame, text="📄 處理結果", 
                                   font=("Arial", 10, "bold"), padx=10, pady=10)
        result_frame.pack(fill='both', expand=True)
        
        # 創建結果列表框和滾動條
        result_list_frame = tk.Frame(result_frame)
        result_list_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        # 垂直滾動條
        result_scrollbar_v = tk.Scrollbar(result_list_frame, orient="vertical")
        result_scrollbar_v.pack(side="right", fill="y")
        
        # 橫向滾動條
        result_scrollbar_h = tk.Scrollbar(result_list_frame, orient="horizontal")
        result_scrollbar_h.pack(side="bottom", fill="x")
        
        # 結果列表框
        self.result_listbox = tk.Listbox(result_list_frame, 
                                       font=("Consolas", 8),
                                       bg="#f8f9fa", fg="#2c3e50",
                                       yscrollcommand=result_scrollbar_v.set,
                                       xscrollcommand=result_scrollbar_h.set,
                                       selectmode=tk.SINGLE)
        
        # 配置滾動條
        result_scrollbar_v.config(command=self.result_listbox.yview)
        result_scrollbar_h.config(command=self.result_listbox.xview)
        
        self.result_listbox.pack(side="left", fill="both", expand=True)
        
        # 添加提示
        self.result_listbox.insert(tk.END, "等待處理文件...")
        
        # 右側：處理日誌區域 - 更大的空間
        log_frame = tk.LabelFrame(right_frame, text="📊 處理日誌", 
                                font=("Arial", 10, "bold"), padx=5, pady=5)
        log_frame.pack(fill='both', expand=True)
        
        # 創建文本框和滾動條 - 優化佈局
        text_frame = tk.Frame(log_frame)
        text_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        # 垂直滾動條
        scrollbar_v = tk.Scrollbar(text_frame, orient="vertical")
        scrollbar_v.pack(side="right", fill="y")
        
        # 水平滾動條（可選）
        scrollbar_h = tk.Scrollbar(text_frame, orient="horizontal")
        scrollbar_h.pack(side="bottom", fill="x")
        
        # 文本框 - 現在有更大的空間
        self.log_text = tk.Text(text_frame, 
                              wrap=tk.NONE,  # 改為不自動換行以支持水平滾動
                              font=("Consolas", 9), 
                              bg="#f8f9fa", 
                              fg="#2c3e50",
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
        # 檢查PDF功能是否可用
        try:
            from docx2pdf import convert
            pdf_status = "✅ Word 轉 PDF 功能已啟用"
        except ImportError:
            pdf_status = "❌ PDF 轉換功能未啟用 (缺少 docx2pdf)"
            
        welcome_msg = f"""🎉 歡迎使用 Miki Word 文件格式化工具！

功能：
✅ 自動為 Word 表格添加總計行
✅ 批量處理多個文件
✅ 保持原始文件不變
{pdf_status}

快捷鍵：
• Ctrl+O: 選擇文件
• Ctrl+D: 選擇資料夾  
• F5: 開始處理
• Ctrl+L: 清空日誌
• F1: 顯示此說明

處理後的文件會儲存在對應的資料夾中。
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
        
        # 同時清空結果列表
        self.result_listbox.delete(0, tk.END)
        self.result_listbox.insert(tk.END, "等待處理文件...")
        
        # 重置統計
        self.stats_text.delete(1.0, tk.END)
        self.stats_text.insert(tk.END, "尚未開始處理...")
        
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
        self.status_label.config(text="正在處理中...", fg="#e74c3c")
        self.clear_log()
        
        # 在後台線程中處理
        thread = threading.Thread(target=self.process_files)
        thread.daemon = True
        thread.start()
        
    def update_stats(self, stats_text):
        """更新處理統計"""
        self.root.after(0, self._update_stats, stats_text)
        
    def _update_stats(self, stats_text):
        """在主線程中更新統計信息"""
        self.stats_text.delete(1.0, tk.END)
        self.stats_text.insert(tk.END, stats_text)
        self.root.update_idletasks()
        
    def add_file_result(self, filename, success, has_pdf=False, error_msg=""):
        """添加文件處理結果"""
        self.root.after(0, self._add_file_result, filename, success, has_pdf, error_msg)
        
    def _add_file_result(self, filename, success, has_pdf=False, error_msg=""):
        """在主線程中添加文件結果"""
        # 如果是第一個結果，先清空提示文字
        if self.result_listbox.size() == 1 and self.result_listbox.get(0) == "等待處理文件...":
            self.result_listbox.delete(0, tk.END)
        
        # 格式化文件名（只顯示文件名，不顯示完整路徑）
        display_name = os.path.basename(filename) if filename else "未知文件"
        
        if success:
            pdf_icon = " 📄" if has_pdf else " ❌"
            result_text = f"✅ {display_name}{pdf_icon}"
        else:
            result_text = f"❌ {display_name}"
            if error_msg:
                result_text += f" ({error_msg[:30]}...)" if len(error_msg) > 30 else f" ({error_msg})"
        
        self.result_listbox.insert(tk.END, result_text)
        
        # 自動滾動到最新項目
        self.result_listbox.see(tk.END)
        self.root.update_idletasks()
        
    def process_files(self):
        try:
            input_path = self.path_var.get().strip()
            recursive = self.recursive_var.get()
            pattern = "*.docx"  # 固定使用 docx 模式
            
            self.log_message("開始處理...")
            self.log_message(f"輸入路徑: {input_path}")
            self.log_message(f"搜尋子資料夾: {'是' if recursive else '否'}")
            self.log_message("=" * 50)
            
            # 初始化統計
            self.update_stats("開始處理...\n正在掃描文件...")
            
            # 重定向 print 輸出到 GUI
            import sys
            from io import StringIO
            
            old_stdout = sys.stdout
            sys.stdout = StringIO()
            
            # 用於統計的變量
            processed_count = 0
            failed_count = 0
            pdf_success_count = 0
            
            try:
                batch_process_documents(input_path, recursive, pattern)
                output = sys.stdout.getvalue()
                
                # 解析輸出來更新統計和結果列表
                lines = output.split('\n')
                current_processing_file = None
                
                for line in lines:
                    if line.strip():
                        self.log_message(line)
                        
                        # 檢測正在處理的文件
                        if "Processing file" in line and ".docx" in line:
                            # 提取文件名
                            if ":" in line:
                                try:
                                    # 從 "Processing file (1/3): filename.docx" 中提取文件名
                                    if "):" in line:
                                        current_processing_file = line.split("):")[1].strip()
                                    else:
                                        # 從 "Processing file: full_path" 中提取文件名
                                        current_processing_file = os.path.basename(line.split(":", 1)[1].strip())
                                except:
                                    current_processing_file = "未知文件"
                        
                        # 檢測成功處理的文件
                        elif "✓ Successfully processed:" in line and not line.startswith("   ✓") and not line.endswith(" files"):
                            if ".docx" in line or current_processing_file:
                                processed_count += 1
                                
                                # 提取文件名
                                filename_to_show = current_processing_file
                                if not filename_to_show:
                                    try:
                                        filename_to_show = os.path.basename(line.split(":", 1)[1].strip())
                                    except:
                                        filename_to_show = "未知文件"
                                
                                # 等待PDF結果，先不添加到列表
                                
                        # 檢測處理失敗的文件
                        elif "✗ Processing failed:" in line and not line.startswith("   ✗") and not line.endswith(" files"):
                            if ".docx" in line or current_processing_file:
                                failed_count += 1
                                
                                # 提取文件名和錯誤信息
                                filename_to_show = current_processing_file
                                if not filename_to_show:
                                    try:
                                        parts = line.split(":", 1)[1].split(" - Error:", 1)
                                        filename_to_show = os.path.basename(parts[0].strip())
                                    except:
                                        filename_to_show = "未知文件"
                                
                                # 提取錯誤信息
                                error_msg = ""
                                if " - Error:" in line:
                                    try:
                                        error_msg = line.split(" - Error:", 1)[1].strip()
                                    except:
                                        pass
                                
                                # 添加失敗結果到列表
                                self.add_file_result(filename_to_show, False, False, error_msg)
                                current_processing_file = None
                        
                        # 檢測PDF結果
                        elif "→ Word: ✓ | PDF:" in line:
                            if current_processing_file:
                                has_pdf = "✓ PDF generated" in line
                                if "✓ PDF generated" in line:
                                    pdf_success_count += 1
                                
                                # 添加成功結果到列表
                                self.add_file_result(current_processing_file, True, has_pdf)
                                current_processing_file = None
                                
                        # 更新實時統計
                        if processed_count > 0 or failed_count > 0:
                            stats_text = f"處理統計:\n\n"
                            stats_text += f"✅ 成功: {processed_count} 個文件\n"
                            stats_text += f"❌ 失敗: {failed_count} 個文件\n"
                            stats_text += f"📄 PDF成功: {pdf_success_count} 個文件\n\n"
                            if processed_count > 0:
                                pdf_rate = (pdf_success_count / processed_count) * 100
                                stats_text += f"PDF成功率: {pdf_rate:.1f}%"
                            self.update_stats(stats_text)
                            
            finally:
                sys.stdout = old_stdout
                
            self.log_message("\n🎉 處理完成！")
            self.log_message("處理後的文件已儲存在對應的資料夾中。")
            
            # 最終統計
            final_stats = f"最終統計:\n\n"
            final_stats += f"✅ 成功處理: {processed_count} 個文件\n"
            final_stats += f"❌ 處理失敗: {failed_count} 個文件\n"
            final_stats += f"📄 PDF成功: {pdf_success_count} 個文件\n\n"
            if processed_count > 0:
                pdf_rate = (pdf_success_count / processed_count) * 100
                final_stats += f"PDF成功率: {pdf_rate:.1f}%\n\n"
            final_stats += "✨ 處理完成！"
            self.update_stats(final_stats)
            
            # 檢查 PDF 功能狀態
            try:
                from docx2pdf import convert
                completion_msg = "文件處理完成！\n\n處理後的文件已儲存在 'success_docx' 和 'success_pdf' 資料夾中。"
            except ImportError:
                completion_msg = "文件處理完成！\n\n處理後的 Word 文件已儲存在 'success_docx' 資料夾中。\n\n注意：PDF 轉換功能未啟用，僅生成了 Word 文件。"
            
            # 顯示完成對話框
            self.root.after(0, lambda: messagebox.showinfo("完成", completion_msg))
            
        except Exception as e:
            self.log_message(f"處理過程中發生錯誤: {str(e)}")
            self.update_stats(f"錯誤:\n\n{str(e)}")
            self.root.after(0, lambda: messagebox.showerror("錯誤", f"處理失敗: {str(e)}"))
        finally:
            # 恢復 UI 狀態
            self.root.after(0, self.finish_processing)
            
    def finish_processing(self):
        self.progress.stop()
        self.process_button.config(state='normal', text="🚀 開始處理")
        self.status_label.config(text="處理完成", fg="#27ae60")
        self.processing = False

def main():
    root = tk.Tk()
    app = WordFormatterGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
