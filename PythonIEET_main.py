import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import subprocess
import threading
import os
import sys

# 定義按鈕與對應的 Python 檔案
SCRIPTS = [
    ("1. 匯入學生成績 (STscoreRead)", "STscoreRead.py"),
    ("2. 匯入大學部問卷 (LeavDepUdataRead)", "LeavDepUdataRead.py"),
    ("3. 匯入研究所問卷 (LeavDepGdataRead)", "LeavDepGdataRead.py"),
    ("---------------------------------", None),
    ("4. 執行成績分析 (STscoreAnalyze)", "STscoreAnalyze.py"),
    ("5. 執行大學部問卷分析 (LDU_DataAnalyze)", "LDU_DataAnalyze.py"),
    ("6. 執行研究所問卷分析 (LDG_DataAnalyze)", "LDG_DataAnalyze.py"),
    ("---------------------------------", None),
    ("7. 匯出成績平均報表 (STscoreAVG_export)", "STscoreAVG_export.py"),
    ("8. 匯出成績分布報表 (STscoreDistribution_export)", "STscoreDistribution_export.py"),
    ("9. 匯出問卷分析報表 (LD_AnalyzeExport)", "LD_AnalyzeExport.py"),
]

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("IEET 資料分析控制台")
        self.root.geometry("700x750")

        self.font_style = ("Microsoft JhengHei", 10)
        tk.Label(root, text="IEET 資料處理流程", font=("Microsoft JhengHei", 16, "bold")).pack(pady=10)

        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=5, fill=tk.BOTH, expand=True)

        self.buttons = []
        for label_text, script_name in SCRIPTS:
            if script_name is None:
                ttk.Separator(btn_frame, orient='horizontal').pack(fill='x', pady=5, padx=20)
            else:
                btn = tk.Button(
                    btn_frame, 
                    text=label_text, 
                    font=self.font_style, 
                    bg="#f0f0f0",
                    command=lambda s=script_name: self.run_script_thread(s)
                )
                btn.pack(fill='x', padx=50, pady=2)
                self.buttons.append(btn)

        tk.Label(root, text="執行紀錄：", font=("Microsoft JhengHei", 10, "bold")).pack(anchor="w", padx=10, pady=(10, 0))
        self.log_area = scrolledtext.ScrolledText(root, height=12, font=("Consolas", 9))
        self.log_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    def log(self, message):
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)

    def run_script_thread(self, script_name):
        if not os.path.exists(script_name):
            messagebox.showerror("錯誤", f"找不到檔案：{script_name}\n請確認該檔案是否存在於同一目錄下。")
            return

        for btn in self.buttons:
            btn.config(state=tk.DISABLED)
        
        thread = threading.Thread(target=self.run_script, args=(script_name,))
        thread.start()

    def run_script(self, script_name):
        self.log(f"=== 開始執行 {script_name} ===")
        try:
            python_exe = sys.executable
            
            # 【關鍵修改】這裡強制使用 cp950 (Windows 繁體中文) 解碼
            process = subprocess.Popen(
                [python_exe, script_name],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding='cp950',  # 這裡是用來解決 0xb6 錯誤的關鍵
                errors='replace',  # 如果還有亂碼，用 ? 取代，不要當機
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
            )

            # 即時讀取輸出
            for line in process.stdout:
                self.log(line.strip())
            
            stderr = process.communicate()[1]
            if stderr:
                self.log("【錯誤訊息】:")
                self.log(stderr)

            if process.returncode == 0:
                self.log(f"=== {script_name} 執行完畢 (成功) ===\n")
            else:
                self.log(f"=== {script_name} 執行結束 (異常代碼: {process.returncode}) ===\n")

        except Exception as e:
            self.log(f"執行發生例外錯誤: {e}")
        finally:
            for btn in self.buttons:
                btn.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()