# main.py
"""
Copyright © 2025 Github:OUKUI All Rights Reserved.
"""
import tkinter as tk
from tkinter import font as tkfont

class ScannerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("现场扫码系统 - 单行显示版")
        
        # --- 窗口设置 ---
        self.root.state('zoomed')  # 最大化窗口
        
        # --- 深色主题配置 ---
        self.COLOR_BG_MAIN = "#1e1e1e"
        self.COLOR_BG_SEC = "#2d2d2d"
        self.COLOR_TEXT = "#ffffff"
        self.COLOR_TEXT_DIM = "#aaaaaa"
        self.COLOR_ACCENT = "#0078d7"
        
        self.root.configure(bg=self.COLOR_BG_MAIN)

        # --- 业务逻辑变量 ---
        self.target_text = ""
        self.ok_duration = 3
        self.ng_duration = 5
        
        # --- 字体 ---
        self.font_large = tkfont.Font(family="Microsoft YaHei", size=20, weight="bold")
        self.font_normal = tkfont.Font(family="Microsoft YaHei", size=14)
        self.font_kb = tkfont.Font(family="Arial", size=14, weight="bold")
        # 状态字体调大，适应单行显示
        self.font_status = tkfont.Font(family="Microsoft YaHei", size=80, weight="bold")

        self.create_login_screen()

    # ================= 界面构建 =================

    def clear_screen(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def create_login_screen(self):
        """登录/设置界面"""
        self.clear_screen()
        
        # 标题
        header_frame = tk.Frame(self.root, bg=self.COLOR_BG_MAIN)
        header_frame.pack(fill='x', pady=10)
        
        tk.Label(header_frame, text="系统启动设置", font=self.font_large, 
                 bg=self.COLOR_BG_MAIN, fg=self.COLOR_TEXT).pack()
        tk.Label(header_frame, text="请输入目标编号及判定显示时间", 
                 font=self.font_normal, bg=self.COLOR_BG_MAIN, fg=self.COLOR_TEXT_DIM).pack(pady=5)

        # 目标编号输入
        input_frame = tk.Frame(self.root, bg=self.COLOR_BG_SEC, bd=1, relief="solid")
        input_frame.pack(pady=10, ipady=10, fill='x', padx=100)
        
        self.login_entry = tk.Entry(input_frame, font=("Arial", 20), justify='center',
                                    bg=self.COLOR_BG_SEC, fg=self.COLOR_TEXT, insertbackground='white',
                                    relief="flat")
        self.login_entry.pack(fill='x', padx=10, pady=5)
        
        # 时间设置区域
        time_frame = tk.Frame(self.root, bg=self.COLOR_BG_MAIN)
        time_frame.pack(fill='x', padx=100, pady=5)
        
        tk.Label(time_frame, text="OK显示时长(秒):", font=self.font_normal, 
                 bg=self.COLOR_BG_MAIN, fg=self.COLOR_TEXT).pack(side='left', padx=10)
        self.ok_time_entry = tk.Entry(time_frame, font=("Arial", 14), width=5, justify='center',
                                      bg=self.COLOR_BG_SEC, fg=self.COLOR_TEXT, relief="flat")
        self.ok_time_entry.insert(0, "3")
        self.ok_time_entry.pack(side='left', padx=5, ipady=5)
        
        tk.Label(time_frame, text="NG显示时长(秒):", font=self.font_normal, 
                 bg=self.COLOR_BG_MAIN, fg=self.COLOR_TEXT).pack(side='left', padx=20)
        self.ng_time_entry = tk.Entry(time_frame, font=("Arial", 14), width=5, justify='center',
                                      bg=self.COLOR_BG_SEC, fg=self.COLOR_TEXT, relief="flat")
        self.ng_time_entry.insert(0, "5")
        self.ng_time_entry.pack(side='left', padx=5, ipady=5)

        # 虚拟键盘
        self.create_keyboard(self.login_entry, is_login=True)

    def start_app(self):
        """启动主程序"""
        val = self.login_entry.get().strip()
        if not val:
            self.login_entry.config(bg="#ffcccc")
            self.root.after(200, lambda: self.login_entry.config(bg=self.COLOR_BG_SEC))
            return
        
        self.target_text = val
        
        try:
            self.ok_duration = int(self.ok_time_entry.get())
            self.ng_duration = int(self.ng_time_entry.get())
        except ValueError:
            self.ok_duration = 2
            self.ng_duration = 4
            
        self.create_main_screen()

    def create_main_screen(self):
        """作业主界面"""
        self.clear_screen()
        
        # 顶部提示板
        self.top_label = tk.Label(self.root, text=f"目标编号：{self.target_text}", 
                                  font=("Microsoft YaHei", 28, "bold"), bg="#f1c40f", fg="black", pady=10)
        self.top_label.pack(fill='x')

        # 中间区域
        main_frame = tk.Frame(self.root, bg=self.COLOR_BG_MAIN)
        main_frame.pack(fill='both', expand=True)

        tk.Label(main_frame, text="扫码输入区", font=self.font_normal, 
                 bg=self.COLOR_BG_MAIN, fg=self.COLOR_TEXT_DIM).pack(pady=10)
        
        self.scan_entry = tk.Entry(main_frame, font=("Arial", 30), justify='center',
                                   bg=self.COLOR_BG_SEC, fg=self.COLOR_TEXT, insertbackground='white',
                                   relief="flat")
        self.scan_entry.pack(pady=10, ipady=15, fill='x', padx=200)
        self.scan_entry.bind('<Return>', self.check_scan)
        
        # --- 修改：单一状态显示框 ---
        self.status_label = tk.Label(main_frame, text="● 待机中", font=self.font_status, 
                                     bg="#7f8c8d", fg="white", pady=80)
        self.status_label.pack(fill='both', expand=True, padx=50, pady=30)

        # 底部控制
        footer_frame = tk.Frame(self.root, bg=self.COLOR_BG_SEC, height=60)
        footer_frame.pack(fill='x', side='bottom')
        footer_frame.pack_propagate(False)
        
        tk.Button(footer_frame, text="重新设定", font=self.font_normal, 
                  command=self.create_login_screen,
                  bg="#c0392b", fg="white", relief="flat", padx=20, pady=5).pack(side='right', padx=20, pady=10)

        self.scan_entry.focus_set()

    def create_keyboard(self, target_entry, is_login=False):
        """虚拟键盘"""
        kb_frame = tk.Frame(self.root, bg=self.COLOR_BG_MAIN)
        kb_frame.pack(fill='both', expand=True, padx=10, pady=10)

        keys = [
            ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0'],
            ['Q', 'W', 'E', 'R', 'T', 'Y', 'U', 'I', 'O', 'P'],
            ['A', 'S', 'D', 'F', 'G', 'H', 'J', 'K', 'L', 'DEL'],
            ['Z', 'X', 'C', 'V', 'B', 'N', 'M', '-', '_', '.'],
            ['/', '(', ')', 'START']
        ]

        for row_keys in keys:
            row_frame = tk.Frame(kb_frame, bg=self.COLOR_BG_MAIN)
            row_frame.pack(fill='x', expand=True, pady=2)
            for key in row_keys:
                btn_bg = self.COLOR_ACCENT if key == 'START' else self.COLOR_BG_SEC
                cmd = lambda k=key: self.on_kb_click(k, target_entry, is_login)
                
                btn = tk.Button(row_frame, text=key, font=self.font_kb, 
                                bg=btn_bg, fg=self.COLOR_TEXT, activebackground="#3e3e3e",
                                relief="flat", command=cmd)
                btn.pack(side='left', fill='both', expand=True, padx=1, ipady=8)

    def on_kb_click(self, key, entry_widget, is_login):
        """键盘点击逻辑"""
        if key == 'DEL':
            current = entry_widget.get()
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, current[:-1])
        elif key == 'START':
            if is_login:
                self.start_app()
        else:
            entry_widget.insert(tk.END, key)

    def check_scan(self, event=None):
        """扫码判定逻辑"""
        scanned_val = self.scan_entry.get().strip()
        
        if not scanned_val:
            return

        # 锁定输入框
        self.scan_entry.config(state='disabled')
        
        target_clean = str(self.target_text).strip()
        scanned_clean = str(scanned_val).strip()

        if scanned_clean == target_clean:
            self.set_status("OK", "#27ae60", self.ok_duration)
        else:
            self.set_status("NG", "#c0392b", self.ng_duration)

    def set_status(self, text, color, duration):
        """单一文本框显示状态与倒计时（主线程 after 实现，tkinter 安全）"""
        self.status_label.config(text=f"{text} ({duration})", bg=color)
        self._countdown_tick(text, color, duration - 1)

    def _countdown_tick(self, text, color, remaining):
        if remaining > 0:
            self.status_label.config(text=f"{text} ({remaining})", bg=color)
            self.root.after(1000, self._countdown_tick, text, color, remaining - 1)
        else:
            # 倒计时结束：先解锁再清空，否则 disabled 状态下 delete 无效
            self.scan_entry.config(state='normal')
            self.scan_entry.delete(0, tk.END)
            self.scan_entry.focus_set()
            self.status_label.config(text="● 待机中", bg="#7f8c8d")

if __name__ == "__main__":
    root = tk.Tk()
    app = ScannerApp(root)
    root.mainloop()
