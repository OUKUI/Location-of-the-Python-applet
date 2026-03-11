import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from scipy.stats import norm
import ctypes
import re
import os
import tempfile
from datetime import datetime
import traceback

# 尝试导入 reportlab
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm, inch
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# 尝试导入 pandas 和 openpyxl
try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False

# ==========================================
# 1. 高分屏适配 (HiDPI)
# ==========================================
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    ScaleFactor = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100
except:
    ScaleFactor = 1.0

# ==========================================
# 2. 全局深色主题
# ==========================================
plt.style.use('dark_background')
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['figure.dpi'] = 100 * ScaleFactor

THEME = {
    'bg': '#1e1e1e', 'panel': '#252526', 'fg': '#cccccc',
    'accent': '#007acc', 'input_bg': '#3c3c3c', 'success': '#4ec9b0',
    'danger': '#f44747', 'border': '#444444', 'warning': '#ffcc00'
}

# ==========================================
# 3. 核心计算
# ==========================================
class CpkCalculator:
    @staticmethod
    def calculate(data, usl, lsl):
        if data is None or len(data) < 2: 
            return {"Error": "数据不足 (N < 2)"}
        
        n = len(data)
        mu = np.mean(data)
        sigma = np.std(data, ddof=1)
        
        if sigma <= 1e-9:
            return {"Error": "标准差为 0，无法计算"}

        cp = None; cpk = None; ppm = 0; cpu = None; cpl = None
        
        if usl is not None and lsl is not None:
            if lsl >= usl: return {"Error": "LSL 必须小于 USL"}
            cpu = (usl - mu) / (3 * sigma)
            cpl = (mu - lsl) / (3 * sigma)
            cpk = min(cpu, cpl)
            cp = (usl - lsl) / (6 * sigma)
            p_upper = 1 - norm.cdf(usl, mu, sigma)
            p_lower = norm.cdf(lsl, mu, sigma)
            ppm = (p_upper + p_lower) * 1_000_000
        elif usl is not None:
            cpu = (usl - mu) / (3 * sigma); cpk = cpu; cp = None
            ppm = (1 - norm.cdf(usl, mu, sigma)) * 1_000_000
        elif lsl is not None:
            cpl = (mu - lsl) / (3 * sigma); cpk = cpl; cp = None
            ppm = norm.cdf(lsl, mu, sigma) * 1_000_000
        else:
            return {"Error": "请至少输入一个规格限"}

        cpk_level = ""
        if cpk is not None:
            if cpk >= 1.67: cpk_level = "优秀"
            elif cpk >= 1.33: cpk_level = "良好"
            elif cpk >= 1.0: cpk_level = "一般"
            elif cpk >= 0.67: cpk_level = "较差"
            else: cpk_level = "很差"

        return {
            "Count": n, "Mean": mu, "StdDev": sigma,
            "USL": usl, "LSL": lsl, "Cp": cp, "Cpk": cpk, "PPM": ppm,
            "CPU": cpu, "CPL": cpl, "CPK_LEVEL": cpk_level
        }

    @staticmethod
    def simulate(target_cpk, target_mean, usl, lsl, count, decimals):
        if target_cpk <= 0: return np.array([])
        sigma = 0
        if usl is not None and lsl is not None:
            sigma = min(usl - target_mean, target_mean - lsl) / (3 * target_cpk)
        elif usl is not None:
            sigma = abs(usl - target_mean) / (3 * target_cpk)
        elif lsl is not None:
            sigma = abs(target_mean - lsl) / (3 * target_cpk)
        else: return None
        return np.round(np.random.normal(target_mean, sigma, count), decimals)

# ==========================================
# 4. 界面组件
# ==========================================
class DarkMessageBox(tk.Toplevel):
    def __init__(self, parent, title, message, is_error=True):
        super().__init__(parent)
        self.configure(bg=THEME['panel'])
        self.title(title)
        base_width, base_height = 500, 200
        msg_lines = message.count('\n') + 1
        estimated_lines = max(msg_lines, len(message) // 50)
        h = min(max(base_height, 150 + estimated_lines * 25), 600)
        w = min(max(base_width, min(len(message) * 8, 700)), 800)
        
        x = parent.winfo_x() + (parent.winfo_width()//2) - (w//2)
        y = parent.winfo_y() + (parent.winfo_height()//2) - (h//2)
        self.geometry(f"{int(w)}x{int(h)}+{int(x)}+{int(y)}")
        self.transient(parent); self.grab_set()

        color = THEME['danger'] if is_error else THEME['success']
        symbol = "❌ 错误" if is_error else "✅ 提示"
        tk.Label(self, text=symbol, font=("Segoe UI", 14, "bold"), bg=THEME['panel'], fg=color).pack(pady=(20, 10))

        msg_frame = tk.Frame(self, bg=THEME['panel'])
        msg_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        scrollbar = tk.Scrollbar(msg_frame, bg=THEME['input_bg'])
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget = tk.Text(msg_frame, font=("Microsoft YaHei", 10), bg=THEME['input_bg'], fg='#ddd',
                             wrap=tk.WORD, height=max(3, min(estimated_lines, 15)), relief=tk.FLAT, padx=10, pady=10, yscrollcommand=scrollbar.set)
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=text_widget.yview)
        text_widget.insert('1.0', message)
        text_widget.config(state=tk.DISABLED)

        btn = tk.Button(self, text="确定", bg=THEME['input_bg'], fg='white', relief=tk.FLAT, command=self.destroy, width=15, font=("Microsoft YaHei", 10))
        btn.pack(pady=15)
        self.bind('<Return>', lambda e: self.destroy())
        self.bind('<Escape>', lambda e: self.destroy())
        self.focus_force()
        self.wait_window()

class CpkApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CPK Tool Pro - Merge to Single PDF")
        self.root.geometry(f"{int(1700)}x{int(850)}") 
        self.root.configure(bg=THEME['bg'])
        
        self.current_data = None
        self.current_stats = None
        self.current_usl = None
        self.current_lsl = None
        self.project_name = ""
        
        self.excel_projects = [] 
        self.current_excel_index = -1
        
        self.setup_ui()

    def setup_ui(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TNotebook", background=THEME['bg'], borderwidth=0)
        style.configure("TNotebook.Tab", background=THEME['panel'], foreground=THEME['fg'], padding=[25, 12], font=('Microsoft YaHei', 11))
        style.map("TNotebook.Tab", background=[("selected", THEME['accent'])], foreground=[("selected", 'white')])
        
        style.configure("Treeview", 
                        background=THEME['input_bg'], 
                        foreground=THEME['fg'], 
                        fieldbackground=THEME['input_bg'], 
                        rowheight=28,
                        font=('Microsoft YaHei', 9))
        style.map("Treeview", 
                  background=[('selected', THEME['accent'])], 
                  foreground=[('selected', 'white')])
        
        style.configure("Treeview.Heading", 
                        background=THEME['panel'], 
                        foreground='#aaaaaa', 
                        font=('Microsoft YaHei', 9, 'bold'))
        style.map("Treeview.Heading",
                  background=[('active', THEME['border'])])

        main = tk.Frame(self.root, bg=THEME['bg'])
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        left = tk.Frame(main, bg=THEME['panel'], width=460)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 5))
        left.pack_propagate(False)

        right = tk.Frame(main, bg=THEME['panel'], width=350)
        right.pack(side=tk.RIGHT, fill=tk.Y, padx=(5, 0))
        right.pack_propagate(False)
        
        center = tk.Frame(main, bg=THEME['bg'])
        center.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.init_left_panel(left)
        self.init_stats_panel(right)
        self.init_chart_panel(center)
        self.add_about_link()

    def init_left_panel(self, parent):
        tk.Label(parent, text="⚙️ 控制台", bg=THEME['panel'], fg='white', font=("Segoe UI", 16, "bold")).pack(pady=(15, 10))
        
        proj_frame = tk.Frame(parent, bg=THEME['panel'])
        proj_frame.pack(fill=tk.X, padx=15, pady=(0, 10))
        tk.Label(proj_frame, text="📁 项目名称 (手动/模拟):", bg=THEME['panel'], fg=THEME['warning'], font=('Microsoft YaHei', 9, 'bold')).pack(anchor='w')
        self.inp_project = tk.Entry(proj_frame, bg=THEME['input_bg'], fg='white', insertbackground='white', 
                                    relief=tk.FLAT, font=('Microsoft YaHei', 10))
        self.inp_project.pack(fill=tk.X, ipady=5, pady=(5,0))
        self.inp_project.insert(0, "未命名项目")
        
        nb = ttk.Notebook(parent)
        nb.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
        
        t1 = ttk.Frame(nb); nb.add(t1, text=' 数据分析 ')
        t2 = ttk.Frame(nb); nb.add(t2, text=' 模拟生成 ')
        t3 = ttk.Frame(nb); nb.add(t3, text=' Excel 导入 ') 
        
        self.setup_tab1(t1)
        self.setup_tab2(t2)
        self.setup_tab3(t3) 
        
        export_frame = tk.Frame(parent, bg=THEME['panel'])
        export_frame.pack(fill=tk.X, padx=15, pady=(10, 20))
        
        self.btn_export = tk.Button(
            export_frame, text="📄 导出当前报告 (PDF)", bg=THEME['success'], fg='#1e1e1e', 
            relief=tk.FLAT, font=('Microsoft YaHei', 11, 'bold'), command=self.export_report, cursor="hand2"
        )
        self.btn_export.pack(fill=tk.X, ipady=8)
        
        # 按钮文字修改，明确提示是“合并”
        self.btn_batch_export = tk.Button(
            export_frame, text="📦 合并导出所有为一份PDF", bg=THEME['accent'], fg='white', 
            relief=tk.FLAT, font=('Microsoft YaHei', 11, 'bold'), command=self.export_merged_report, cursor="hand2"
        )

        if not REPORTLAB_AVAILABLE:
            self.btn_export.config(state=tk.DISABLED, text="❌ 缺少 reportlab")
            self.btn_batch_export.config(state=tk.DISABLED, text="❌ 缺少 reportlab")
            tk.Label(parent, text="(需安装 reportlab)", bg=THEME['panel'], fg='#666', font=("Microsoft YaHei", 8)).pack(pady=(0, 10))
        
        if not PANDAS_AVAILABLE:
            tk.Label(parent, text="⚠️ 缺少 pandas/openpyxl\nExcel 导入功能不可用", bg=THEME['panel'], fg=THEME['danger'], font=("Microsoft YaHei", 9)).pack(pady=(10, 10))

        nb.bind("<<NotebookTabChanged>>", self.on_tab_changed)
        self.main_notebook = nb

    def on_tab_changed(self, event):
        selected_tab = event.widget.tab('current')['text']
        if "Excel" in selected_tab:
            self.btn_batch_export.pack(fill=tk.X, ipady=8, pady=(5, 0))
            if self.excel_projects and self.current_excel_index != -1:
                self.btn_export.config(state=tk.NORMAL)
            else:
                self.btn_export.config(state=tk.DISABLED)
        else:
            self.btn_batch_export.pack_forget()
            if self.current_stats:
                self.btn_export.config(state=tk.NORMAL)
            else:
                self.btn_export.config(state=tk.DISABLED)

    def init_stats_panel(self, parent):
        self.lbl_proj_display = tk.Label(parent, text="项目：未命名", bg=THEME['panel'], fg=THEME['accent'], font=("Segoe UI", 12, "bold"))
        self.lbl_proj_display.pack(pady=(15, 5))
        tk.Label(parent, text="📊 结果汇总", bg=THEME['panel'], fg='white', font=("Segoe UI", 14, "bold")).pack(pady=5)
        
        self.stat_labels = {}
        fields = [
            ("Count", "样本数 N"), ("Mean", "均值 Mean"), ("StdDev", "标准差 Std"),
            (None, None),
            ("USL", "规格上限"), ("LSL", "规格下限"),
            (None, None),
            ("Cp", "Cp"), ("Cpk", "Cpk"), ("CPK_LEVEL", "等级"),
            (None, None),
            ("CPU", "CPU"), ("CPL", "CPL"), ("PPM", "PPM")
        ]
        
        tbl = tk.Frame(parent, bg=THEME['panel'])
        tbl.pack(fill=tk.X, padx=20)
        
        for i, (key, label) in enumerate(fields):
            if key is None:
                tk.Frame(tbl, bg=THEME['border'], height=1).grid(row=i, column=0, columnspan=2, sticky='ew', pady=8)
            else:
                tk.Label(tbl, text=label, bg=THEME['panel'], fg='#888', anchor='w', width=12).grid(row=i, column=0, sticky='w', padx=(0, 5))
                val = tk.Label(tbl, text="-", bg=THEME['panel'], fg='white', font=("Consolas", 11, "bold"), anchor='w', width=18)
                val.grid(row=i, column=1, sticky='ew')
                self.stat_labels[key] = val
        tbl.columnconfigure(1, weight=1)

    def update_stats_display(self, stats, project_name=None):
        if project_name:
            pname = project_name
        else:
            pname = self.inp_project.get().strip() or "未命名项目"
            
        self.lbl_proj_display.config(text=f"项目：{pname}")
        if not project_name: 
            self.project_name = pname

        def fmt_val(v, precision=4):
            if v is None: return "N/A"
            return f"{v:.{precision}f}"

        for key, lbl in self.stat_labels.items():
            if key in stats:
                val = stats[key]
                if key == "Count": txt = f"{int(val)}"
                elif key == "PPM": txt = f"{int(val)}"
                elif key == "CPK_LEVEL": txt = f"{val}"
                elif key in ["USL", "LSL"] and val is None: txt = "Not Set"
                else: txt = fmt_val(val, 3 if key in ['Cp', 'Cpk', 'CPU', 'CPL'] else 4)
                
                if key == "CPK_LEVEL":
                    level_colors = {"优秀": THEME['success'], "良好": "#aaff00", "一般": THEME['warning'], "较差": "#ffaa00", "很差": THEME['danger']}
                    color = level_colors.get(val, 'white')
                else:
                    color = THEME['success'] if val is not None else '#666'
                lbl.config(text=txt, fg=color)

    def init_chart_panel(self, parent):
        self.fig, self.ax = plt.subplots(figsize=(5, 5))
        self.fig.patch.set_facecolor(THEME['bg'])
        self.canvas = FigureCanvasTkAgg(self.fig, master=parent)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        self.reset_chart()

    def create_input(self, parent, label, row, default=""):
        tk.Label(parent, text=label, bg=THEME['panel'], fg=THEME['fg']).grid(row=row, column=0, sticky='w', pady=8)
        e = tk.Entry(parent, bg=THEME['input_bg'], fg='white', insertbackground='white', relief=tk.FLAT, font=('Consolas', 11))
        if default: e.insert(0, default)
        e.grid(row=row, column=1, sticky='ew', padx=10, pady=8)
        return e

    def setup_tab1(self, f):
        inner = tk.Frame(f, bg=THEME['panel'], padx=20, pady=20)
        inner.pack(fill=tk.BOTH, expand=True)
        inner.columnconfigure(1, weight=1)
        
        tk.Label(inner, text="规格设置:", bg=THEME['panel'], fg='#888', font=('Microsoft YaHei', 9)).grid(row=0, column=0, columnspan=2, sticky='w')
        self.inp_an_usl = self.create_input(inner, "上限 (USL)", 1)
        self.inp_an_lsl = self.create_input(inner, "下限 (LSL)", 2)
        tk.Label(inner, text="测量数据:", bg=THEME['panel'], fg=THEME['fg']).grid(row=3, column=0, sticky='w', pady=(20, 5))
        self.txt_data = tk.Text(inner, bg=THEME['input_bg'], fg='white', height=18, relief=tk.FLAT, font=('Consolas', 10), width=45)
        self.txt_data.grid(row=4, column=0, columnspan=2, sticky='nsew')
        self.create_btn_bar(inner, 5, self.on_analyze, self.on_clear_tab1, "开始分析")

    def setup_tab2(self, f):
        inner = tk.Frame(f, bg=THEME['panel'], padx=20, pady=20)
        inner.pack(fill=tk.BOTH, expand=True)
        inner.columnconfigure(1, weight=1)
        
        tk.Label(inner, text="规格设置:", bg=THEME['panel'], fg='#888', font=('Microsoft YaHei', 9)).grid(row=0, column=0, columnspan=2, sticky='w')
        self.inp_sim_usl = self.create_input(inner, "上限 (USL)", 1)
        self.inp_sim_lsl = self.create_input(inner, "下限 (LSL)", 2)
        self.inp_sim_cpk = self.create_input(inner, "目标 Cpk", 3, "1.33")
        self.inp_sim_mean = self.create_input(inner, "目标均值", 4, "10.0")
        self.inp_sim_cnt = self.create_input(inner, "数量", 5, "50")
        self.inp_sim_prec = self.create_input(inner, "小数精度", 6, "3")
        self.create_btn_bar(inner, 7, self.on_simulate, self.on_clear_tab2, "生成数据")
        tk.Label(inner, text="结果预览:", bg=THEME['panel'], fg=THEME['fg']).grid(row=8, column=0, sticky='w', pady=(10,0))
        self.txt_sim = tk.Text(inner, bg=THEME['input_bg'], fg=THEME['success'], height=10, relief=tk.FLAT, font=('Consolas', 10), width=45)
        self.txt_sim.grid(row=9, column=0, columnspan=2, sticky='nsew')
        tk.Button(inner, text="📋 复制结果", bg='#444', fg='white', relief=tk.FLAT, command=self.on_copy).grid(row=10, column=0, columnspan=2, sticky='ew', pady=5)

    def setup_tab3(self, f):
        f.columnconfigure(1, weight=1) 
        f.rowconfigure(1, weight=1)   
        
        top_frame = tk.Frame(f, bg=THEME['panel'])
        top_frame.grid(row=0, column=0, columnspan=2, sticky='ew', pady=15, padx=20)
        
        tk.Button(top_frame, text="📂 导入 Excel 文件", bg=THEME['accent'], fg='white', relief=tk.FLAT, 
                  font=('Microsoft YaHei', 10, 'bold'), command=self.load_excel_file, width=15).pack(side=tk.LEFT, padx=(0, 10))
        
        tk.Button(top_frame, text="🗑️ 清空列表", bg=THEME['danger'], fg='white', relief=tk.FLAT, 
                  font=('Microsoft YaHei', 10), command=self.clear_excel_data, width=12).pack(side=tk.LEFT)
        
        info_label = tk.Label(top_frame, text="导入后点击列表项即可查看分析结果", bg=THEME['panel'], fg='#888', font=('Microsoft YaHei', 9))
        info_label.pack(side=tk.RIGHT)

        list_frame = tk.Frame(f, bg=THEME['panel'], width=200)
        list_frame.grid(row=1, column=0, sticky='nsew', padx=(20, 10), pady=(0, 20))
        list_frame.pack_propagate(False)
        
        cols = ('cpk', 'level')
        self.tree_projects = ttk.Treeview(list_frame, columns=cols, displaycolumns=cols, selectmode='browse')
        self.tree_projects.heading('#0', text='项目名称', anchor='w')
        self.tree_projects.heading('cpk', text='Cpk', anchor='center')
        self.tree_projects.heading('level', text='等级', anchor='center')
        
        self.tree_projects.column('#0', width=140, anchor='w')
        self.tree_projects.column('cpk', width=50, anchor='center')
        self.tree_projects.column('level', width=50, anchor='center')
        
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree_projects.yview)
        self.tree_projects.configure(yscrollcommand=scrollbar.set)
        
        self.tree_projects.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tree_projects.bind('<<TreeviewSelect>>', self.on_excel_item_select)

        preview_frame = tk.Frame(f, bg=THEME['panel'])
        preview_frame.grid(row=1, column=1, sticky='nsew', padx=(0, 20), pady=(0, 20))
        
        tk.Label(preview_frame, text="📋 数据预览 & 详情", bg=THEME['panel'], fg=THEME['fg'], font=('Microsoft YaHei', 10, 'bold')).pack(anchor='w', pady=(0, 10))
        
        self.txt_excel_preview = tk.Text(preview_frame, bg=THEME['input_bg'], fg='#ddd', relief=tk.FLAT, font=('Consolas', 9))
        self.txt_excel_preview.pack(fill=tk.BOTH, expand=True)

    def create_btn_bar(self, parent, row, cmd1, cmd2, lbl1):
        box = tk.Frame(parent, bg=THEME['panel'])
        box.grid(row=row, column=0, columnspan=2, sticky='ew', pady=15)
        box.columnconfigure(0, weight=1); box.columnconfigure(1, weight=1)
        tk.Button(box, text=lbl1, bg=THEME['accent'], fg='white', relief=tk.FLAT, font=('bold'), command=cmd1).grid(row=0, column=0, sticky='ew', padx=(0,5), ipady=5)
        tk.Button(box, text="清空", bg=THEME['danger'], fg='white', relief=tk.FLAT, command=cmd2).grid(row=0, column=1, sticky='ew', padx=(5,0), ipady=5)

    def get_val(self, entry, is_int=False, allow_empty=False):
        val_str = entry.get().strip()
        if not val_str: return None if allow_empty else False
        try:
            val = float(val_str)
            return int(val) if is_int else val
        except: return False

    def on_analyze(self):
        usl = self.get_val(self.inp_an_usl, allow_empty=True)
        lsl = self.get_val(self.inp_an_lsl, allow_empty=True)
        if usl is False or lsl is False: DarkMessageBox(self.root, "输入错误", "规格值必须是数字"); return
        if usl is None and lsl is None: DarkMessageBox(self.root, "缺失规格", "请至少输入一个规格限"); return
        raw = self.txt_data.get("1.0", tk.END)
        nums = re.findall(r"[-+]?\d*\.\d+|\d+", raw)
        try:
            data = np.array([float(x) for x in nums])
            if len(data) < 2: raise ValueError
        except: DarkMessageBox(self.root, "数据错误", "请检查输入数据"); return
        self.process_result(data, usl, lsl)

    def on_simulate(self):
        usl = self.get_val(self.inp_sim_usl, allow_empty=True)
        lsl = self.get_val(self.inp_sim_lsl, allow_empty=True)
        cpk = self.get_val(self.inp_sim_cpk); mean = self.get_val(self.inp_sim_mean)
        cnt = self.get_val(self.inp_sim_cnt, True); prec = self.get_val(self.inp_sim_prec, True)
        if any(x is False for x in [usl, lsl, cpk, mean, cnt, prec]): DarkMessageBox(self.root, "输入错误", "请检查数值格式"); return
        if usl is None and lsl is None: DarkMessageBox(self.root, "缺失规格", "请至少输入一个规格限"); return
        data = CpkCalculator.simulate(cpk, mean, usl, lsl, cnt, max(0, min(prec, 10)))
        if data is None: DarkMessageBox(self.root, "错误", "无法生成数据"); return
        fmt = f"{{:.{prec}f}}"
        self.txt_sim.delete("1.0", tk.END)
        self.txt_sim.insert(tk.END, "\n".join([fmt.format(x) for x in data]))
        self.process_result(data, usl, lsl)

    def process_result(self, data, usl, lsl, project_name=None):
        stats = CpkCalculator.calculate(data, usl, lsl)
        if "Error" in stats: DarkMessageBox(self.root, "计算错误", stats["Error"]); return
        self.current_data = data; self.current_stats = stats
        self.current_usl = usl; self.current_lsl = lsl
        self.update_stats_display(stats, project_name)
        self.draw_chart(data, stats)
        if "Excel" not in self.main_notebook.tab('current', 'text'):
            self.btn_export.config(state=tk.NORMAL)

    def draw_chart(self, data, stats):
        if not hasattr(self, '_is_exporting') or not self._is_exporting:
            self.ax.clear(); self.ax.set_facecolor(THEME['bg'])
            mu, sigma = stats['Mean'], stats['StdDev']
            usl, lsl = stats['USL'], stats['LSL']
            
            self.ax.hist(data, bins=30, density=True, alpha=0.6, color=THEME['accent'], edgecolor='none')
            
            xmin, xmax = self.ax.get_xlim()
            base_span = 6 * sigma if sigma > 0 else 1.0
            plot_min = lsl - base_span*0.2 if lsl is not None else min(xmin, mu - 4*sigma)
            plot_max = usl + base_span*0.2 if usl is not None else max(xmax, mu + 4*sigma)
            
            x = np.linspace(plot_min, plot_max, 500)
            y = norm.pdf(x, mu, sigma)
            self.ax.plot(x, y, color=THEME['success'], linewidth=2)
            self.ax.fill_between(x, y, alpha=0.2, color=THEME['success'])
            
            ymax = max(y) * 1.2 if len(y)>0 else 1
            self.ax.set_ylim(0, ymax)
            
            if usl is not None:
                self.ax.axvline(usl, c=THEME['danger'], ls='--', lw=1.5)
                self.ax.text(usl, ymax*0.95, "USL", c=THEME['danger'], ha='center', fontsize=9)
            if lsl is not None:
                self.ax.axvline(lsl, c=THEME['danger'], ls='--', lw=1.5)
                self.ax.text(lsl, ymax*0.95, "LSL", c=THEME['danger'], ha='center', fontsize=9)
            
            self.ax.axvline(mu, c=THEME['warning'], ls='-', lw=1.5, alpha=0.8)
            self.ax.text(mu, ymax*0.85, f"μ={mu:.3f}", c=THEME['warning'], ha='center', fontsize=9)
            
            self.ax.set_xlabel("测量值 (Value)", fontsize=10, color='#aaaaaa', labelpad=5)
            self.ax.set_ylabel("概率密度 (Density)", fontsize=10, color='#aaaaaa', labelpad=5)
            self.ax.tick_params(colors='#888', labelsize=9)
            
            self.ax.spines['top'].set_visible(False); self.ax.spines['right'].set_visible(False)
            self.ax.spines['left'].set_color(THEME['border']); self.ax.spines['bottom'].set_color(THEME['border'])
            
            self.canvas.draw()

    def reset_chart(self):
        if not hasattr(self, '_is_exporting') or not self._is_exporting:
            self.ax.clear(); self.ax.axis('off')
            self.ax.text(0.5, 0.5, "等待数据...", color='#555', ha='center', transform=self.ax.transAxes)
            self.canvas.draw()
            for k, l in self.stat_labels.items(): l.config(text="-", fg='white')
            self.lbl_proj_display.config(text="项目：未命名")
            self.current_data = None; self.current_stats = None

    def on_clear_tab1(self):
        self.inp_an_usl.delete(0, tk.END); self.inp_an_lsl.delete(0, tk.END)
        self.txt_data.delete("1.0", tk.END); self.reset_chart()
        self.btn_export.config(state=tk.DISABLED)

    def on_clear_tab2(self):
        for e in [self.inp_sim_usl, self.inp_sim_lsl, self.inp_sim_cpk, self.inp_sim_mean, self.inp_sim_cnt, self.inp_sim_prec]: e.delete(0, tk.END)
        self.txt_sim.delete("1.0", tk.END); self.reset_chart()
        self.btn_export.config(state=tk.DISABLED)

    def on_copy(self):
        self.root.clipboard_clear(); self.root.clipboard_append(self.txt_sim.get("1.0", tk.END))
        DarkMessageBox(self.root, "复制成功", "内容已复制到剪贴板", False)

    def add_about_link(self):
        about_label = tk.Label(self.root, text="关于软件", bg=THEME['bg'], fg=THEME['accent'], font=("Microsoft YaHei", 9, "underline"), cursor="hand2")
        about_label.place(relx=1.0, rely=1.0, anchor='se', x=-15, y=-10)
        about_label.bind("<Button-1>", lambda e: self.show_about())
        about_label.bind("<Enter>", lambda e: about_label.config(fg=THEME['success']))
        about_label.bind("<Leave>", lambda e: about_label.config(fg=THEME['accent']))

    def show_about(self):
        about_text = "CPK 统计分析工具 V6.0 (Merge PDF)\n\n新功能:\n• 单个导出：标准“另存为”对话框\n• 批量导出：将所有项目合并为【一份】PDF文件\n  - 调用标准“另存为”对话框\n  - 一次保存，生成汇总报告"
        DarkMessageBox(self.root, "关于软件", about_text, is_error=False)

    # ==========================================
    # 5. Excel 导入相关功能
    # ==========================================
    def load_excel_file(self):
        if not PANDAS_AVAILABLE:
            DarkMessageBox(self.root, "缺少依赖", "未安装 pandas 或 openpyxl。\n请运行：pip install pandas openpyxl")
            return

        file_path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if not file_path:
            return

        try:
            df = pd.read_excel(file_path, header=None)
            
            if df.shape[0] < 4:
                DarkMessageBox(self.root, "格式错误", "Excel 文件至少需要 4 行数据:\n1. 项目名\n2. USL\n3. LSL\n4. 数据起始行")
                return

            new_projects = []
            num_cols = df.shape[1]
            error_logs = []

            for col_idx in range(num_cols):
                col_data = df.iloc[:, col_idx]
                
                project_name = col_data.iloc[0]
                if pd.isna(project_name) or str(project_name).strip() == "":
                    project_name = f"Project_{col_idx + 1}"
                else:
                    project_name = str(project_name).strip()

                usl_val = col_data.iloc[1]
                usl = float(usl_val) if not pd.isna(usl_val) else None

                lsl_val = col_data.iloc[2]
                lsl = float(lsl_val) if not pd.isna(lsl_val) else None

                raw_data = col_data.iloc[3:].dropna()
                try:
                    data_array = np.array(raw_data.astype(float))
                except ValueError:
                    error_logs.append(f"{project_name}: 数据格式错误")
                    continue 

                if len(data_array) < 2:
                    error_logs.append(f"{project_name}: 数据量不足")
                    continue 

                stats = CpkCalculator.calculate(data_array, usl, lsl)
                
                if "Error" in stats:
                    error_logs.append(f"{project_name}: {stats['Error']}")
                    continue

                new_projects.append({
                    "name": project_name,
                    "data": data_array,
                    "usl": usl,
                    "lsl": lsl,
                    "stats": stats,
                    "cpk_val": stats['Cpk'],
                    "level": stats['CPK_LEVEL']
                })

            if not new_projects:
                msg = "未找到有效数据。"
                if error_logs:
                    msg += "\n\n错误详情:\n" + "\n".join(error_logs[:5])
                DarkMessageBox(self.root, "导入失败", msg)
                return

            self.excel_projects = new_projects
            self.refresh_excel_treeview()
            
            if self.excel_projects:
                self.tree_projects.selection_set(self.tree_projects.get_children()[0])
                self.on_excel_item_select(None)
            
            msg = f"成功导入 {len(self.excel_projects)} 个项目。"
            if error_logs:
                msg += f"\n跳过 {len(error_logs)} 个无效项目。"
            DarkMessageBox(self.root, "导入成功", msg, is_error=False)

        except Exception as e:
            DarkMessageBox(self.root, "导入失败", f"读取 Excel 文件时出错:\n{str(e)}")

    def refresh_excel_treeview(self):
        for item in self.tree_projects.get_children():
            self.tree_projects.delete(item)
        
        for i, proj in enumerate(self.excel_projects):
            cpk = proj['cpk_val']
            level = proj['level']
            self.tree_projects.insert('', 'end', iid=str(i), text=proj['name'], values=(f"{cpk:.3f}", level))

    def on_excel_item_select(self, event):
        selection = self.tree_projects.selection()
        if not selection:
            return
        
        iid = selection[0]
        idx = int(iid)
        
        if idx < 0 or idx >= len(self.excel_projects):
            return
            
        self.current_excel_index = idx
        project = self.excel_projects[idx]
        
        self.update_stats_display(project['stats'], project_name=project['name'])
        self.draw_chart(project['data'], project['stats'])
        
        self.txt_excel_preview.delete("1.0", tk.END)
        preview_text = f"【{project['name']}】详细报告\n"
        preview_text += "="*30 + "\n"
        preview_text += f"Cpk: {project['cpk_val']:.4f} ({project['level']})\n"
        preview_text += f"USL: {project['usl']} | LSL: {project['lsl']}\n"
        preview_text += f"Mean: {project['stats']['Mean']:.4f} | Std: {project['stats']['StdDev']:.4f}\n"
        preview_text += f"PPM: {int(project['stats']['PPM'])}\n"
        preview_text += "="*30 + "\n\n数据前 30 行:\n"
        
        for i, val in enumerate(project['data'][:30]):
            preview_text += f"{val}\n"
        if len(project['data']) > 30:
            preview_text += f"... 共 {len(project['data'])} 条数据"
        
        self.txt_excel_preview.insert("1.0", preview_text)
        
        if "Excel" in self.main_notebook.tab('current', 'text'):
            self.btn_export.config(state=tk.NORMAL)

    def clear_excel_data(self):
        self.excel_projects = []
        self.current_excel_index = -1
        for item in self.tree_projects.get_children():
            self.tree_projects.delete(item)
        self.txt_excel_preview.delete("1.0", tk.END)
        self.reset_chart()
        if "Excel" in self.main_notebook.tab('current', 'text'):
            self.btn_export.config(state=tk.DISABLED)

    # 【核心修改】批量导出：合并为一份PDF，使用标准另存为对话框
    def export_merged_report(self):
        if not self.excel_projects:
            DarkMessageBox(self.root, "无数据", "没有可导出的项目数据。请先导入 Excel。")
            return
        
        if not REPORTLAB_AVAILABLE:
            DarkMessageBox(self.root, "缺少依赖", "未安装 reportlab。")
            return

        # 生成默认文件名
        default_filename = f"CPK_汇总报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"

        # 【关键】调用标准的“另存为”对话框，让用户决定这一个文件的名字
        file_path = filedialog.asksaveasfilename(
            title="保存汇总报告 (所有项目将合并为此文件)",
            defaultextension=".pdf",
            initialfile=default_filename,
            filetypes=[("PDF 文件", "*.pdf"), ("所有文件", "*.*")],
            initialdir=os.path.expanduser("~")
        )
        
        if not file_path:
            return

        # 确保扩展名
        if not file_path.lower().endswith('.pdf'):
            file_path += '.pdf'

        try:
            self.btn_batch_export.config(state=tk.DISABLED, text="生成中...")
            self.root.update_idletasks()
            
            # 调用合并生成函数
            self._generate_merged_pdf_report(file_path, self.excel_projects)
            
            DarkMessageBox(self.root, "导出成功", f"所有 {len(self.excel_projects)} 个项目已合并保存至:\n{file_path}", is_error=False)

        except Exception as e:
            DarkMessageBox(self.root, "导出失败", f"错误:\n{str(e)}\n{traceback.format_exc()}")
        finally:
            self.btn_batch_export.config(text="📦 合并导出所有为一份PDF")

    # 【单个导出】保持标准另存为
    def export_report(self):
        if not REPORTLAB_AVAILABLE:
            DarkMessageBox(self.root, "缺少依赖", "未安装 reportlab。")
            return

        current_tab_text = self.main_notebook.tab('current', 'text')
        if "Excel" in current_tab_text:
            if not self.excel_projects or self.current_excel_index == -1:
                DarkMessageBox(self.root, "无数据", "请先在列表中选择一个项目。")
                return
            
            project = self.excel_projects[self.current_excel_index]
            data_to_export = project['data']
            stats_to_export = project['stats']
            usl_to_export = project['usl']
            lsl_to_export = project['lsl']
            name_to_export = project['name']
        
        elif self.current_stats is None:
            DarkMessageBox(self.root, "无数据", "请先进行分析或模拟生成数据。")
            return
        else:
            data_to_export = self.current_data
            stats_to_export = self.current_stats
            usl_to_export = self.current_usl
            lsl_to_export = self.current_lsl
            name_to_export = self.project_name

        safe_name = re.sub(r'[^\w\-_\. ]', '_', name_to_export)
        safe_name = safe_name[:50]
        default_filename = f"CPK_{safe_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"

        file_path = filedialog.asksaveasfilename(
            title="保存 CPK 报告",
            defaultextension=".pdf",
            initialfile=default_filename,
            filetypes=[("PDF 文件", "*.pdf"), ("所有文件", "*.*")],
            initialdir=os.path.expanduser("~")
        )
        
        if not file_path:
            return

        if not file_path.lower().endswith('.pdf'):
            file_path += '.pdf'

        try:
            old_data = self.current_data
            old_stats = self.current_stats
            old_usl = self.current_usl
            old_lsl = self.current_lsl
            old_name = self.project_name

            self.current_data = data_to_export
            self.current_stats = stats_to_export
            self.current_usl = usl_to_export
            self.current_lsl = lsl_to_export
            self.project_name = name_to_export

            self._generate_pdf_report(file_path)
            
            self.current_data = old_data
            self.current_stats = old_stats
            self.current_usl = old_usl
            self.current_lsl = old_lsl
            self.project_name = old_name

            DarkMessageBox(self.root, "导出成功", f"报告已保存至:\n{file_path}", is_error=False)
        
        except Exception as e:
            DarkMessageBox(self.root, "导出失败", f"错误:\n{str(e)}\n{traceback.format_exc()}")

    def _generate_merged_pdf_report(self, file_path, projects_list):
        """生成包含所有项目的单一PDF文件"""
        temp_img_paths = []
        figs = []
        
        try:
            doc = SimpleDocTemplate(file_path, pagesize=A4, 
                                    rightMargin=1.5*cm, leftMargin=1.5*cm, 
                                    topMargin=1.5*cm, bottomMargin=1.5*cm)
            story = []
            styles = getSampleStyleSheet()
            
            font_name = "Helvetica"
            try:
                if os.name == 'nt':
                    path = r"C:\Windows\Fonts\simhei.ttf"
                    if os.path.exists(path):
                        pdfmetrics.registerFont(TTFont('SimHei', path)); font_name = 'SimHei'
            except: pass

            title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontName=font_name, fontSize=18, alignment=TA_CENTER, spaceAfter=10, textColor=colors.black)
            sub_style = ParagraphStyle('Sub', parent=styles['Normal'], fontName=font_name, fontSize=10, alignment=TA_CENTER, spaceAfter=20, textColor=colors.gray)
            head_style = ParagraphStyle('Head', parent=styles['Heading3'], fontName=font_name, fontSize=12, spaceBefore=10, spaceAfter=5, textColor=colors.darkblue)
            normal_style = ParagraphStyle('Norm', parent=styles['Normal'], fontName=font_name, fontSize=9, leading=12, textColor=colors.black)

            # 封面/总标题
            story.append(Paragraph("CPK 过程能力汇总分析报告", title_style))
            story.append(Paragraph(f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}  |  包含项目数：{len(projects_list)}", sub_style))
            story.append(PageBreak())

            # 目录（可选，简单列出）
            story.append(Paragraph("目录", head_style))
            toc_data = []
            for i, proj in enumerate(projects_list):
                toc_data.append([f"{i+1}. {proj['name']}", f"Cpk: {proj['cpk_val']:.3f} ({proj['level']})"])
            
            t_toc = Table(toc_data, colWidths=[12*cm, 4*cm])
            t_toc.setStyle(TableStyle([
                ('FONTNAME', (0, 0), (-1, -1), font_name), ('FONTSIZE', (0, 0), (-1, -1), 9),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ('GRID', (0, 0), (-1, -1), 0.25, colors.lightgrey),
            ]))
            story.append(t_toc)
            story.append(PageBreak())

            # 循环添加每个项目
            for i, proj in enumerate(projects_list):
                # 临时设置当前状态以便复用绘图和表格逻辑
                self.current_data = proj['data']
                self.current_stats = proj['stats']
                self.current_usl = proj['usl']
                self.current_lsl = proj['lsl']
                self.project_name = proj['name']
                
                # 项目标题
                story.append(Paragraph(f"项目 {i+1}: {proj['name']}", head_style))
                story.append(Spacer(1, 0.1*inch))

                # 1. 添加统计表格
                s = self.current_stats
                highlight_color = colors.Color(0.93, 0.96, 1.0)
                core_data = [
                    [
                        Paragraph("<b>USL</b><br/>" + ("-" if s['USL'] is None else f"{s['USL']:.4f}"), normal_style),
                        Paragraph("<b>LSL</b><br/>" + ("-" if s['LSL'] is None else f"{s['LSL']:.4f}"), normal_style),
                        Paragraph("<b>Cpk</b><br/><font size=12 color='darkred'>" + ("-" if s['Cpk'] is None else f"{s['Cpk']:.4f}") + "</font>", normal_style),
                        Paragraph("<b>PPM</b><br/><font size=12 color='darkred'>" + ("-" if s['PPM'] is None else f"{int(s['PPM'])}") + "</font>", normal_style)
                    ],
                    [
                        Paragraph("Mean: " + ("-" if s['Mean'] is None else f"{s['Mean']:.4f}") + "<br/>Std: " + ("-" if s['StdDev'] is None else f"{s['StdDev']:.4f}"), ParagraphStyle('Small', parent=normal_style, fontSize=8)),
                        Paragraph("Cp: " + ("-" if s['Cp'] is None else f"{s['Cp']:.4f}"), ParagraphStyle('Small', parent=normal_style, fontSize=8)),
                        Paragraph("等级：<b>" + f"{s['CPK_LEVEL']}" + "</b>", ParagraphStyle('Small', parent=normal_style, fontSize=8)),
                        Paragraph("N: <b>" + f"{int(s['Count'])}" + "</b>", ParagraphStyle('Small', parent=normal_style, fontSize=8))
                    ]
                ]
                
                t_core = Table(core_data, colWidths=[4.0*cm, 4.0*cm, 4.0*cm, 4.0*cm])
                t_core.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), highlight_color),
                    ('BOX', (0, 0), (-1, 0), 1.5, colors.darkblue),
                    ('INNERGRID', (0, 0), (-1, 0), 0.5, colors.white),
                    ('ALIGN', (0, 0), (-1, 0), 'CENTER'), ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 6), ('TOPPADDING', (0, 0), (-1, 0), 6),
                    ('FONTNAME', (0, 0), (-1, 0), font_name), ('FONTSIZE', (0, 0), (-1, 0), 9),
                    ('BACKGROUND', (0, 1), (-1, 1), colors.white),
                    ('BOX', (0, 1), (-1, 1), 0.5, colors.lightgrey),
                    ('ALIGN', (0, 1), (-1, 1), 'CENTER'), ('VALIGN', (0, 1), (-1, 1), 'MIDDLE'),
                    ('FONTNAME', (0, 1), (-1, 1), font_name), ('FONTSIZE', (0, 1), (-1, 1), 8),
                ]))
                story.append(t_core)
                story.append(Spacer(1, 0.15*inch))

                # 2. 添加图表
                # 生成临时图片
                with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp_file:
                    temp_img_path = tmp_file.name
                    temp_img_paths.append(temp_img_path)
                
                fig, ax = plt.subplots(figsize=(7.5, 4.8), dpi=300)
                figs.append(fig) # 保存引用以便关闭
                fig.patch.set_facecolor('white')
                ax.set_facecolor('white')
                
                data = self.current_data
                mu, sigma = s['Mean'], s['StdDev']
                usl, lsl = s['USL'], s['LSL']
                
                if data is not None and len(data) > 0:
                    ax.hist(data, bins=30, density=True, alpha=0.7, color='#3b82f6', edgecolor='white', linewidth=0.5)
                    xmin, xmax = ax.get_xlim()
                    base_span = 6 * sigma if sigma > 0 else 1.0
                    plot_min = lsl - base_span*0.2 if lsl is not None else min(xmin, mu - 4*sigma)
                    plot_max = usl + base_span*0.2 if usl is not None else max(xmax, mu + 4*sigma)
                    x = np.linspace(plot_min, plot_max, 500)
                    y = norm.pdf(x, mu, sigma)
                    ax.plot(x, y, color='#dc2626', linewidth=2.5)
                    ax.fill_between(x, y, alpha=0.2, color='#dc2626')
                    
                    ymax = max(y) * 1.15 if len(y)>0 else 1
                    ax.set_ylim(0, ymax)
                    
                    if usl is not None:
                        ax.axvline(usl, c='#ef4444', ls='--', lw=2)
                        ax.text(usl, ymax*0.95, f"USL", c='#ef4444', ha='center', fontweight='bold', fontsize=9)
                    if lsl is not None:
                        ax.axvline(lsl, c='#ef4444', ls='--', lw=2)
                        ax.text(lsl, ymax*0.95, f"LSL", c='#ef4444', ha='center', fontweight='bold', fontsize=9)
                    ax.axvline(mu, c='#16a34a', ls='-', lw=2)
                    
                    ax.set_title(f"Cpk = {s['Cpk']:.3f}", fontsize=12, fontweight='bold', pad=10)
                    ax.set_xlabel("Value", fontsize=9, color='black')
                    ax.set_ylabel("Density", fontsize=9, color='black')
                    ax.tick_params(colors='black', labelsize=8)
                    ax.grid(True, linestyle='--', alpha=0.3)
                    ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
                    
                    fig.tight_layout()
                    fig.savefig(temp_img_path, dpi=300, bbox_inches='tight', facecolor='white')
                
                plt.close(fig)
                
                img = Image(temp_img_path, width=6.5*inch, height=4.2*inch)
                story.append(img)
                story.append(Spacer(1, 0.1*inch))

                # 3. 添加部分数据明细 (限制行数以防文件过大)
                story.append(Paragraph("数据明细 (前50条)", ParagraphStyle('SmallHead', parent=normal_style, fontSize=9, spaceBefore=5)))
                d_list = self.current_data
                if isinstance(d_list, np.ndarray):
                    d_list = d_list.tolist()
                
                cols = 8
                rows_data = []
                subset = d_list[:50] if len(d_list) > 50 else d_list
                
                for i_row in range(0, len(subset), cols):
                    row_slice = subset[i_row : i_row+cols]
                    while len(row_slice) < cols:
                        row_slice.append("")
                    formatted_row = [f"{x:.3f}" if x != "" else "" for x in row_slice]
                    rows_data.append(formatted_row)
                
                if rows_data:
                    t_data = Table(rows_data, colWidths=[2.0*cm] * cols)
                    t_data.setStyle(TableStyle([
                        ('FONTNAME', (0, 0), (-1, -1), font_name), ('FONTSIZE', (0, 0), (-1, -1), 7),
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                        ('GRID', (0, 0), (-1, -1), 0.25, colors.lightgrey),
                        ('BACKGROUND', (0, 0), (-1, 0), colors.Color(0.9, 0.9, 0.9)),
                    ]))
                    story.append(t_data)
                
                if len(d_list) > 50:
                    story.append(Paragraph(f"... 共 {len(d_list)} 条数据，此处仅显示前 50 条", ParagraphStyle('Note', parent=normal_style, fontSize=7, textColor=colors.gray)))

                # 如果不是最后一个项目，加分页符
                if i < len(projects_list) - 1:
                    story.append(PageBreak())

            doc.build(story)
            
        finally:
            # 清理临时图片
            for path in temp_img_paths:
                if os.path.exists(path):
                    try: os.unlink(path)
                    except: pass
            # 关闭图表
            for fig in figs:
                try: plt.close(fig)
                except: pass

    def _generate_pdf_report(self, file_path):
        """原有的单个报告生成逻辑（略作简化，复用上面的逻辑片段）"""
        # 为了代码简洁，这里直接调用合并逻辑的简化版，或者保留原有逻辑
        # 既然上面已经有了完整的单项目生成逻辑片段，这里为了稳健性，保留原有独立函数结构
        # 但为了减少重复代码，实际生产中可以重构。这里保持原样以确保单个导出逻辑不变。
        
        temp_img_path = None
        fig = None
        try:
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp_file:
                temp_img_path = tmp_file.name
            
            fig, ax = plt.subplots(figsize=(7.5, 4.8), dpi=300)
            fig.patch.set_facecolor('white')
            ax.set_facecolor('white')
            
            data = self.current_data
            stats = self.current_stats
            if data is None or stats is None:
                raise ValueError("当前无有效数据可导出")
                
            mu, sigma = stats['Mean'], stats['StdDev']
            usl, lsl = stats['USL'], stats['LSL']
            
            ax.hist(data, bins=30, density=True, alpha=0.7, color='#3b82f6', edgecolor='white', linewidth=0.5)
            xmin, xmax = ax.get_xlim()
            base_span = 6 * sigma if sigma > 0 else 1.0
            plot_min = lsl - base_span*0.2 if lsl is not None else min(xmin, mu - 4*sigma)
            plot_max = usl + base_span*0.2 if usl is not None else max(xmax, mu + 4*sigma)
            x = np.linspace(plot_min, plot_max, 500)
            y = norm.pdf(x, mu, sigma)
            ax.plot(x, y, color='#dc2626', linewidth=2.5)
            ax.fill_between(x, y, alpha=0.2, color='#dc2626')
            
            ymax = max(y) * 1.15 if len(y)>0 else 1
            ax.set_ylim(0, ymax)
            
            if usl is not None:
                ax.axvline(usl, c='#ef4444', ls='--', lw=2)
                ax.text(usl, ymax*0.95, f"USL", c='#ef4444', ha='center', fontweight='bold', fontsize=9)
            if lsl is not None:
                ax.axvline(lsl, c='#ef4444', ls='--', lw=2)
                ax.text(lsl, ymax*0.95, f"LSL", c='#ef4444', ha='center', fontweight='bold', fontsize=9)
            ax.axvline(mu, c='#16a34a', ls='-', lw=2)
            
            title_str = f"Project: {self.project_name}  |  Cpk = {stats['Cpk']:.3f}"
            ax.set_title(title_str, fontsize=12, fontweight='bold', pad=10)
            ax.set_xlabel("Measurement Value", fontsize=10, color='black', labelpad=5)
            ax.set_ylabel("Probability Density", fontsize=10, color='black', labelpad=5)
            ax.tick_params(colors='black', labelsize=9)
            
            ax.grid(True, linestyle='--', alpha=0.3, color='gray')
            ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
            ax.spines['left'].set_color('black'); ax.spines['bottom'].set_color('black')
            
            fig.tight_layout()
            fig.savefig(temp_img_path, dpi=300, bbox_inches='tight', facecolor='white')
            plt.close(fig) 
            fig = None

            doc = SimpleDocTemplate(file_path, pagesize=A4, 
                                    rightMargin=1.5*cm, leftMargin=1.5*cm, 
                                    topMargin=1.5*cm, bottomMargin=1.5*cm)
            story = []
            styles = getSampleStyleSheet()
            
            font_name = "Helvetica"
            try:
                if os.name == 'nt':
                    path = r"C:\Windows\Fonts\simhei.ttf"
                    if os.path.exists(path):
                        pdfmetrics.registerFont(TTFont('SimHei', path)); font_name = 'SimHei'
            except: pass

            title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontName=font_name, fontSize=16, alignment=TA_CENTER, spaceAfter=5, textColor=colors.black)
            sub_style = ParagraphStyle('Sub', parent=styles['Normal'], fontName=font_name, fontSize=9, alignment=TA_CENTER, spaceAfter=12, textColor=colors.gray)
            head_style = ParagraphStyle('Head', parent=styles['Heading3'], fontName=font_name, fontSize=10, spaceBefore=8, spaceAfter=4, textColor=colors.darkblue)
            normal_style = ParagraphStyle('Norm', parent=styles['Normal'], fontName=font_name, fontSize=9, leading=12, textColor=colors.black)
            
            story.append(Paragraph("CPK 过程能力分析报表", title_style))
            story.append(Paragraph(f"项目名称：{self.project_name}  |  生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}", sub_style))
            
            def safe_fmt(val, fmt_str="{:.4f}"):
                if val is None: return "-"
                return fmt_str.format(val)

            s = self.current_stats
            highlight_color = colors.Color(0.93, 0.96, 1.0)
            core_data = [
                [
                    Paragraph("<b>USL (上限)</b><br/>" + safe_fmt(s['USL']), normal_style),
                    Paragraph("<b>LSL (下限)</b><br/>" + safe_fmt(s['LSL']), normal_style),
                    Paragraph("<b>Cpk (能力指数)</b><br/><font size=14 color='darkred'>" + safe_fmt(s['Cpk']) + "</font>", normal_style),
                    Paragraph("<b>PPM (不良率)</b><br/><font size=14 color='darkred'>" + (f"{int(s['PPM'])}" if s['PPM'] is not None else "-") + "</font>", normal_style)
                ],
                [
                    Paragraph("Mean: " + safe_fmt(s['Mean']) + "<br/>Std: " + safe_fmt(s['StdDev']), ParagraphStyle('Small', parent=normal_style, fontSize=8)),
                    Paragraph("Cp: " + safe_fmt(s['Cp']), ParagraphStyle('Small', parent=normal_style, fontSize=8)),
                    Paragraph("等级：<b>" + f"{s['CPK_LEVEL']}" + "</b>", ParagraphStyle('Small', parent=normal_style, fontSize=8)),
                    Paragraph("样本数 N: <b>" + f"{int(s['Count'])}" + "</b>", ParagraphStyle('Small', parent=normal_style, fontSize=8))
                ]
            ]
            
            t_core = Table(core_data, colWidths=[4.0*cm, 4.0*cm, 4.0*cm, 4.0*cm])
            t_core.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), highlight_color),
                ('BOX', (0, 0), (-1, 0), 1.5, colors.darkblue),
                ('INNERGRID', (0, 0), (-1, 0), 0.5, colors.white),
                ('ALIGN', (0, 0), (-1, 0), 'CENTER'), ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 8), ('TOPPADDING', (0, 0), (-1, 0), 8),
                ('FONTNAME', (0, 0), (-1, 0), font_name), ('FONTSIZE', (0, 0), (-1, 0), 9),
                ('BACKGROUND', (0, 1), (-1, 1), colors.white),
                ('BOX', (0, 1), (-1, 1), 0.5, colors.lightgrey),
                ('ALIGN', (0, 1), (-1, 1), 'CENTER'), ('VALIGN', (0, 1), (-1, 1), 'MIDDLE'),
                ('BOTTOMPADDING', (0, 1), (-1, 1), 5), ('TOPPADDING', (0, 1), (-1, 1), 5),
                ('FONTNAME', (0, 1), (-1, 1), font_name), ('FONTSIZE', (0, 1), (-1, 1), 8),
                ('TEXTCOLOR', (0, 1), (-1, 1), colors.darkgrey),
            ]))
            story.append(t_core)
            story.append(Spacer(1, 0.15*inch))
            
            story.append(Paragraph("分布直方图", head_style))
            img = Image(temp_img_path, width=6.5*inch, height=4.2*inch)
            story.append(img)
            story.append(Spacer(1, 0.1*inch))
            
            story.append(Paragraph("原始数据明细", head_style))
            d_list = self.current_data
            if isinstance(d_list, np.ndarray):
                d_list = d_list.tolist()
            
            cols = 10
            rows_data = []
            display_limit = 200 
            subset = d_list[:display_limit] if len(d_list) > display_limit else d_list
            
            for i in range(0, len(subset), cols):
                row_slice = subset[i : i+cols]
                while len(row_slice) < cols:
                    row_slice.append("")
                formatted_row = [f"{x:.3f}" if x != "" else "" for x in row_slice]
                rows_data.append(formatted_row)
            
            t_data = Table(rows_data, colWidths=[1.5*cm] * cols)
            data_table_style = [
                ('FONTNAME', (0, 0), (-1, -1), font_name), ('FONTSIZE', (0, 0), (-1, -1), 7),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 2), ('RIGHTPADDING', (0, 0), (-1, -1), 2),
                ('TOPPADDING', (0, 0), (-1, -1), 2), ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
                ('GRID', (0, 0), (-1, -1), 0.25, colors.lightgrey),
            ]
            for i in range(len(rows_data)):
                bg_color = colors.Color(0.97, 0.97, 0.97) if i % 2 == 0 else colors.white
                data_table_style.append(('BACKGROUND', (0, i), (-1, i), bg_color))
            
            t_data.setStyle(TableStyle(data_table_style))
            story.append(t_data)
            
            if len(self.current_data) > display_limit:
                note_style = ParagraphStyle('Note', parent=normal_style, fontSize=7, textColor=colors.gray, alignment=TA_CENTER)
                story.append(Paragraph(f"... 共 {len(self.current_data)} 条数据，此处仅显示前 {display_limit} 条", note_style))
            
            doc.build(story)
            
        finally:
            if temp_img_path and os.path.exists(temp_img_path): 
                try:
                    os.unlink(temp_img_path)
                except: pass
            if fig is not None:
                try:
                    plt.close(fig)
                except: pass

if __name__ == "__main__":
    root = tk.Tk()
    app = CpkApp(root)
    root.mainloop()
