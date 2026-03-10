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

# 尝试导入 reportlab
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm, inch
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# ==========================================
# 1. 高分屏适配 (HiDPI)
# ==========================================
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    ScaleFactor = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100
except:
    ScaleFactor = 1.0

# ==========================================
# 2. 全局深色主题 (仅用于 GUI)
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
            if cpk >= 1.67: cpk_level = "优秀 (Excellent)"
            elif cpk >= 1.33: cpk_level = "良好 (Good)"
            elif cpk >= 1.0: cpk_level = "一般 (Adequate)"
            elif cpk >= 0.67: cpk_level = "较差 (Poor)"
            else: cpk_level = "很差 (Inadequate)"

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
        self.root.title("CPK Tool Pro - Fixed Zero Bug")
        self.root.geometry(f"{int(1700)}x{int(850)}") 
        self.root.configure(bg=THEME['bg'])
        
        self.current_data = None
        self.current_stats = None
        self.current_usl = None
        self.current_lsl = None
        self.project_name = ""
        
        self.setup_ui()

    def setup_ui(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TNotebook", background=THEME['bg'], borderwidth=0)
        style.configure("TNotebook.Tab", background=THEME['panel'], foreground=THEME['fg'], padding=[25, 12], font=('Microsoft YaHei', 11))
        style.map("TNotebook.Tab", background=[("selected", THEME['accent'])], foreground=[("selected", 'white')])

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
        tk.Label(proj_frame, text="📁 项目名称 (可选):", bg=THEME['panel'], fg=THEME['warning'], font=('Microsoft YaHei', 9, 'bold')).pack(anchor='w')
        self.inp_project = tk.Entry(proj_frame, bg=THEME['input_bg'], fg='white', insertbackground='white', 
                                    relief=tk.FLAT, font=('Microsoft YaHei', 10))
        self.inp_project.pack(fill=tk.X, ipady=5, pady=(5,0))
        self.inp_project.insert(0, "未命名项目")
        
        nb = ttk.Notebook(parent)
        nb.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
        
        t1 = ttk.Frame(nb); nb.add(t1, text=' 数据分析 ')
        t2 = ttk.Frame(nb); nb.add(t2, text=' 模拟生成 ')
        
        self.setup_tab1(t1)
        self.setup_tab2(t2)
        
        export_frame = tk.Frame(parent, bg=THEME['panel'])
        export_frame.pack(fill=tk.X, padx=15, pady=(10, 20))
        
        self.btn_export = tk.Button(
            export_frame, text="📄 导出专业报告 (PDF)", bg=THEME['success'], fg='#1e1e1e', 
            relief=tk.FLAT, font=('Microsoft YaHei', 11, 'bold'), command=self.export_report, cursor="hand2"
        )
        self.btn_export.pack(fill=tk.X, ipady=8)
        
        if not REPORTLAB_AVAILABLE:
            self.btn_export.config(state=tk.DISABLED, text="❌ 缺少 reportlab\npip install reportlab")
            tk.Label(parent, text="(需安装 reportlab)", bg=THEME['panel'], fg='#666', font=("Microsoft YaHei", 8)).pack(pady=(0, 10))

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

    def update_stats_display(self, stats):
        pname = self.inp_project.get().strip() or "未命名项目"
        self.lbl_proj_display.config(text=f"项目：{pname}")
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
                # 【修复点】这里也应用了 is not None 逻辑，确保界面显示正确
                elif key in ["USL", "LSL"] and val is None: txt = "Not Set"
                else: txt = fmt_val(val, 3 if key in ['Cp', 'Cpk', 'CPU', 'CPL'] else 4)
                
                if key == "CPK_LEVEL":
                    level_colors = {"优秀 (Excellent)": THEME['success'], "良好 (Good)": "#aaff00", "一般 (Adequate)": THEME['warning'], "较差 (Poor)": "#ffaa00", "很差 (Inadequate)": THEME['danger']}
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
        f = tk.Frame(f, bg=THEME['panel'], padx=20, pady=20)
        f.pack(fill=tk.BOTH, expand=True)
        f.columnconfigure(1, weight=1)
        tk.Label(f, text="规格设置:", bg=THEME['panel'], fg='#888', font=('Microsoft YaHei', 9)).grid(row=0, column=0, columnspan=2, sticky='w')
        self.inp_an_usl = self.create_input(f, "上限 (USL)", 1)
        self.inp_an_lsl = self.create_input(f, "下限 (LSL)", 2)
        tk.Label(f, text="测量数据:", bg=THEME['panel'], fg=THEME['fg']).grid(row=3, column=0, sticky='w', pady=(20, 5))
        self.txt_data = tk.Text(f, bg=THEME['input_bg'], fg='white', height=18, relief=tk.FLAT, font=('Consolas', 10), width=45)
        self.txt_data.grid(row=4, column=0, columnspan=2, sticky='nsew')
        self.create_btn_bar(f, 5, self.on_analyze, self.on_clear_tab1, "开始分析")

    def setup_tab2(self, f):
        f = tk.Frame(f, bg=THEME['panel'], padx=20, pady=20)
        f.pack(fill=tk.BOTH, expand=True)
        f.columnconfigure(1, weight=1)
        tk.Label(f, text="规格设置:", bg=THEME['panel'], fg='#888', font=('Microsoft YaHei', 9)).grid(row=0, column=0, columnspan=2, sticky='w')
        self.inp_sim_usl = self.create_input(f, "上限 (USL)", 1)
        self.inp_sim_lsl = self.create_input(f, "下限 (LSL)", 2)
        self.inp_sim_cpk = self.create_input(f, "目标 Cpk", 3, "1.33")
        self.inp_sim_mean = self.create_input(f, "目标均值", 4, "10.0")
        self.inp_sim_cnt = self.create_input(f, "数量", 5, "50")
        self.inp_sim_prec = self.create_input(f, "小数精度", 6, "3")
        self.create_btn_bar(f, 7, self.on_simulate, self.on_clear_tab2, "生成数据")
        tk.Label(f, text="结果预览:", bg=THEME['panel'], fg=THEME['fg']).grid(row=8, column=0, sticky='w', pady=(10,0))
        self.txt_sim = tk.Text(f, bg=THEME['input_bg'], fg=THEME['success'], height=10, relief=tk.FLAT, font=('Consolas', 10), width=45)
        self.txt_sim.grid(row=9, column=0, columnspan=2, sticky='nsew')
        tk.Button(f, text="📋 复制结果", bg='#444', fg='white', relief=tk.FLAT, command=self.on_copy).grid(row=10, column=0, columnspan=2, sticky='ew', pady=5)

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

    def process_result(self, data, usl, lsl):
        stats = CpkCalculator.calculate(data, usl, lsl)
        if "Error" in stats: DarkMessageBox(self.root, "计算错误", stats["Error"]); return
        self.current_data = data; self.current_stats = stats
        self.current_usl = usl; self.current_lsl = lsl
        self.update_stats_display(stats)
        self.draw_chart(data, stats)

    def draw_chart(self, data, stats):
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
        self.ax.clear(); self.ax.axis('off')
        self.ax.text(0.5, 0.5, "等待数据...", color='#555', ha='center', transform=self.ax.transAxes)
        self.canvas.draw()
        for k, l in self.stat_labels.items(): l.config(text="-", fg='white')
        self.lbl_proj_display.config(text="项目：未命名")
        self.current_data = None; self.current_stats = None

    def on_clear_tab1(self):
        self.inp_an_usl.delete(0, tk.END); self.inp_an_lsl.delete(0, tk.END)
        self.txt_data.delete("1.0", tk.END); self.reset_chart()

    def on_clear_tab2(self):
        for e in [self.inp_sim_usl, self.inp_sim_lsl, self.inp_sim_cpk, self.inp_sim_mean, self.inp_sim_cnt, self.inp_sim_prec]: e.delete(0, tk.END)
        self.txt_sim.delete("1.0", tk.END); self.reset_chart()

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
        about_text = "CPK统计分析工具 V2.6 (Zero Value Fix)\n\n修复:\n• 修复规格限为 0 时显示为空的问题\n• 图表增加横纵坐标轴标签\n• 关键指标高亮显示"
        DarkMessageBox(self.root, "关于软件", about_text, is_error=False)

    # ==========================================
    # 5. 核心优化：专业排版导出 (已修复 0 值显示 bug)
    # ==========================================
    def export_report(self):
        if not REPORTLAB_AVAILABLE:
            DarkMessageBox(self.root, "缺少依赖", "未安装 reportlab。\n请运行：pip install reportlab")
            return
        if self.current_stats is None:
            DarkMessageBox(self.root, "无数据", "请先进行分析或模拟。")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")],
            title="保存分析报告",
            initialfile=f"CPK_{self.project_name}_{datetime.now().strftime('%Y%m%d')}.pdf"
        )
        if not file_path: return

        try:
            # --- 1. 生成浅色高清图表 ---
            temp_img_path = None
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp_file:
                temp_img_path = tmp_file.name
            
            export_fig, export_ax = plt.subplots(figsize=(7.5, 4.8), dpi=300)
            export_fig.patch.set_facecolor('white')
            export_ax.set_facecolor('white')
            
            data = self.current_data; stats = self.current_stats
            mu, sigma = stats['Mean'], stats['StdDev']
            usl, lsl = stats['USL'], stats['LSL']
            
            export_ax.hist(data, bins=30, density=True, alpha=0.7, color='#3b82f6', edgecolor='white', linewidth=0.5)
            xmin, xmax = export_ax.get_xlim()
            base_span = 6 * sigma if sigma > 0 else 1.0
            plot_min = lsl - base_span*0.2 if lsl is not None else min(xmin, mu - 4*sigma)
            plot_max = usl + base_span*0.2 if usl is not None else max(xmax, mu + 4*sigma)
            x = np.linspace(plot_min, plot_max, 500)
            y = norm.pdf(x, mu, sigma)
            export_ax.plot(x, y, color='#dc2626', linewidth=2.5)
            export_ax.fill_between(x, y, alpha=0.2, color='#dc2626')
            
            ymax = max(y) * 1.15 if len(y)>0 else 1
            export_ax.set_ylim(0, ymax)
            
            if usl is not None:
                export_ax.axvline(usl, c='#ef4444', ls='--', lw=2)
                export_ax.text(usl, ymax*0.95, f"USL", c='#ef4444', ha='center', fontweight='bold', fontsize=9)
            if lsl is not None:
                export_ax.axvline(lsl, c='#ef4444', ls='--', lw=2)
                export_ax.text(lsl, ymax*0.95, f"LSL", c='#ef4444', ha='center', fontweight='bold', fontsize=9)
            export_ax.axvline(mu, c='#16a34a', ls='-', lw=2)
            
            export_ax.set_title(f"Project: {self.project_name}  |  Cpk = {stats['Cpk']:.3f}", fontsize=12, fontweight='bold', pad=10)
            export_ax.set_xlabel("Measurement Value", fontsize=10, color='black', labelpad=5)
            export_ax.set_ylabel("Probability Density", fontsize=10, color='black', labelpad=5)
            export_ax.tick_params(colors='black', labelsize=9)
            
            export_ax.grid(True, linestyle='--', alpha=0.3, color='gray')
            export_ax.spines['top'].set_visible(False); export_ax.spines['right'].set_visible(False)
            export_ax.spines['left'].set_color('black'); export_ax.spines['bottom'].set_color('black')
            
            export_fig.tight_layout()
            export_fig.savefig(temp_img_path, dpi=300, bbox_inches='tight', facecolor='white')
            plt.close(export_fig)
            
            # --- 2. 构建专业单页 PDF ---
            doc = SimpleDocTemplate(file_path, pagesize=A4, 
                                    rightMargin=1.5*cm, leftMargin=1.5*cm, 
                                    topMargin=1.5*cm, bottomMargin=1.5*cm)
            story = []
            styles = getSampleStyleSheet()
            
            # 字体注册
            font_name = "Helvetica"
            try:
                if os.name == 'nt':
                    path = r"C:\Windows\Fonts\simhei.ttf"
                    if os.path.exists(path):
                        pdfmetrics.registerFont(TTFont('SimHei', path)); font_name = 'SimHei'
                elif 'Darwin' in os.uname().sysname:
                    path = "/System/Library/Fonts/STHeiti Medium.ttc"
                    if os.path.exists(path):
                        pdfmetrics.registerFont(TTFont('STHeiti', path)); font_name = 'STHeiti'
            except: pass

            title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontName=font_name, fontSize=16, alignment=TA_CENTER, spaceAfter=5, textColor=colors.black)
            sub_style = ParagraphStyle('Sub', parent=styles['Normal'], fontName=font_name, fontSize=9, alignment=TA_CENTER, spaceAfter=12, textColor=colors.gray)
            head_style = ParagraphStyle('Head', parent=styles['Heading3'], fontName=font_name, fontSize=10, spaceBefore=8, spaceAfter=4, textColor=colors.darkblue)
            normal_style = ParagraphStyle('Norm', parent=styles['Normal'], fontName=font_name, fontSize=9, leading=12, textColor=colors.black)
            
            story.append(Paragraph("CPK 过程能力分析报表", title_style))
            story.append(Paragraph(f"项目名称：{self.project_name}  |  生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}", sub_style))
            
            # --- 3. 重点指标表格 (修复 0 值显示) ---
            s = self.current_stats
            
            # 【修复核心】使用辅助函数安全地格式化数值，区分 0 和 None
            def safe_fmt(val, fmt_str="{:.4f}"):
                if val is None:
                    return "-"
                return fmt_str.format(val)

            highlight_color = colors.Color(0.93, 0.96, 1.0)
            core_data = [
                [
                    # 修改前: if s['USL'] else "-"  (0 会被误判)
                    # 修改后: if s['USL'] is not None else "-" (0 正常显示)
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
            
            # --- 4. 图表 ---
            story.append(Paragraph("分布直方图", head_style))
            img = Image(temp_img_path, width=6.5*inch, height=4.2*inch)
            story.append(img)
            story.append(Spacer(1, 0.1*inch))
            
            # --- 5. 原始数据表格 ---
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
            
            if os.path.exists(temp_img_path): os.unlink(temp_img_path)
            DarkMessageBox(self.root, "导出成功", f"专业报告已保存至:\n{file_path}", is_error=False)
            
        except Exception as e:
            if temp_img_path and os.path.exists(temp_img_path): os.unlink(temp_img_path)
            DarkMessageBox(self.root, "导出失败", f"错误:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = CpkApp(root)
    root.mainloop()
