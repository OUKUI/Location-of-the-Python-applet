import tkinter as tk
from tkinter import ttk
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from scipy.stats import norm
import ctypes
import re
import json
import os

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
# 3. 核心计算 (支持单边规格)
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

        # 初始化变量
        cp = None
        cpk = None
        ppm = 0
        cpu = None
        cpl = None
        
        # --- 情况 A: 双边规格 (USL 和 LSL 都有) ---
        if usl is not None and lsl is not None:
            if lsl >= usl: return {"Error": "LSL 必须小于 USL"}
            cpu = (usl - mu) / (3 * sigma)
            cpl = (mu - lsl) / (3 * sigma)
            cpk = min(cpu, cpl)
            cp = (usl - lsl) / (6 * sigma)
            p_upper = 1 - norm.cdf(usl, mu, sigma)
            p_lower = norm.cdf(lsl, mu, sigma)
            ppm = (p_upper + p_lower) * 1_000_000
            
        # --- 情况 B: 只有上限 (USL) ---
        elif usl is not None and lsl is None:
            cpu = (usl - mu) / (3 * sigma) # 即 Cpu
            cpk = cpu
            cp = None # 单边无法计算 Cp
            p_upper = 1 - norm.cdf(usl, mu, sigma)
            ppm = p_upper * 1_000_000
            
        # --- 情况 C: 只有下限 (LSL) ---
        elif lsl is not None and usl is None:
            cpl = (mu - lsl) / (3 * sigma) # 即 Cpl
            cpk = cpl
            cp = None
            p_lower = norm.cdf(lsl, mu, sigma)
            ppm = p_lower * 1_000_000
            
        else:
            return {"Error": "请至少输入一个规格限 (USL 或 LSL)"}

        # CPK等级评估
        cpk_level = ""
        if cpk is not None:
            if cpk >= 1.67:
                cpk_level = "优秀 (Excellent)"
            elif cpk >= 1.33:
                cpk_level = "良好 (Good)"
            elif cpk >= 1.0:
                cpk_level = "一般 (Adequate)"
            elif cpk >= 0.67:
                cpk_level = "较差 (Poor)"
            else:
                cpk_level = "很差 (Inadequate)"

        return {
            "Count": n, "Mean": mu, "StdDev": sigma,
            "USL": usl, "LSL": lsl, "Cp": cp, "Cpk": cpk, "PPM": ppm,
            "CPU": cpu, "CPL": cpl, "CPK_LEVEL": cpk_level
        }

    @staticmethod
    def simulate(target_cpk, target_mean, usl, lsl, count, decimals):
        if target_cpk <= 0: return np.array([])
        
        sigma = 0
        # --- 反推 Sigma ---
        if usl is not None and lsl is not None:
            # 双边：取距离较近的一边
            d_u = usl - target_mean
            d_l = target_mean - lsl
            min_d = min(d_u, d_l)
            sigma = min_d / (3 * target_cpk)
            
        elif usl is not None:
            # 单边上限
            distance = usl - target_mean
            # 如果目标均值比 USL 还大，Cpk 逻辑上是负的，这里取绝对距离
            sigma = abs(distance) / (3 * target_cpk)
            
        elif lsl is not None:
            # 单边下限
            distance = target_mean - lsl
            sigma = abs(distance) / (3 * target_cpk)
        
        else:
            return None # 无规格

        raw_data = np.random.normal(target_mean, sigma, count)
        rounded_data = np.round(raw_data, decimals)
        return rounded_data

# ==========================================
# 4. 界面组件
# ==========================================
class DarkMessageBox(tk.Toplevel):
    def __init__(self, parent, title, message, is_error=True):
        super().__init__(parent)
        self.configure(bg=THEME['panel'])
        self.title(title)

        # 动态计算窗口大小
        base_width = 500
        base_height = 200

        # 根据消息长度调整窗口大小
        msg_lines = message.count('\n') + 1
        msg_length = len(message)

        # 计算需要的高度（考虑换行）
        estimated_lines = max(msg_lines, msg_length // 50)
        h = min(max(base_height, 150 + estimated_lines * 25), 600)  # 最大600px
        w = min(max(base_width, min(msg_length * 8, 700)), 800)  # 最大800px

        # 居中显示
        x = parent.winfo_x() + (parent.winfo_width()//2) - (w//2)
        y = parent.winfo_y() + (parent.winfo_height()//2) - (h//2)
        self.geometry(f"{int(w)}x{int(h)}+{int(x)}+{int(y)}")
        self.transient(parent); self.grab_set()

        color = THEME['danger'] if is_error else THEME['success']
        symbol = "❌ 计算失败" if is_error else "✅ 提示"

        # 标题
        tk.Label(self, text=symbol, font=("Segoe UI", 14, "bold"),
                bg=THEME['panel'], fg=color).pack(pady=(20, 10))

        # 消息内容 - 使用Text组件支持滚动和完整显示
        msg_frame = tk.Frame(self, bg=THEME['panel'])
        msg_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

        # 添加滚动条
        scrollbar = tk.Scrollbar(msg_frame, bg=THEME['input_bg'])
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 使用Text组件替代Label以支持更好的显示
        text_widget = tk.Text(msg_frame, font=("Microsoft YaHei", 10),
                             bg=THEME['input_bg'], fg='#ddd',
                             wrap=tk.WORD, height=max(3, min(estimated_lines, 15)),
                             relief=tk.FLAT, padx=10, pady=10,
                             yscrollcommand=scrollbar.set)
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=text_widget.yview)

        # 插入消息并设置为只读
        text_widget.insert('1.0', message)
        text_widget.config(state=tk.DISABLED)

        # 确定按钮
        btn = tk.Button(self, text="确定 (Enter)", bg=THEME['input_bg'], fg='white',
                       relief=tk.FLAT, command=self.destroy, width=15,
                       font=("Microsoft YaHei", 10))
        btn.pack(pady=15)

        # 绑定回车键关闭
        self.bind('<Return>', lambda e: self.destroy())
        self.bind('<KP_Enter>', lambda e: self.destroy())  # 小键盘回车

        # 绑定ESC键关闭
        self.bind('<Escape>', lambda e: self.destroy())

        # 强制设置焦点到窗口本身，确保打开后立即可以按回车关闭
        self.focus_force()

        self.wait_window()

class CpkApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CPK Tool Pro - One Sided Support")
        # 增加整体宽度以适应更宽的右侧
        self.root.geometry(f"{int(1700)}x{int(820)}") 
        self.root.configure(bg=THEME['bg'])
        self.setup_ui()

    def setup_ui(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TNotebook", background=THEME['bg'], borderwidth=0)
        style.configure("TNotebook.Tab", background=THEME['panel'], foreground=THEME['fg'], padding=[25, 12], font=('Microsoft YaHei', 11))
        style.map("TNotebook.Tab", background=[("selected", THEME['accent'])], foreground=[("selected", 'white')])

        main = tk.Frame(self.root, bg=THEME['bg'])
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # === 布局调整：左侧 460，右侧 350，中间自动填充 ===
        left = tk.Frame(main, bg=THEME['panel'], width=460)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 5))
        left.pack_propagate(False)

        right = tk.Frame(main, bg=THEME['panel'], width=350)  # 增加右侧宽度
        right.pack(side=tk.RIGHT, fill=tk.Y, padx=(5, 0))
        right.pack_propagate(False)
        
        center = tk.Frame(main, bg=THEME['bg'])
        center.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.init_left_panel(left)
        self.init_stats_panel(right)
        self.init_chart_panel(center)

        # 添加右下角的"关于"链接
        self.add_about_link()

    def init_left_panel(self, parent):
        tk.Label(parent, text="⚙️ 控制台", bg=THEME['panel'], fg='white', font=("Segoe UI", 16, "bold")).pack(pady=15)
        
        nb = ttk.Notebook(parent)
        nb.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
        
        t1 = ttk.Frame(nb); nb.add(t1, text=' 数据分析 ')
        t2 = ttk.Frame(nb); nb.add(t2, text=' 模拟生成 ')
        
        self.setup_tab1(t1)
        self.setup_tab2(t2)

    def init_stats_panel(self, parent):
        tk.Label(parent, text="📊 结果汇总", bg=THEME['panel'], fg=THEME['accent'], font=("Segoe UI", 14, "bold")).pack(pady=20)
        self.stat_labels = {}
        
        # 格式配置：如果 fmt 为 None，说明需要特殊处理（比如 Cp 可能为 N/A）
        fields = [
            ("Count", "样本数 N"), ("Mean", "均值 Mean"), ("StdDev", "标准差 Std"),
            (None, None),
            ("USL", "规格上限"), ("LSL", "规格下限"),
            (None, None),
            ("Cp", "Cp (精密度)"), ("Cpk", "Cpk (能力)"), ("CPK_LEVEL", "Cpk等级"),
            (None, None),
            ("CPU", "CPU (上限能力)"), ("CPL", "CPL (下限能力)"), ("PPM", "不良率 PPM")
        ]
        
        # 创建表格框架，设置更大的宽度
        tbl = tk.Frame(parent, bg=THEME['panel'])
        tbl.pack(fill=tk.X, padx=20)
        
        for i, (key, label) in enumerate(fields):
            if key is None:
                tk.Frame(tbl, bg=THEME['border'], height=1).grid(row=i, column=0, columnspan=2, sticky='ew', pady=8)
            else:
                # 左侧标签，设置固定宽度
                tk.Label(tbl, text=label, bg=THEME['panel'], fg='#888', anchor='w', width=15).grid(row=i, column=0, sticky='w', padx=(0, 5))
                # 右侧数值，设置更大的宽度
                val = tk.Label(tbl, text="-", bg=THEME['panel'], fg='white', font=("Consolas", 11, "bold"), anchor='w', width=20)
                val.grid(row=i, column=1, sticky='ew')
                self.stat_labels[key] = val
        # 配置列权重以允许扩展
        tbl.columnconfigure(1, weight=1)

    def update_stats_display(self, stats):
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
                
                # 设置颜色
                if key == "CPK_LEVEL":
                    level_colors = {
                        "优秀 (Excellent)": THEME['success'],
                        "良好 (Good)": "#aaff00",
                        "一般 (Adequate)": THEME['warning'],
                        "较差 (Poor)": "#ffaa00",
                        "很差 (Inadequate)": THEME['danger']
                    }
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

    def create_input_with_unit(self, parent, label, unit, row, default=""):
        frame = tk.Frame(parent, bg=THEME['panel'])
        frame.grid(row=row, column=0, columnspan=2, sticky='ew', pady=8)
        tk.Label(frame, text=label, bg=THEME['panel'], fg=THEME['fg']).pack(side=tk.LEFT)
        e = tk.Entry(frame, bg=THEME['input_bg'], fg='white', insertbackground='white', relief=tk.FLAT, font=('Consolas', 11))
        if default: e.insert(0, default)
        e.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))
        tk.Label(frame, text=unit, bg=THEME['panel'], fg=THEME['fg']).pack(side=tk.LEFT, padx=(5, 0))
        return e

    # === Tab 1: 分析 ===
    def setup_tab1(self, f):
        f = tk.Frame(f, bg=THEME['panel'], padx=20, pady=20)
        f.pack(fill=tk.BOTH, expand=True)
        f.columnconfigure(1, weight=1)
        
        tk.Label(f, text="规格设置 (可只填一项):", bg=THEME['panel'], fg='#888', font=('Microsoft YaHei', 9)).grid(row=0, column=0, columnspan=2, sticky='w')
        self.inp_an_usl = self.create_input(f, "上限 (USL)", 1)
        self.inp_an_lsl = self.create_input(f, "下限 (LSL)", 2)
        
        tk.Label(f, text="测量数据 (粘贴):", bg=THEME['panel'], fg=THEME['fg']).grid(row=3, column=0, sticky='w', pady=(20, 5))
        # 宽度扩大，方便看数据
        self.txt_data = tk.Text(f, bg=THEME['input_bg'], fg='white', height=18, relief=tk.FLAT, font=('Consolas', 10), width=45)
        self.txt_data.grid(row=4, column=0, columnspan=2, sticky='nsew')
        
        self.create_btn_bar(f, 5, self.on_analyze, self.on_clear_tab1, "开始分析")

    # === Tab 2: 模拟 ===
    def setup_tab2(self, f):
        f = tk.Frame(f, bg=THEME['panel'], padx=20, pady=20)
        f.pack(fill=tk.BOTH, expand=True)
        f.columnconfigure(1, weight=1)
        
        tk.Label(f, text="规格设置 (可只填一项):", bg=THEME['panel'], fg='#888', font=('Microsoft YaHei', 9)).grid(row=0, column=0, columnspan=2, sticky='w')
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

    # === 逻辑处理 ===
    def get_val(self, entry, is_int=False, allow_empty=False):
        val_str = entry.get().strip()
        if not val_str:
            return None if allow_empty else False # False标记为错误
        try:
            val = float(val_str)
            return int(val) if is_int else val
        except:
            return False # 格式错误

    def on_analyze(self):
        # 允许 USL 或 LSL 为空，但不能同时为空
        usl = self.get_val(self.inp_an_usl, allow_empty=True)
        lsl = self.get_val(self.inp_an_lsl, allow_empty=True)
        
        if usl is False or lsl is False: # 格式错误
            DarkMessageBox(self.root, "输入错误", "规格值必须是数字")
            return
        if usl is None and lsl is None:
            DarkMessageBox(self.root, "缺失规格", "请至少输入一个规格限 (USL 或 LSL)")
            return

        raw = self.txt_data.get("1.0", tk.END)
        nums = re.findall(r"[-+]?\d*\.\d+|\d+", raw)
        try:
            data = np.array([float(x) for x in nums])
            if len(data) < 2: raise ValueError
        except:
            DarkMessageBox(self.root, "数据错误", "请检查输入数据")
            return
            
        self.process_result(data, usl, lsl)

    def on_simulate(self):
        usl = self.get_val(self.inp_sim_usl, allow_empty=True)
        lsl = self.get_val(self.inp_sim_lsl, allow_empty=True)
        cpk = self.get_val(self.inp_sim_cpk)
        mean = self.get_val(self.inp_sim_mean)
        cnt = self.get_val(self.inp_sim_cnt, True)
        prec = self.get_val(self.inp_sim_prec, True)
        
        if any(x is False for x in [usl, lsl, cpk, mean, cnt, prec]): 
            DarkMessageBox(self.root, "输入错误", "请检查数值格式")
            return
        if usl is None and lsl is None:
            DarkMessageBox(self.root, "缺失规格", "请至少输入一个规格限")
            return
            
        data = CpkCalculator.simulate(cpk, mean, usl, lsl, cnt, max(0, min(prec, 10)))
        
        if data is None:
             DarkMessageBox(self.root, "错误", "无法根据当前条件生成数据")
             return

        fmt = f"{{:.{prec}f}}"
        self.txt_sim.delete("1.0", tk.END)
        self.txt_sim.insert(tk.END, "\n".join([fmt.format(x) for x in data]))
        
        self.process_result(data, usl, lsl)

    def process_result(self, data, usl, lsl):
        stats = CpkCalculator.calculate(data, usl, lsl)
        if "Error" in stats:
            DarkMessageBox(self.root, "计算错误", stats["Error"])
            return
        self.update_stats_display(stats)
        self.draw_chart(data, stats)

    def draw_chart(self, data, stats):
        self.ax.clear()
        self.ax.axis('on')
        self.ax.set_facecolor(THEME['bg'])
        
        mu, sigma = stats['Mean'], stats['StdDev']
        usl, lsl = stats['USL'], stats['LSL']
        
        # 绘图
        self.ax.hist(data, bins=30, density=True, alpha=0.5, color=THEME['accent'], edgecolor=None)
        
        # 确定X轴范围
        xmin, xmax = self.ax.get_xlim()
        # 根据存在的规格动态调整范围
        base_span = 6 * sigma if sigma > 0 else 1.0
        
        plot_min = lsl - base_span*0.2 if lsl is not None else min(xmin, mu - 4*sigma)
        plot_max = usl + base_span*0.2 if usl is not None else max(xmax, mu + 4*sigma)
        
        x = np.linspace(plot_min, plot_max, 500)
        y = norm.pdf(x, mu, sigma)
        
        self.ax.plot(x, y, color=THEME['success'], linewidth=2)
        self.ax.fill_between(x, y, alpha=0.15, color=THEME['success'])
        
        # 规格线
        ymax = max(y) * 1.2 if len(y)>0 else 1
        self.ax.set_ylim(0, ymax)
        
        if usl is not None:
            self.ax.axvline(usl, c=THEME['danger'], ls='--', lw=1.5)
            self.ax.text(usl, ymax*0.95, "USL", c=THEME['danger'], ha='center')
            
        if lsl is not None:
            self.ax.axvline(lsl, c=THEME['danger'], ls='--', lw=1.5)
            self.ax.text(lsl, ymax*0.95, "LSL", c=THEME['danger'], ha='center')
        
        # 添加均值线
        self.ax.axvline(mu, c=THEME['warning'], ls='-', lw=1.5, alpha=0.7)
        self.ax.text(mu, ymax*0.85, f"Mean: {mu:.3f}", c=THEME['warning'], ha='center')
        
        # 去除边框
        self.ax.spines['top'].set_visible(False)
        self.ax.spines['right'].set_visible(False)
        self.ax.spines['left'].set_color(THEME['border'])
        self.ax.spines['bottom'].set_color(THEME['border'])
        self.ax.tick_params(colors='#888')
        
        self.canvas.draw()

    def reset_chart(self):
        self.ax.clear(); self.ax.axis('off')
        self.ax.text(0.5, 0.5, "等待数据...", color='#555', ha='center', transform=self.ax.transAxes)
        self.canvas.draw()
        for k, l in self.stat_labels.items(): l.config(text="-", fg='white')

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
        """在窗口右下角添加关于链接"""
        about_label = tk.Label(self.root, text="关于软件",
                              bg=THEME['bg'], fg=THEME['accent'],
                              font=("Microsoft YaHei", 9, "underline"),
                              cursor="hand2")
        about_label.place(relx=1.0, rely=1.0, anchor='se', x=-15, y=-10)
        about_label.bind("<Button-1>", lambda e: self.show_about())

        # 鼠标悬停效果
        about_label.bind("<Enter>", lambda e: about_label.config(fg=THEME['success']))
        about_label.bind("<Leave>", lambda e: about_label.config(fg=THEME['accent']))

    def show_about(self):
        """显示使用条款和版权声明"""
        about_text = """
CPK统计分析工具 V2.0

功能特性：
• 支持双边和单边规格限制的CPK计算
• 提供数据模拟生成功能
• 深色主题界面，美观易用
• 实时图表显示分布情况
• 支持多种统计指标显示

使用说明：
• 在数据分析标签页中输入规格限和测量数据
• 在模拟生成标签页中设置目标参数生成数据
• 结果面板显示详细的统计信息

作者：CPK分析团队
版本：2.0
日期：2024年
        """
        DarkMessageBox(self.root, "关于软件", about_text, is_error=False)

if __name__ == "__main__":
    root = tk.Tk()
    app = CpkApp(root)
    root.mainloop()
