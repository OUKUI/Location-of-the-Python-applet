import tkinter as tk
from tkinter import ttk, filedialog, simpledialog
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from scipy.stats import norm, skew, kurtosis
import ctypes
import re
import os
import tempfile
from datetime import datetime
import traceback

# 尝试导入 reportlab
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import (
        SimpleDocTemplate, BaseDocTemplate, Frame, PageTemplate, FrameBreak,
        NextPageTemplate, Table, TableStyle, Paragraph, Spacer, Image,
        PageBreak, KeepInFrame
    )
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
except Exception:
    ScaleFactor = 1.0

# ==========================================
# 2. 现代化设计系统
# ==========================================
plt.style.use('dark_background')
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['figure.dpi'] = 100 * ScaleFactor

THEME = {
    # surfaces
    'bg': '#0b0f14',
    'panel': '#12181f',
    'card': '#171e27',
    'card_hover': '#1c2530',
    'input_bg': '#0e141b',
    'elevated': '#1a222d',
    # borders / dividers
    'border': '#273140',
    'border_soft': '#1e2834',
    'focus': '#4f7cff',
    # text
    'fg': '#e8eef7',
    'fg_muted': '#8b97a8',
    'fg_dim': '#5c6a7a',
    # brand
    'accent': '#4f7cff',
    'accent_hover': '#6b93ff',
    'accent_soft': '#1a2744',
    'accent_text': '#dbe6ff',
    # semantic
    'success': '#3dd68c',
    'success_soft': '#123526',
    'danger': '#ff6b6b',
    'danger_soft': '#3a1717',
    'warning': '#f5c542',
    'warning_soft': '#3a2e0f',
    'info': '#56c4ff',
    # chart
    'chart_fill': '#4f7cff',
    'chart_curve': '#3dd68c',
    'chart_spec': '#ff6b6b',
    'chart_mean': '#f5c542',
}

FONT_UI = 'Microsoft YaHei UI'
FONT_MONO = 'Consolas'


# ==========================================
# 2.1 现代 UI 组件
# ==========================================
class ModernButton(tk.Frame):
    """扁平圆角感按钮（hover / active / disabled）"""

    def __init__(self, parent, text='', command=None, style='primary',
                 width=None, height=36, font=None, **kwargs):
        super().__init__(parent, bg=parent.cget('bg') if hasattr(parent, 'cget') else THEME['panel'], **kwargs)
        self.command = command
        self.style = style
        self._enabled = True
        self._font = font or (FONT_UI, 10, 'bold')
        self._height = height
        self._width = width
        self._text = text
        self._colors = self._palette(style)

        self.configure(bg=self._colors['bg'], highlightthickness=0, bd=0)
        self.btn = tk.Label(
            self, text=text, bg=self._colors['bg'], fg=self._colors['fg'],
            font=self._font, cursor='hand2', padx=16, pady=8
        )
        self.btn.pack(fill=tk.BOTH, expand=True)
        if width:
            self.btn.configure(width=width)
        # height 参数生效：固定按钮总高，内部 Label 自动填满
        if height:
            self.configure(height=height)
            self.pack_propagate(False)
            self.configure(width=self.btn.winfo_reqwidth())

        for w in (self, self.btn):
            w.bind('<Enter>', self._on_enter)
            w.bind('<Leave>', self._on_leave)
            w.bind('<Button-1>', self._on_press)
            w.bind('<ButtonRelease-1>', self._on_release)

    def _palette(self, style):
        palettes = {
            'primary': {
                'bg': THEME['accent'], 'fg': '#ffffff',
                'hover': THEME['accent_hover'], 'active': '#3d68e8',
                'disabled_bg': '#2a3545', 'disabled_fg': THEME['fg_dim'],
            },
            'success': {
                'bg': '#1f9d63', 'fg': '#ffffff',
                'hover': '#24b572', 'active': '#178552',
                'disabled_bg': '#2a3545', 'disabled_fg': THEME['fg_dim'],
            },
            'danger': {
                'bg': '#c94444', 'fg': '#ffffff',
                'hover': '#e05555', 'active': '#a83636',
                'disabled_bg': '#2a3545', 'disabled_fg': THEME['fg_dim'],
            },
            'ghost': {
                'bg': THEME['elevated'], 'fg': THEME['fg'],
                'hover': THEME['card_hover'], 'active': THEME['border'],
                'disabled_bg': THEME['panel'], 'disabled_fg': THEME['fg_dim'],
            },
            'soft': {
                'bg': THEME['accent_soft'], 'fg': THEME['accent_text'],
                'hover': '#243456', 'active': '#1a2744',
                'disabled_bg': THEME['panel'], 'disabled_fg': THEME['fg_dim'],
            },
        }
        return palettes.get(style, palettes['primary'])

    def _paint(self, bg, fg=None):
        self.configure(bg=bg)
        self.btn.configure(bg=bg, fg=fg if fg is not None else self.btn.cget('fg'))

    def _on_enter(self, _e=None):
        if self._enabled:
            self._paint(self._colors['hover'], self._colors['fg'])

    def _on_leave(self, _e=None):
        if self._enabled:
            self._paint(self._colors['bg'], self._colors['fg'])

    def _on_press(self, _e=None):
        if self._enabled:
            self._paint(self._colors['active'], self._colors['fg'])

    def _on_release(self, _e=None):
        if self._enabled:
            self._paint(self._colors['hover'], self._colors['fg'])
            if self.command:
                self.command()

    def config(self, **kwargs):
        self.configure(**kwargs)

    def configure(self, cnf=None, **kwargs):
        if cnf and isinstance(cnf, dict):
            kwargs = {**cnf, **kwargs}
        if 'text' in kwargs:
            self._text = kwargs.pop('text')
            self.btn.configure(text=self._text)
        if 'command' in kwargs:
            self.command = kwargs.pop('command')
        if 'state' in kwargs:
            state = kwargs.pop('state')
            self._enabled = state != tk.DISABLED and str(state).lower() != 'disabled'
            if self._enabled:
                self.btn.configure(cursor='hand2')
                self._paint(self._colors['bg'], self._colors['fg'])
            else:
                self.btn.configure(cursor='arrow')
                self._paint(self._colors['disabled_bg'], self._colors['disabled_fg'])
        if kwargs:
            super().configure(**kwargs)

    def cget(self, key):
        if key == 'text':
            return self._text
        return super().cget(key)


class Card(tk.Frame):
    """带边框的卡片容器"""

    def __init__(self, parent, pad=12, **kwargs):
        bg = kwargs.pop('bg', THEME['card'])
        super().__init__(
            parent, bg=bg, highlightbackground=THEME['border'],
            highlightthickness=1, bd=0, **kwargs
        )
        self.inner = tk.Frame(self, bg=bg)
        self.inner.pack(fill=tk.BOTH, expand=True, padx=pad, pady=pad)


class ModernEntry(tk.Frame):
    """带标签的输入框"""

    def __init__(self, parent, label='', default='', placeholder='', **kwargs):
        super().__init__(parent, bg=parent.cget('bg'), **kwargs)
        if label:
            tk.Label(
                self, text=label, bg=self.cget('bg'), fg=THEME['fg_muted'],
                font=(FONT_UI, 9), anchor='w'
            ).pack(fill=tk.X, pady=(0, 4))

        self.box = tk.Frame(self, bg=THEME['border'], bd=0)
        self.box.pack(fill=tk.X)
        self.entry = tk.Entry(
            self.box, bg=THEME['input_bg'], fg=THEME['fg'],
            insertbackground=THEME['accent'], relief=tk.FLAT,
            font=(FONT_MONO, 11), bd=0
        )
        self.entry.pack(fill=tk.X, ipady=8, padx=1, pady=1)
        if default:
            self.entry.insert(0, default)

        self.entry.bind('<FocusIn>', self._focus_in)
        self.entry.bind('<FocusOut>', self._focus_out)

    def _focus_in(self, _e=None):
        self.box.configure(bg=THEME['focus'])

    def _focus_out(self, _e=None):
        self.box.configure(bg=THEME['border'])

    def get(self):
        return self.entry.get()

    def delete(self, *args):
        self.entry.delete(*args)

    def insert(self, *args):
        self.entry.insert(*args)

    def bind(self, *args, **kwargs):
        return self.entry.bind(*args, **kwargs)


class SectionTitle(tk.Frame):
    def __init__(self, parent, text='', icon='', **kwargs):
        super().__init__(parent, bg=parent.cget('bg'), **kwargs)
        row = tk.Frame(self, bg=self.cget('bg'))
        row.pack(fill=tk.X)
        if icon:
            tk.Label(row, text=icon, bg=self.cget('bg'), fg=THEME['accent'],
                     font=(FONT_UI, 12)).pack(side=tk.LEFT, padx=(0, 6))
        tk.Label(row, text=text, bg=self.cget('bg'), fg=THEME['fg'],
                 font=(FONT_UI, 12, 'bold')).pack(side=tk.LEFT)
        tk.Frame(self, bg=THEME['border_soft'], height=1).pack(fill=tk.X, pady=(8, 0))


class MetricChip(tk.Frame):
    """指标小卡片"""

    def __init__(self, parent, label='', value='-', **kwargs):
        super().__init__(
            parent, bg=THEME['elevated'], highlightbackground=THEME['border_soft'],
            highlightthickness=1, bd=0, **kwargs
        )
        tk.Label(
            self, text=label, bg=THEME['elevated'], fg=THEME['fg_muted'],
            font=(FONT_UI, 8), anchor='w'
        ).pack(fill=tk.X, padx=10, pady=(8, 0))
        self.value_lbl = tk.Label(
            self, text=value, bg=THEME['elevated'], fg=THEME['fg'],
            font=(FONT_MONO, 11, 'bold'), anchor='w'
        )
        self.value_lbl.pack(fill=tk.X, padx=10, pady=(2, 8))

    def set_value(self, text, color=None):
        self.value_lbl.configure(text=text, fg=color or THEME['fg'])


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
        data_min = np.min(data)
        data_max = np.max(data)
        data_range = data_max - data_min
        data_median = np.median(data)

        data_skew = skew(data)
        data_kurt = kurtosis(data)
        cv = (sigma / abs(mu)) * 100 if mu != 0 else 0.0

        if sigma <= 1e-9:
            return {"Error": "标准差为 0，无法计算"}

        cp = None
        cpk = None
        ppm = 0
        cpu = None
        cpl = None

        out_of_spec_count = 0
        if usl is not None:
            out_of_spec_count += np.sum(data > usl)
        if lsl is not None:
            out_of_spec_count += np.sum(data < lsl)

        if usl is not None and lsl is not None:
            if lsl >= usl:
                return {"Error": "LSL 必须小于 USL"}
            cpu = (usl - mu) / (3 * sigma)
            cpl = (mu - lsl) / (3 * sigma)
            cpk = min(cpu, cpl)
            cp = (usl - lsl) / (6 * sigma)

            p_upper = 1 - norm.cdf(usl, mu, sigma)
            p_lower = norm.cdf(lsl, mu, sigma)
            ppm = (p_upper + p_lower) * 1_000_000

            sigma_level = 3 * cpk

        elif usl is not None:
            cpu = (usl - mu) / (3 * sigma)
            cpk = cpu
            cp = None
            ppm = (1 - norm.cdf(usl, mu, sigma)) * 1_000_000
            sigma_level = 3 * cpk
        elif lsl is not None:
            cpl = (mu - lsl) / (3 * sigma)
            cpk = cpl
            cp = None
            ppm = norm.cdf(lsl, mu, sigma) * 1_000_000
            sigma_level = 3 * cpk
        else:
            return {"Error": "请至少输入一个规格限"}

        cpk_level = ""
        if cpk is not None:
            if cpk >= 1.67:
                cpk_level = "优秀"
            elif cpk >= 1.33:
                cpk_level = "良好"
            elif cpk >= 1.0:
                cpk_level = "一般"
            elif cpk >= 0.67:
                cpk_level = "较差"
            else:
                cpk_level = "很差"

        return {
            "Count": n,
            "Mean": mu,
            "StdDev": sigma,
            "Max": data_max,
            "Min": data_min,
            "Range": data_range,
            "Median": data_median,
            "Skewness": data_skew,
            "Kurtosis": data_kurt,
            "CV": cv,
            "USL": usl, "LSL": lsl,
            "Cp": cp, "Cpk": cpk,
            "PPM": ppm,
            "CPU": cpu, "CPL": cpl,
            "CPK_LEVEL": cpk_level,
            "OutOfSpecCount": int(out_of_spec_count),
            "SigmaLevel": sigma_level if 'sigma_level' in locals() else 0
        }

    @staticmethod
    def simulate(target_cpk, target_mean, usl, lsl, count, decimals):
        if target_cpk <= 0:
            return np.array([])
        sigma = 0
        if usl is not None and lsl is not None:
            sigma = min(usl - target_mean, target_mean - lsl) / (3 * target_cpk)
        elif usl is not None:
            sigma = abs(usl - target_mean) / (3 * target_cpk)
        elif lsl is not None:
            sigma = abs(target_mean - lsl) / (3 * target_cpk)
        else:
            return None
        return np.round(np.random.normal(target_mean, sigma, count), decimals)


# ==========================================
# 4. 消息对话框
# ==========================================
def show_msg(parent, title, msg, is_error=True):
    dlg = tk.Toplevel(parent)
    dlg.title(title)
    dlg.configure(bg=THEME['panel'])
    dlg.resizable(False, False)
    dlg.transient(parent)
    dlg.grab_set()
    dlg.configure(highlightbackground=THEME['border'], highlightthickness=1)

    color = THEME['danger'] if is_error else THEME['success']
    soft = THEME['danger_soft'] if is_error else THEME['success_soft']
    icon = "!" if is_error else "✓"

    top = tk.Frame(dlg, bg=THEME['panel'])
    top.pack(fill=tk.X, padx=24, pady=(22, 10))

    badge = tk.Label(
        top, text=icon, bg=soft, fg=color, font=(FONT_UI, 14, 'bold'),
        width=3, height=1
    )
    badge.pack(side=tk.LEFT, padx=(0, 12))
    tk.Label(
        top, text=title, font=(FONT_UI, 13, 'bold'),
        bg=THEME['panel'], fg=THEME['fg']
    ).pack(side=tk.LEFT)

    msg_frame = tk.Frame(dlg, bg=THEME['panel'])
    msg_frame.pack(fill=tk.BOTH, expand=True, padx=28, pady=(0, 18))
    tk.Label(
        msg_frame, text=msg, font=(FONT_UI, 10), bg=THEME['panel'],
        fg=THEME['fg_muted'], wraplength=440, justify=tk.LEFT
    ).pack(anchor='w')

    btn_frame = tk.Frame(dlg, bg=THEME['card'])
    btn_frame.pack(fill=tk.X)
    inner = tk.Frame(btn_frame, bg=THEME['card'])
    inner.pack(pady=12)
    ModernButton(inner, text="确定", command=dlg.destroy, style='primary').pack()

    dlg.update_idletasks()
    pw, ph = parent.winfo_width(), parent.winfo_height()
    px = parent.winfo_x() + max(0, (pw - dlg.winfo_width()) // 2)
    py = parent.winfo_y() + max(0, (ph - dlg.winfo_height()) // 2)
    dlg.geometry(f"+{px}+{py}")

    dlg.bind('<Return>', lambda e: dlg.destroy())
    dlg.bind('<Escape>', lambda e: dlg.destroy())
    dlg.focus_force()
    dlg.wait_window()


class CpkApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CPK 统计分析工具 Pro")
        self.root.geometry("1600x900")
        self.root.state('zoomed')
        self.root.configure(bg=THEME['bg'])
        self.root.minsize(1200, 720)
        self._set_dark_titlebar()

        self.app_dir = os.path.dirname(os.path.abspath(__file__))

        self.current_data = None
        self.current_stats = None
        self.current_usl = None
        self.current_lsl = None
        self.project_name = ""

        self.excel_projects = []
        self.current_excel_index = -1

        self.setup_styles()
        self.setup_ui()
        self.root.bind('<Control-Shift-D>', self._show_debug_prompt)

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')

        style.configure(
            "Modern.TNotebook",
            background=THEME['panel'],
            borderwidth=0,
            tabmargins=[0, 0, 0, 0]
        )
        style.configure(
            "Modern.TNotebook.Tab",
            background=THEME['elevated'],
            foreground=THEME['fg_muted'],
            padding=[18, 10],
            font=(FONT_UI, 10),
            borderwidth=0,
            lightcolor=THEME['panel'],
            darkcolor=THEME['panel'],
            focuscolor=THEME['panel'],
        )
        style.map(
            "Modern.TNotebook.Tab",
            background=[("selected", THEME['accent_soft']), ("active", THEME['card_hover'])],
            foreground=[("selected", THEME['accent_text']), ("active", THEME['fg'])],
        )

        style.configure(
            "Modern.Treeview",
            background=THEME['input_bg'],
            foreground=THEME['fg'],
            fieldbackground=THEME['input_bg'],
            rowheight=32,
            font=(FONT_UI, 9),
            borderwidth=0,
            relief='flat',
        )
        style.map(
            "Modern.Treeview",
            background=[('selected', THEME['accent_soft'])],
            foreground=[('selected', THEME['accent_text'])],
        )
        style.configure(
            "Modern.Treeview.Heading",
            background=THEME['elevated'],
            foreground=THEME['fg_muted'],
            font=(FONT_UI, 9, 'bold'),
            relief='flat',
            borderwidth=0,
            padding=6,
        )
        style.map(
            "Modern.Treeview.Heading",
            background=[('active', THEME['card_hover'])],
            foreground=[('active', THEME['fg'])],
        )
        style.configure(
            "Modern.Vertical.TScrollbar",
            background=THEME['elevated'],
            troughcolor=THEME['input_bg'],
            bordercolor=THEME['input_bg'],
            arrowcolor=THEME['fg_muted'],
            relief='flat',
        )
        style.map(
            "Modern.Vertical.TScrollbar",
            background=[('active', THEME['border']), ('pressed', THEME['accent'])],
        )

    def setup_ui(self):
        # 顶部栏
        self._build_topbar()

        main = tk.Frame(self.root, bg=THEME['bg'])
        main.pack(fill=tk.BOTH, expand=True, padx=14, pady=(0, 14))

        # 三栏布局
        left = tk.Frame(main, bg=THEME['panel'], width=520, highlightbackground=THEME['border'], highlightthickness=1)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        left.pack_propagate(False)

        right = tk.Frame(main, bg=THEME['panel'], width=360, highlightbackground=THEME['border'], highlightthickness=1)
        right.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0))
        right.pack_propagate(False)

        center = tk.Frame(main, bg=THEME['panel'], highlightbackground=THEME['border'], highlightthickness=1)
        center.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.init_left_panel(left)
        self.init_stats_panel(right)
        self.init_chart_panel(center)

    def _build_topbar(self):
        bar = tk.Frame(self.root, bg=THEME['panel'], height=56, highlightbackground=THEME['border'], highlightthickness=1)
        bar.pack(fill=tk.X, padx=14, pady=14)
        bar.pack_propagate(False)

        left = tk.Frame(bar, bg=THEME['panel'])
        left.pack(side=tk.LEFT, fill=tk.Y, padx=16)

        logo = tk.Label(
            left, text="CPK", bg=THEME['accent'], fg='#ffffff',
            font=(FONT_UI, 11, 'bold'), padx=10, pady=4
        )
        logo.pack(side=tk.LEFT, pady=12)
        titles = tk.Frame(left, bg=THEME['panel'])
        titles.pack(side=tk.LEFT, padx=12, pady=10)
        tk.Label(
            titles, text="过程能力分析", bg=THEME['panel'], fg=THEME['fg'],
            font=(FONT_UI, 13, 'bold')
        ).pack(anchor='w')
        tk.Label(
            titles, text="Statistical Process Capability · Pro", bg=THEME['panel'],
            fg=THEME['fg_dim'], font=(FONT_UI, 8)
        ).pack(anchor='w')

        right = tk.Frame(bar, bg=THEME['panel'])
        right.pack(side=tk.RIGHT, padx=16)
        about = tk.Label(
            right, text="关于", bg=THEME['elevated'], fg=THEME['fg_muted'],
            font=(FONT_UI, 9), cursor='hand2', padx=12, pady=6
        )
        about.pack(side=tk.RIGHT, pady=12)
        about.bind('<Button-1>', lambda e: self.show_about())
        about.bind('<Enter>', lambda e: about.configure(fg=THEME['fg'], bg=THEME['card_hover']))
        about.bind('<Leave>', lambda e: about.configure(fg=THEME['fg_muted'], bg=THEME['elevated']))

        self._status_chip = tk.Label(
            right, text="就绪", bg=THEME['success_soft'], fg=THEME['success'],
            font=(FONT_UI, 9), padx=12, pady=6
        )
        self._status_chip.pack(side=tk.RIGHT, padx=(0, 10), pady=12)

    def set_status(self, text, kind='ok'):
        colors = {
            'ok': (THEME['success_soft'], THEME['success']),
            'warn': (THEME['warning_soft'], THEME['warning']),
            'err': (THEME['danger_soft'], THEME['danger']),
            'info': (THEME['accent_soft'], THEME['accent_text']),
        }
        bg, fg = colors.get(kind, colors['ok'])
        self._status_chip.configure(text=text, bg=bg, fg=fg)

    def init_left_panel(self, parent):
        header = tk.Frame(parent, bg=THEME['panel'])
        header.pack(fill=tk.X, padx=16, pady=(16, 8))
        tk.Label(
            header, text="控制台", bg=THEME['panel'], fg=THEME['fg'],
            font=(FONT_UI, 14, 'bold')
        ).pack(anchor='w')
        tk.Label(
            header, text="配置项目、输入数据或导入 Excel", bg=THEME['panel'],
            fg=THEME['fg_dim'], font=(FONT_UI, 9)
        ).pack(anchor='w', pady=(2, 0))

        # 项目名称卡片
        proj_card = Card(parent, pad=12)
        proj_card.pack(fill=tk.X, padx=14, pady=(4, 10))
        tk.Label(
            proj_card.inner, text="项目名称", bg=THEME['card'], fg=THEME['fg_muted'],
            font=(FONT_UI, 9)
        ).pack(anchor='w')
        box = tk.Frame(proj_card.inner, bg=THEME['border'])
        box.pack(fill=tk.X, pady=(6, 0))
        self.inp_project = tk.Entry(
            box, bg=THEME['input_bg'], fg=THEME['fg'], insertbackground=THEME['accent'],
            relief=tk.FLAT, font=(FONT_UI, 11), bd=0
        )
        self.inp_project.pack(fill=tk.X, ipady=8, padx=1, pady=1)
        self.inp_project.insert(0, "未命名项目")
        self.inp_project.bind('<FocusIn>', lambda e: box.configure(bg=THEME['focus']))
        self.inp_project.bind('<FocusOut>', lambda e: box.configure(bg=THEME['border']))

        # Notebook
        nb_wrap = tk.Frame(parent, bg=THEME['panel'])
        nb_wrap.pack(fill=tk.BOTH, expand=True, padx=14, pady=(0, 8))

        nb = ttk.Notebook(nb_wrap, style="Modern.TNotebook")
        nb.pack(fill=tk.BOTH, expand=True)

        t1 = tk.Frame(nb, bg=THEME['panel'])
        t3 = tk.Frame(nb, bg=THEME['panel'])
        nb.add(t1, text='  数据分析  ')
        nb.add(t3, text='  Excel 导入  ')

        self.setup_tab1(t1)
        self.setup_tab3(t3)

        self._notebook = nb
        self._debug_tab_active = False
        self.main_notebook = nb

        # 导出区
        export_card = Card(parent, pad=12)
        export_card.pack(fill=tk.X, padx=14, pady=(0, 10))
        tk.Label(
            export_card.inner, text="导出报告", bg=THEME['card'], fg=THEME['fg_muted'],
            font=(FONT_UI, 9)
        ).pack(anchor='w', pady=(0, 8))

        self.btn_export = ModernButton(
            export_card.inner, text="导出当前报告 (PDF)", style='success',
            command=self.export_report, height=52, font=(FONT_UI, 12, 'bold')
        )
        self.btn_export.pack(fill=tk.X, pady=(0, 6))

        self.btn_batch_export = ModernButton(
            export_card.inner, text="合并导出全部 PDF", style='primary',
            command=self.export_merged_report, height=52, font=(FONT_UI, 12, 'bold')
        )
        # 常驻显示（无数据时点击会提示先导入 Excel），不再依赖 tab 切换事件
        self.btn_batch_export.pack(fill=tk.X, pady=(0, 6))

        if not REPORTLAB_AVAILABLE:
            self.btn_export.configure(state=tk.DISABLED, text="缺少 reportlab")
            self.btn_batch_export.configure(state=tk.DISABLED, text="缺少 reportlab")
            tk.Label(
                export_card.inner, text="pip install reportlab", bg=THEME['card'],
                fg=THEME['fg_dim'], font=(FONT_UI, 8)
            ).pack(anchor='w')

        if not PANDAS_AVAILABLE:
            tk.Label(
                parent, text="缺少 pandas/openpyxl，Excel 导入不可用",
                bg=THEME['panel'], fg=THEME['danger'], font=(FONT_UI, 9)
            ).pack(pady=(0, 6))

        nb.bind("<<NotebookTabChanged>>", self.on_tab_changed)

        self._debug_btn = tk.Label(
            parent, text="调试模式", bg=THEME['panel'], fg=THEME['fg_dim'],
            font=(FONT_UI, 8), cursor="hand2"
        )
        self._debug_btn.pack(side=tk.BOTTOM, pady=(0, 12))
        self._debug_btn.bind("<Button-1>", lambda e: self._toggle_debug())
        self._debug_btn.bind("<Enter>", lambda e: self._debug_btn.configure(fg=THEME['fg_muted']))
        self._debug_btn.bind("<Leave>", lambda e: self._debug_btn.configure(
            fg=THEME['success'] if self._debug_tab_active else THEME['fg_dim']
        ))

    def on_tab_changed(self, event):
        # 批量导出按钮已常驻显示，这里只管理"导出当前报告"按钮的可用状态
        selected_tab = event.widget.tab('current')['text']
        if "Excel" in selected_tab:
            if self.excel_projects and self.current_excel_index != -1:
                self.btn_export.configure(state=tk.NORMAL)
            else:
                self.btn_export.configure(state=tk.DISABLED)
        else:
            if self.current_stats:
                self.btn_export.configure(state=tk.NORMAL)
            else:
                self.btn_export.configure(state=tk.DISABLED)

    def init_stats_panel(self, parent):
        header = tk.Frame(parent, bg=THEME['panel'])
        header.pack(fill=tk.X, padx=14, pady=(16, 8))

        self.lbl_proj_display = tk.Label(
            header, text="未命名项目", bg=THEME['panel'], fg=THEME['accent_text'],
            font=(FONT_UI, 12, 'bold'), wraplength=320, justify=tk.LEFT, anchor='w'
        )
        self.lbl_proj_display.pack(anchor='w')
        tk.Label(
            header, text="详细质量指标", bg=THEME['panel'], fg=THEME['fg_dim'],
            font=(FONT_UI, 9)
        ).pack(anchor='w', pady=(2, 0))

        # Cpk 大号高亮
        hero = Card(parent, pad=14)
        hero.pack(fill=tk.X, padx=12, pady=(4, 10))
        top = tk.Frame(hero.inner, bg=THEME['card'])
        top.pack(fill=tk.X)
        left = tk.Frame(top, bg=THEME['card'])
        left.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(left, text="Cpk", bg=THEME['card'], fg=THEME['fg_muted'],
                 font=(FONT_UI, 9)).pack(anchor='w')
        self.hero_cpk = tk.Label(
            left, text="—", bg=THEME['card'], fg=THEME['fg'],
            font=(FONT_MONO, 28, 'bold')
        )
        self.hero_cpk.pack(anchor='w')
        self.hero_level = tk.Label(
            top, text="待分析", bg=THEME['elevated'], fg=THEME['fg_muted'],
            font=(FONT_UI, 10, 'bold'), padx=12, pady=6
        )
        self.hero_level.pack(side=tk.RIGHT, anchor='n')

        # 指标网格
        grid_wrap = tk.Frame(parent, bg=THEME['panel'])
        grid_wrap.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 12))

        self.stat_labels = {}
        fields = [
            ("Count", "样本数 N"), ("Mean", "均值"),
            ("StdDev", "标准差"), ("Median", "中位数"),
            ("CV", "变异系数%"), ("Range", "极差"),
            ("Max", "最大值"), ("Min", "最小值"),
            ("Skewness", "偏度"), ("Kurtosis", "峰度"),
            ("USL", "规格上限"), ("LSL", "规格下限"),
            ("Cp", "Cp"), ("SigmaLevel", "Sigma水平"),
            ("PPM", "PPM"), ("OutOfSpecCount", "超规数"),
        ]

        for i, (key, label) in enumerate(fields):
            r, c = divmod(i, 2)
            chip = MetricChip(grid_wrap, label=label, value="-")
            chip.grid(row=r, column=c, sticky='nsew', padx=3, pady=3)
            self.stat_labels[key] = chip

        grid_wrap.columnconfigure(0, weight=1)
        grid_wrap.columnconfigure(1, weight=1)
        # Cpk / CPK_LEVEL 在 hero 中展示，但仍保留映射兼容
        self.stat_labels["Cpk"] = None
        self.stat_labels["CPK_LEVEL"] = None

    def update_stats_display(self, stats, project_name=None):
        if project_name:
            pname = project_name
        else:
            pname = self.inp_project.get().strip() or "未命名项目"

        self.lbl_proj_display.config(text=pname)
        if not project_name:
            self.project_name = pname

        def fmt_val(v, precision=4):
            if v is None:
                return "N/A"
            return f"{v:.{precision}f}"

        # Hero Cpk
        cpk = stats.get('Cpk')
        level = stats.get('CPK_LEVEL', '')
        level_colors = {
            "优秀": THEME['success'], "良好": "#a3e635",
            "一般": THEME['warning'], "较差": "#fb923c", "很差": THEME['danger']
        }
        level_bgs = {
            "优秀": THEME['success_soft'], "良好": "#243018",
            "一般": THEME['warning_soft'], "较差": "#3a2410", "很差": THEME['danger_soft']
        }
        if cpk is not None:
            self.hero_cpk.configure(text=f"{cpk:.3f}", fg=level_colors.get(level, THEME['fg']))
            self.hero_level.configure(
                text=level or "—",
                fg=level_colors.get(level, THEME['fg_muted']),
                bg=level_bgs.get(level, THEME['elevated'])
            )
        else:
            self.hero_cpk.configure(text="—", fg=THEME['fg'])
            self.hero_level.configure(text="—", fg=THEME['fg_muted'], bg=THEME['elevated'])

        for key, chip in self.stat_labels.items():
            if chip is None or key not in stats:
                continue
            val = stats[key]
            color = THEME['fg']

            if key in ["Count", "OutOfSpecCount"]:
                txt = f"{int(val)}"
                if key == "OutOfSpecCount":
                    color = THEME['danger'] if int(val) > 0 else THEME['success']
            elif key == "PPM":
                txt = f"{int(val)}"
            elif key == "SigmaLevel":
                txt = f"{val:.2f}σ"
            elif key == "CV":
                txt = f"{val:.2f}%"
            elif key in ["USL", "LSL"] and val is None:
                txt = "未设置"
                color = THEME['fg_dim']
            elif key in ["Skewness", "Kurtosis"]:
                txt = fmt_val(val, 3)
                color = THEME['warning'] if abs(val) > 1.0 else THEME['success']
            elif key in ['Cp', 'Cpk', 'CPU', 'CPL']:
                txt = fmt_val(val, 3)
            else:
                txt = fmt_val(val, 4)

            if val is None and key not in ["USL", "LSL"]:
                color = THEME['fg_dim']

            chip.set_value(txt, color)

        self.set_status(f"Cpk {cpk:.3f} · {level}" if cpk is not None else "已更新", 'ok')

    def init_chart_panel(self, parent):
        header = tk.Frame(parent, bg=THEME['panel'])
        header.pack(fill=tk.X, padx=16, pady=(14, 6))
        tk.Label(
            header, text="分布直方图", bg=THEME['panel'], fg=THEME['fg'],
            font=(FONT_UI, 13, 'bold')
        ).pack(side=tk.LEFT)
        tk.Label(
            header, text="正态拟合 · 规格限 · 均值", bg=THEME['panel'],
            fg=THEME['fg_dim'], font=(FONT_UI, 9)
        ).pack(side=tk.LEFT, padx=10)

        chart_box = tk.Frame(parent, bg=THEME['bg'])
        chart_box.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        self.fig, self.ax = plt.subplots(figsize=(5, 5))
        self.fig.patch.set_facecolor(THEME['bg'])
        self.canvas = FigureCanvasTkAgg(self.fig, master=chart_box)
        self.canvas.get_tk_widget().configure(bg=THEME['bg'], highlightthickness=0)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        self.reset_chart()

    def create_input(self, parent, label, row, default=""):
        tk.Label(
            parent, text=label, bg=parent.cget('bg'), fg=THEME['fg_muted'],
            font=(FONT_UI, 9)
        ).grid(row=row, column=0, sticky='w', pady=6)
        wrap = tk.Frame(parent, bg=THEME['border'])
        wrap.grid(row=row, column=1, sticky='ew', padx=(10, 0), pady=6)
        e = tk.Entry(
            wrap, bg=THEME['input_bg'], fg=THEME['fg'], insertbackground=THEME['accent'],
            relief=tk.FLAT, font=(FONT_MONO, 11), bd=0
        )
        if default:
            e.insert(0, default)
        e.pack(fill=tk.X, ipady=7, padx=1, pady=1)
        e.bind('<FocusIn>', lambda ev: wrap.configure(bg=THEME['focus']))
        e.bind('<FocusOut>', lambda ev: wrap.configure(bg=THEME['border']))
        return e

    def setup_tab1(self, f):
        f.configure(bg=THEME['panel'])
        # 可滚动区域
        canvas = tk.Canvas(f, bg=THEME['panel'], highlightthickness=0, bd=0)
        scroll = ttk.Scrollbar(f, orient='vertical', style='Modern.Vertical.TScrollbar', command=canvas.yview)
        inner = tk.Frame(canvas, bg=THEME['panel'])
        inner.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=inner, anchor='nw')
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        pad = tk.Frame(inner, bg=THEME['panel'], padx=12, pady=12)
        pad.pack(fill=tk.BOTH, expand=True)
        pad.columnconfigure(1, weight=1)

        tk.Label(
            pad, text="规格限", bg=THEME['panel'], fg=THEME['fg'],
            font=(FONT_UI, 10, 'bold')
        ).grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 4))
        self.inp_an_usl = self.create_input(pad, "上限 USL", 1)
        self.inp_an_lsl = self.create_input(pad, "下限 LSL", 2)

        tk.Label(
            pad, text="测量数据", bg=THEME['panel'], fg=THEME['fg'],
            font=(FONT_UI, 10, 'bold')
        ).grid(row=3, column=0, columnspan=2, sticky='w', pady=(14, 4))
        tk.Label(
            pad, text="支持空格 / 换行 / 逗号分隔", bg=THEME['panel'],
            fg=THEME['fg_dim'], font=(FONT_UI, 8)
        ).grid(row=4, column=0, columnspan=2, sticky='w')

        text_wrap = tk.Frame(pad, bg=THEME['border'])
        text_wrap.grid(row=5, column=0, columnspan=2, sticky='nsew', pady=(6, 0))
        self.txt_data = tk.Text(
            text_wrap, bg=THEME['input_bg'], fg=THEME['fg'], height=14,
            relief=tk.FLAT, font=(FONT_MONO, 10), insertbackground=THEME['accent'],
            bd=0, padx=8, pady=8, selectbackground=THEME['accent_soft'],
            selectforeground=THEME['accent_text']
        )
        self.txt_data.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        pad.rowconfigure(5, weight=1)

        self.create_btn_bar(pad, 6, self.on_analyze, self.on_clear_tab1, "开始分析")

    def setup_tab2(self, f):
        f.configure(bg=THEME['panel'])
        pad = tk.Frame(f, bg=THEME['panel'], padx=12, pady=12)
        pad.pack(fill=tk.BOTH, expand=True)
        pad.columnconfigure(1, weight=1)

        tk.Label(
            pad, text="模拟参数", bg=THEME['panel'], fg=THEME['fg'],
            font=(FONT_UI, 10, 'bold')
        ).grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 4))

        self.inp_sim_usl = self.create_input(pad, "上限 USL", 1)
        self.inp_sim_lsl = self.create_input(pad, "下限 LSL", 2)
        self.inp_sim_cpk = self.create_input(pad, "目标 Cpk", 3, "1.33")
        self.inp_sim_mean = self.create_input(pad, "目标均值", 4, "10.0")
        self.inp_sim_cnt = self.create_input(pad, "数量", 5, "50")
        self.inp_sim_prec = self.create_input(pad, "小数精度", 6, "3")
        self.create_btn_bar(pad, 7, self.on_simulate, self.on_clear_tab2, "生成数据")

        tk.Label(
            pad, text="结果预览", bg=THEME['panel'], fg=THEME['fg'],
            font=(FONT_UI, 10, 'bold')
        ).grid(row=8, column=0, columnspan=2, sticky='w', pady=(12, 4))
        text_wrap = tk.Frame(pad, bg=THEME['border'])
        text_wrap.grid(row=9, column=0, columnspan=2, sticky='nsew')
        self.txt_sim = tk.Text(
            text_wrap, bg=THEME['input_bg'], fg=THEME['success'], height=8,
            relief=tk.FLAT, font=(FONT_MONO, 10), insertbackground=THEME['accent'],
            bd=0, padx=8, pady=8
        )
        self.txt_sim.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        ModernButton(pad, text="复制结果", style='ghost', command=self.on_copy).grid(
            row=10, column=0, columnspan=2, sticky='ew', pady=8
        )

    def setup_tab3(self, f):
        f.configure(bg=THEME['panel'])
        f.columnconfigure(0, weight=1)
        # 纵向布局：项目列表在上（占大头），数据预览在下。
        # 注意：不能用左右两列布局——预览文本框请求尺寸很大（默认 80 字符宽），
        # grid 空间不足时会把项目列表列压成 1px，导致列表不可见、无法切换项目。
        f.rowconfigure(1, weight=3)
        f.rowconfigure(2, weight=2)

        top = tk.Frame(f, bg=THEME['panel'])
        top.grid(row=0, column=0, columnspan=2, sticky='ew', pady=10, padx=12)

        ModernButton(top, text="导入 Excel", style='primary', command=self.load_excel_file).pack(side=tk.LEFT, padx=(0, 8))
        ModernButton(top, text="清空列表", style='danger', command=self.clear_excel_data).pack(side=tk.LEFT)
        tk.Label(
            top, text="点击列表项查看分析", bg=THEME['panel'], fg=THEME['fg_dim'],
            font=(FONT_UI, 8)
        ).pack(side=tk.RIGHT)

        list_card = Card(f, pad=0)
        list_card.grid(row=1, column=0, columnspan=2, sticky='nsew', padx=12, pady=(0, 6))
        list_inner = list_card.inner
        list_inner.configure(bg=THEME['card'])

        cols = ('sheet', 'cpk', 'level')
        self.tree_projects = ttk.Treeview(
            list_inner, columns=cols, displaycolumns=cols,
            selectmode='browse', style='Modern.Treeview'
        )
        self.tree_projects.heading('#0', text='项目', anchor='w')
        self.tree_projects.heading('sheet', text='子表', anchor='center')
        self.tree_projects.heading('cpk', text='Cpk', anchor='center')
        self.tree_projects.heading('level', text='等级', anchor='center')

        self.tree_projects.column('#0', width=140, anchor='w')
        self.tree_projects.column('sheet', width=70, anchor='center')
        self.tree_projects.column('cpk', width=60, anchor='center')
        self.tree_projects.column('level', width=60, anchor='center')

        scrollbar = ttk.Scrollbar(list_inner, orient="vertical", style='Modern.Vertical.TScrollbar',
                                  command=self.tree_projects.yview)
        self.tree_projects.configure(yscrollcommand=scrollbar.set)
        self.tree_projects.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4, pady=4)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=4, padx=(0, 4))
        self.tree_projects.bind('<<TreeviewSelect>>', self.on_excel_item_select)
        self.tree_projects.bind('<ButtonRelease-1>', self.on_excel_item_select)

        preview_card = Card(f, pad=10)
        preview_card.grid(row=2, column=0, columnspan=2, sticky='nsew', padx=12, pady=(6, 12))
        tk.Label(
            preview_card.inner, text="数据预览", bg=THEME['card'], fg=THEME['fg'],
            font=(FONT_UI, 10, 'bold')
        ).pack(anchor='w', pady=(0, 6))
        text_wrap = tk.Frame(preview_card.inner, bg=THEME['border'])
        text_wrap.pack(fill=tk.BOTH, expand=True)
        self.txt_excel_preview = tk.Text(
            text_wrap, bg=THEME['input_bg'], fg=THEME['fg_muted'], relief=tk.FLAT,
            font=(FONT_MONO, 9), bd=0, padx=8, pady=8, insertbackground=THEME['accent'],
            width=1, height=1  # 请求尺寸设小，实际尺寸由 pack 拉伸决定
        )
        self.txt_excel_preview.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

    def create_btn_bar(self, parent, row, cmd1, cmd2, lbl1):
        box = tk.Frame(parent, bg=parent.cget('bg'))
        box.grid(row=row, column=0, columnspan=2, sticky='ew', pady=12)
        box.columnconfigure(0, weight=3)
        box.columnconfigure(1, weight=1)
        ModernButton(box, text=lbl1, style='primary', command=cmd1).grid(
            row=0, column=0, sticky='ew', padx=(0, 6)
        )
        ModernButton(box, text="清空", style='ghost', command=cmd2).grid(
            row=0, column=1, sticky='ew'
        )

    def get_val(self, entry, is_int=False, allow_empty=False):
        val_str = entry.get().strip()
        if not val_str:
            return None if allow_empty else False
        try:
            val = float(val_str)
            return int(val) if is_int else val
        except Exception:
            return False

    def on_analyze(self):
        usl = self.get_val(self.inp_an_usl, allow_empty=True)
        lsl = self.get_val(self.inp_an_lsl, allow_empty=True)
        if usl is False or lsl is False:
            show_msg(self.root, "输入错误", "规格值必须是数字")
            return
        if usl is None and lsl is None:
            show_msg(self.root, "缺失规格", "请至少输入一个规格限")
            return

        raw = self.txt_data.get("1.0", tk.END)
        nums = re.findall(r"[-+]?\d*\.?\d+|\d+", raw)
        try:
            data = np.array([float(x) for x in nums])
            if len(data) < 2:
                raise ValueError
        except Exception:
            show_msg(self.root, "数据错误", "请检查输入数据")
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
            show_msg(self.root, "输入错误", "请检查数值格式")
            return
        if usl is None and lsl is None:
            show_msg(self.root, "缺失规格", "请至少输入一个规格限")
            return

        data = CpkCalculator.simulate(cpk, mean, usl, lsl, cnt, max(0, min(prec, 10)))
        if data is None:
            show_msg(self.root, "错误", "无法生成数据")
            return

        fmt = f"{{:.{prec}f}}"
        self.txt_sim.delete("1.0", tk.END)
        self.txt_sim.insert(tk.END, "\n".join([fmt.format(x) for x in data]))
        self.process_result(data, usl, lsl)

    def process_result(self, data, usl, lsl, project_name=None):
        stats = CpkCalculator.calculate(data, usl, lsl)
        if "Error" in stats:
            show_msg(self.root, "计算错误", stats["Error"])
            self.set_status("计算失败", 'err')
            return

        self.current_data = data
        self.current_stats = stats
        self.current_usl = usl
        self.current_lsl = lsl
        self.update_stats_display(stats, project_name)
        self.draw_chart(data, stats)
        if "Excel" not in self.main_notebook.tab('current', 'text'):
            self.btn_export.configure(state=tk.NORMAL)

    def draw_chart(self, data, stats):
        if not hasattr(self, '_is_exporting') or not self._is_exporting:
            self.ax.clear()
            self.ax.set_facecolor(THEME['bg'])
            mu, sigma = stats['Mean'], stats['StdDev']
            usl, lsl = stats['USL'], stats['LSL']

            self.ax.hist(
                data, bins=30, density=True, alpha=0.55,
                color=THEME['chart_fill'], edgecolor='none'
            )
            xmin, xmax = self.ax.get_xlim()
            base_span = 6 * sigma if sigma > 0 else 1.0
            plot_min = lsl - base_span * 0.2 if lsl is not None else min(xmin, mu - 4 * sigma)
            plot_max = usl + base_span * 0.2 if usl is not None else max(xmax, mu + 4 * sigma)

            x = np.linspace(plot_min, plot_max, 500)
            y = norm.pdf(x, mu, sigma)
            self.ax.plot(x, y, color=THEME['chart_curve'], linewidth=2.2)
            self.ax.fill_between(x, y, alpha=0.15, color=THEME['chart_curve'])

            ymax = max(y) * 1.2 if len(y) > 0 else 1
            self.ax.set_ylim(0, ymax)

            if usl is not None:
                self.ax.axvline(usl, c=THEME['chart_spec'], ls='--', lw=1.6)
                self.ax.text(usl, ymax * 0.95, "USL", c=THEME['chart_spec'], ha='center', fontsize=9)
            if lsl is not None:
                self.ax.axvline(lsl, c=THEME['chart_spec'], ls='--', lw=1.6)
                self.ax.text(lsl, ymax * 0.95, "LSL", c=THEME['chart_spec'], ha='center', fontsize=9)

            self.ax.axvline(mu, c=THEME['chart_mean'], ls='-', lw=1.5, alpha=0.9)
            self.ax.text(mu, ymax * 0.85, f"μ={mu:.3f}", c=THEME['chart_mean'], ha='center', fontsize=9)

            self.ax.set_xlabel("测量值", fontsize=10, color=THEME['fg_muted'], labelpad=6)
            self.ax.set_ylabel("概率密度", fontsize=10, color=THEME['fg_muted'], labelpad=6)
            self.ax.tick_params(colors=THEME['fg_dim'], labelsize=9)
            self.ax.grid(True, linestyle='--', alpha=0.12, color=THEME['fg_muted'])

            for spine in ['top', 'right']:
                self.ax.spines[spine].set_visible(False)
            self.ax.spines['left'].set_color(THEME['border'])
            self.ax.spines['bottom'].set_color(THEME['border'])

            self.canvas.draw()

    def reset_chart(self):
        if not hasattr(self, '_is_exporting') or not self._is_exporting:
            self.ax.clear()
            self.ax.set_facecolor(THEME['bg'])
            self.ax.axis('off')
            self.ax.text(
                0.5, 0.5, "等待数据…\n在左侧输入测量值或导入 Excel",
                color=THEME['fg_dim'], ha='center', va='center',
                fontsize=12, transform=self.ax.transAxes,
                linespacing=1.6
            )
            self.canvas.draw()
            for k, chip in self.stat_labels.items():
                if chip is not None:
                    chip.set_value("-", THEME['fg_dim'])
            if hasattr(self, 'hero_cpk'):
                self.hero_cpk.configure(text="—", fg=THEME['fg'])
                self.hero_level.configure(text="待分析", fg=THEME['fg_muted'], bg=THEME['elevated'])
            self.lbl_proj_display.config(text="未命名项目")
            self.current_data = None
            self.current_stats = None
            self.set_status("就绪", 'ok')

    def on_clear_tab1(self):
        self.inp_an_usl.delete(0, tk.END)
        self.inp_an_lsl.delete(0, tk.END)
        self.txt_data.delete("1.0", tk.END)
        self.reset_chart()
        self.btn_export.configure(state=tk.DISABLED)

    def on_clear_tab2(self):
        for e in [self.inp_sim_usl, self.inp_sim_lsl, self.inp_sim_cpk,
                  self.inp_sim_mean, self.inp_sim_cnt, self.inp_sim_prec]:
            e.delete(0, tk.END)
        self.txt_sim.delete("1.0", tk.END)
        self.reset_chart()
        self.btn_export.configure(state=tk.DISABLED)

    def on_copy(self):
        self.root.clipboard_clear()
        self.root.clipboard_append(self.txt_sim.get("1.0", tk.END))
        show_msg(self.root, "复制成功", "内容已复制到剪贴板", False)

    def _set_dark_titlebar(self):
        try:
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            hwnd = self.root.winfo_id()
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE,
                ctypes.byref(ctypes.c_int(1)),
                ctypes.sizeof(ctypes.c_int(1))
            )
        except Exception:
            pass

    def show_about(self):
        about_text = (
            "CPK 统计分析工具 V7.4 (Modern UI)\n\n"
            "更新内容:\n"
            "• 现代化深色界面与卡片式布局\n"
            "• 指标芯片、焦点反馈与状态提示\n"
            "• 极致压缩 PDF，单页容纳约 150 条数据\n"
            "• 导出默认路径为软件所在目录"
        )
        show_msg(self.root, "关于软件", about_text, is_error=False)

    def _show_debug_prompt(self, event=None):
        if self._debug_tab_active:
            return
        pwd = simpledialog.askstring("调试模式", "请输入调试密码：", show='*', parent=self.root)
        if pwd == "114514":
            t2 = tk.Frame(self._notebook, bg=THEME['panel'])
            self._notebook.insert(1, t2, text='  模拟生成  ')
            self.setup_tab2(t2)
            self._debug_tab_active = True
            self._debug_btn.config(fg=THEME['success'], text="调试模式 ✓")
            show_msg(self.root, "调试模式", "模拟生成模块已激活", is_error=False)
            self.set_status("调试模式", 'info')

    def _toggle_debug(self, event=None):
        self._show_debug_prompt()

    # ==========================================
    # 5. Excel 导入相关功能
    # ==========================================
    def load_excel_file(self):
        if not PANDAS_AVAILABLE:
            show_msg(self.root, "缺少依赖", "未安装 pandas 或 openpyxl。\n请运行：pip install pandas openpyxl")
            return

        file_path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if not file_path:
            return

        try:
            self.set_status("导入中…", 'info')
            xl = pd.ExcelFile(file_path)
            sheet_names = xl.sheet_names
            is_comparison_mode = len(sheet_names) > 1

            all_projects = []
            error_logs = []

            for sheet_name in sheet_names:
                df = xl.parse(sheet_name, header=None)

                if df.shape[0] < 4:
                    error_logs.append(f"[{sheet_name}] 数据行不足 (需要≥4行)")
                    continue

                num_cols = df.shape[1]

                for col_idx in range(num_cols):
                    col_data = df.iloc[:, col_idx]

                    raw_project_name = col_data.iloc[0]
                    if pd.isna(raw_project_name) or str(raw_project_name).strip() == "":
                        project_name = f"Project_{col_idx + 1}"
                    else:
                        project_name = str(raw_project_name).strip()

                    usl_val = col_data.iloc[1]
                    usl = float(usl_val) if not pd.isna(usl_val) else None

                    lsl_val = col_data.iloc[2]
                    lsl = float(lsl_val) if not pd.isna(lsl_val) else None

                    raw_values = col_data.iloc[3:]
                    # 严格过滤空值：NaN、空字符串、纯空格、非数值内容一律忽略
                    def _is_valid_numeric(v):
                        if pd.isna(v):
                            return False
                        s = str(v).strip()
                        if s == '':
                            return False
                        try:
                            float(s)
                            return True
                        except ValueError:
                            return False

                    valid_mask = raw_values.apply(_is_valid_numeric)
                    skipped_count = (~valid_mask).sum()
                    raw_data = raw_values[valid_mask]

                    if len(raw_data) < 2:
                        error_logs.append(
                            f"[{sheet_name}] {project_name}: 有效数据不足 (仅 {len(raw_data)} 条，已忽略 {skipped_count} 个空值)"
                        )
                        continue

                    try:
                        data_array = np.array(raw_data.astype(float))
                    except ValueError:
                        error_logs.append(f"[{sheet_name}] {project_name}: 数据格式错误")
                        continue

                    if skipped_count > 0:
                        error_logs.append(
                            f"[{sheet_name}] {project_name}: 已自动忽略 {skipped_count} 个空值/无效值，有效数据 {len(data_array)} 条"
                        )

                    stats = CpkCalculator.calculate(data_array, usl, lsl)

                    if "Error" in stats:
                        error_logs.append(f"[{sheet_name}] {project_name}: {stats['Error']}")
                        continue

                    all_projects.append({
                        "name": project_name,
                        "sheet_name": sheet_name,
                        "data": data_array,
                        "usl": usl,
                        "lsl": lsl,
                        "stats": stats,
                        "cpk_val": stats['Cpk'],
                        "level": stats['CPK_LEVEL']
                    })

            if is_comparison_mode:
                sheet_groups = {}
                for sn in sheet_names:
                    sheet_groups[sn] = []
                for p in all_projects:
                    sheet_groups[p['sheet_name']].append(p)

                max_count = max(len(v) for v in sheet_groups.values()) if sheet_groups else 0
                interleaved = []
                for i in range(max_count):
                    for sn in sheet_names:
                        if i < len(sheet_groups[sn]):
                            interleaved.append(sheet_groups[sn][i])
                all_projects = interleaved

            if not all_projects:
                msg = "未找到有效数据。"
                if error_logs:
                    msg += "\n\n错误详情:\n" + "\n".join(error_logs[:5])
                show_msg(self.root, "导入失败", msg)
                self.set_status("导入失败", 'err')
                return

            self.excel_projects = all_projects
            self.refresh_excel_treeview()

            if self.excel_projects:
                self.tree_projects.selection_set(self.tree_projects.get_children()[0])
                self.on_excel_item_select(None)

            msg = f"成功导入 {len(self.excel_projects)} 个项目"
            if is_comparison_mode:
                msg += f"（{len(sheet_names)} 个子表，交叉对比模式）"
            msg += "。"
            if error_logs:
                msg += f"\n跳过 {len(error_logs)} 个无效项目。"
            show_msg(self.root, "导入成功", msg, is_error=False)
            self.set_status(f"已导入 {len(self.excel_projects)} 项", 'ok')

        except Exception as e:
            show_msg(self.root, "导入失败", f"读取 Excel 文件时出错:\n{str(e)}")
            self.set_status("导入失败", 'err')

    def refresh_excel_treeview(self):
        for item in self.tree_projects.get_children():
            self.tree_projects.delete(item)

        for i, proj in enumerate(self.excel_projects):
            cpk = proj['cpk_val']
            level = proj['level']
            sheet_name = proj.get('sheet_name', 'Sheet1')
            self.tree_projects.insert(
                '', 'end', iid=str(i), text=proj['name'],
                values=(sheet_name, f"{cpk:.3f}", level)
            )

    def on_excel_item_select(self, event=None):
        selection = self.tree_projects.selection()
        if not selection:
            return
        iid = selection[0]
        if not iid or iid == '':
            return
        try:
            idx = int(iid)
        except ValueError:
            return

        if idx < 0 or idx >= len(self.excel_projects):
            return

        self.current_excel_index = idx
        project = self.excel_projects[idx]

        self.update_stats_display(project['stats'], project_name=project['name'])
        self.draw_chart(project['data'], project['stats'])

        self.txt_excel_preview.delete("1.0", tk.END)
        s = project['stats']
        preview_text = f"【{project['name']}】详细报告\n"
        if project.get('sheet_name'):
            preview_text += f"子表: {project['sheet_name']}\n"
        preview_text += "=" * 30 + "\n"
        preview_text += f"Cpk: {s['Cpk']:.4f} ({s['CPK_LEVEL']})\n"
        preview_text += f"USL: {s['USL']} | LSL: {s['LSL']}\n"
        preview_text += f"均值：{s['Mean']:.4f} | 中位数：{s['Median']:.4f}\n"
        preview_text += f"标准差：{s['StdDev']:.4f} | CV: {s['CV']:.2f}%\n"
        preview_text += f"偏度：{s['Skewness']:.3f} | 峰度：{s['Kurtosis']:.3f}\n"
        preview_text += f"最大：{s['Max']:.4f} | 最小：{s['Min']:.4f} | 极差：{s['Range']:.4f}\n"
        preview_text += f"PPM: {int(s['PPM'])} | 超规数：{s['OutOfSpecCount']}\n"
        preview_text += "=" * 30 + "\n\n数据前 30 行:\n"

        for i, val in enumerate(project['data'][:30]):
            preview_text += f"{val}\n"
        if len(project['data']) > 30:
            preview_text += f"... 共 {len(project['data'])} 条数据"

        self.txt_excel_preview.insert("1.0", preview_text)
        self.root.update_idletasks()

        if "Excel" in self.main_notebook.tab('current', 'text'):
            self.btn_export.configure(state=tk.NORMAL)

    def clear_excel_data(self):
        self.excel_projects = []
        self.current_excel_index = -1
        for item in self.tree_projects.get_children():
            self.tree_projects.delete(item)
        self.txt_excel_preview.delete("1.0", tk.END)
        self.reset_chart()
        if "Excel" in self.main_notebook.tab('current', 'text'):
            self.btn_export.configure(state=tk.DISABLED)
        self.set_status("列表已清空", 'info')

    def export_merged_report(self):
        if not self.excel_projects:
            show_msg(self.root, "无数据", "没有可导出的项目数据。请先导入 Excel。")
            return

        if not REPORTLAB_AVAILABLE:
            show_msg(self.root, "缺少依赖", "未安装 reportlab。")
            return

        default_filename = f"CPK_汇总报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"

        file_path = filedialog.asksaveasfilename(
            title="保存汇总报告 (所有项目将合并为此文件)",
            defaultextension=".pdf",
            initialfile=default_filename,
            filetypes=[("PDF 文件", "*.pdf"), ("所有文件", "*.*")],
            initialdir=self.app_dir
        )

        if not file_path:
            return
        if not file_path.lower().endswith('.pdf'):
            file_path += '.pdf'

        try:
            self.btn_batch_export.configure(state=tk.DISABLED, text="生成中...")
            self.set_status("导出中…", 'info')
            self.root.update_idletasks()

            self._generate_merged_pdf_report(file_path, self.excel_projects)

            show_msg(self.root, "导出成功", f"所有 {len(self.excel_projects)} 个项目已合并保存至:\n{file_path}", is_error=False)
            self.set_status("导出完成", 'ok')
        except Exception as e:
            show_msg(self.root, "导出失败", f"错误:\n{str(e)}\n{traceback.format_exc()}")
            self.set_status("导出失败", 'err')
        finally:
            self.btn_batch_export.configure(state=tk.NORMAL, text="合并导出全部 PDF")

    def export_report(self):
        if not REPORTLAB_AVAILABLE:
            show_msg(self.root, "缺少依赖", "未安装 reportlab。")
            return

        current_tab_text = self.main_notebook.tab('current', 'text')
        if "Excel" in current_tab_text:
            if not self.excel_projects or self.current_excel_index == -1:
                show_msg(self.root, "无数据", "请先在列表中选择一个项目。")
                return
            project = self.excel_projects[self.current_excel_index]
            data_to_export = project['data']
            stats_to_export = project['stats']
            usl_to_export = project['usl']
            lsl_to_export = project['lsl']
            name_to_export = project['name']
        elif self.current_stats is None:
            show_msg(self.root, "无数据", "请先进行分析或模拟生成数据。")
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
            initialdir=self.app_dir
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

            self.set_status("导出中…", 'info')
            self._generate_single_pdf_logic(file_path)

            self.current_data = old_data
            self.current_stats = old_stats
            self.current_usl = old_usl
            self.current_lsl = old_lsl
            self.project_name = old_name

            show_msg(self.root, "导出成功", f"报告已保存至:\n{file_path}", is_error=False)
            self.set_status("导出完成", 'ok')
        except Exception as e:
            show_msg(self.root, "导出失败", f"错误:\n{str(e)}\n{traceback.format_exc()}")
            self.set_status("导出失败", 'err')

    def _get_pdf_font(self):
        font_name = "Helvetica"
        try:
            if os.name == 'nt':
                path = r"C:\Windows\Fonts\simhei.ttf"
                if os.path.exists(path):
                    pdfmetrics.registerFont(TTFont('SimHei', path))
                    font_name = 'SimHei'
        except Exception:
            pass
        return font_name

    def _generate_merged_pdf_report(self, file_path, projects_list):
        """批量导出：一页两个项目（上下双 Frame），固定区域避免重叠/分页拆散。"""
        temp_img_paths = []
        try:
            page_margin = 0.7 * cm
            pw, ph = A4
            usable_w = pw - 2 * page_margin
            usable_h = ph - 2 * page_margin
            gap_h = 0.30 * cm
            half_h = (usable_h - gap_h) / 2.0

            # 目录页：整页单 Frame
            frame_full = Frame(
                page_margin, page_margin, usable_w, usable_h,
                id='full',
                leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
                showBoundary=0,
            )
            # 详情页：上/下两个固定半页 Frame —— 保证一页两项、互不侵占
            frame_top = Frame(
                page_margin, page_margin + half_h + gap_h, usable_w, half_h,
                id='top',
                leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
                showBoundary=0,
            )
            frame_bot = Frame(
                page_margin, page_margin, usable_w, half_h,
                id='bot',
                leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
                showBoundary=0,
            )

            doc = BaseDocTemplate(
                file_path, pagesize=A4,
                leftMargin=page_margin, rightMargin=page_margin,
                topMargin=page_margin, bottomMargin=page_margin,
            )
            doc.addPageTemplates([
                PageTemplate(id='toc', frames=[frame_full]),
                PageTemplate(id='two_up', frames=[frame_top, frame_bot]),
            ])

            story = []
            styles = getSampleStyleSheet()
            font_name = self._get_pdf_font()

            title_style = ParagraphStyle(
                'BatchTitle', parent=styles['Heading1'], fontName=font_name,
                fontSize=15, alignment=TA_CENTER, spaceAfter=4,
                textColor=colors.HexColor('#0f172a'), leading=18
            )
            sub_style = ParagraphStyle(
                'BatchSub', parent=styles['Normal'], fontName=font_name,
                fontSize=8, alignment=TA_CENTER, spaceAfter=6,
                textColor=colors.HexColor('#64748b'), leading=10
            )
            head_style = ParagraphStyle(
                'BatchHead', parent=styles['Heading3'], fontName=font_name,
                fontSize=10, spaceBefore=2, spaceAfter=3,
                textColor=colors.HexColor('#1e40af'), leading=12
            )

            # ---- 目录页 ----
            toc_level_colors = {
                '优秀': '#2563eb', '良好': '#16a34a', '一般': '#ca8a04',
                '较差': '#dc2626', '很差': '#dc2626'
            }
            cell_left = ParagraphStyle('TCL', fontName=font_name, fontSize=6, alignment=TA_LEFT, leading=8)
            cell_center = ParagraphStyle('TCC', fontName=font_name, fontSize=6, alignment=TA_CENTER, leading=8)
            cell_right = ParagraphStyle('TCR', fontName=font_name, fontSize=6, alignment=TA_RIGHT, leading=8)
            toc_header_bold = [Paragraph(f"<b>{h}</b>", cell_center if h in ("项目名","子表名") else cell_right if h in ("#","Cpk","PPM","均值","中位数","最大值","最小值") else cell_left) for h in ["#", "项目名", "子表名", "Cpk", "PPM", "LSL/USL", "均值", "中位数", "最大值", "最小值"]]
            toc_data = [toc_header_bold]
            for i, proj in enumerate(projects_list):
                s = proj['stats']
                cpk_color = toc_level_colors.get(proj['level'], '#334155')
                lsl_usl = f"{s['LSL']:.3f}" if s['LSL'] is not None else "-"
                lsl_usl += " / "
                lsl_usl += f"{s['USL']:.3f}" if s['USL'] is not None else "-"
                toc_data.append([
                    Paragraph(f"<b>{i + 1}</b>", cell_right),
                    Paragraph(f"<b>{str(proj['name'])[:20]}</b>", cell_left),
                    Paragraph(f"<b>{str(proj.get('sheet_name', 'Sheet1'))[:14]}</b>", cell_left),
                    Paragraph(f"<font color='{cpk_color}'><b>{proj['cpk_val']:.3f}</b></font>", cell_right),
                    Paragraph(f"<b>{int(s['PPM'])}</b>" if s['PPM'] is not None else "<b>-</b>", cell_right),
                    Paragraph(f"<b>{lsl_usl}</b>", cell_left),
                    Paragraph(f"<b>{s['Mean']:.3f}</b>", cell_right),
                    Paragraph(f"<b>{s['Median']:.3f}</b>", cell_right),
                    Paragraph(f"<b>{s['Max']:.3f}</b>", cell_right),
                    Paragraph(f"<b>{s['Min']:.3f}</b>", cell_right),
                ])
            col_w = [0.7*cm, 3.2*cm, 2.5*cm, 1.0*cm, 1.4*cm, 2.8*cm, 1.8*cm, 1.8*cm, 1.8*cm, 1.6*cm]
            t_toc = Table(toc_data, colWidths=col_w)
            t_toc.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('TOPPADDING', (0, 0), (-1, -1), 2),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
                ('GRID', (0, 0), (-1, -1), 0.3, colors.HexColor('#cbd5e1')),
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#e0e7ff')),
                ('FONTNAME', (0, 0), (-1, 0), font_name),
            ]))

            # ---- 判定标准脚标（页面最底部）----
            foot_style = ParagraphStyle(
                'Footnote', fontName=font_name, fontSize=7, leading=10,
                textColor=colors.HexColor('#64748b'), alignment=TA_CENTER
            )
            foot_text = (
                "判定标准："
                f"<font color='#2563eb'><b>优秀</b></font>(Cpk≥1.67)  "
                f"<font color='#16a34a'><b>良好</b></font>(1.33≤Cpk<1.67)  "
                f"<font color='#ca8a04'><b>一般</b></font>(1.00≤Cpk<1.33)  "
                f"<font color='#dc2626'><b>较差</b></font>(0.67≤Cpk<1.00)  "
                f"<font color='black'><b>很差</b></font>(Cpk<0.67)"
            )
            # 目录内容与脚标打包，脚标固定在页面底部
            toc_inner = Table([
                [Paragraph("CPK 过程能力汇总分析报告", title_style)],
                [Paragraph(f"时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}  |  项目数：{len(projects_list)}", sub_style)],
                [Paragraph("目 录", head_style)],
                [t_toc],
            ], colWidths=[usable_w])
            toc_inner.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('LEFTPADDING', (0, 0), (-1, -1), 0),
                ('RIGHTPADDING', (0, 0), (-1, -1), 0),
                ('TOPPADDING', (0, 0), (-1, -1), 0),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
            ]))
            toc_page = Table([
                [toc_inner],
                [Paragraph(foot_text, foot_style)],
            ], colWidths=[usable_w], rowHeights=[usable_h - 20, 20])
            toc_page.setStyle(TableStyle([
                ('VALIGN', (0, 0), (0, 0), 'TOP'),
                ('VALIGN', (0, -1), (0, -1), 'BOTTOM'),
                ('LEFTPADDING', (0, 0), (-1, -1), 0),
                ('RIGHTPADDING', (0, 0), (-1, -1), 0),
                ('TOPPADDING', (0, 0), (-1, -1), 0),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
            ]))
            story.append(toc_page)

            # ---- 详情页：切换到双 Frame 模板 ----
            story.append(NextPageTemplate('two_up'))
            story.append(PageBreak())

            n_proj = len(projects_list)
            for i, proj in enumerate(projects_list):
                card = self._build_batch_project_card(
                    proj=proj,
                    index=i + 1,
                    font_name=font_name,
                    card_w=usable_w,
                    card_h=half_h,
                    temp_img_paths=temp_img_paths,
                )
                story.append(card)

                # 偶数项（0-based 奇数）在底部 Frame；之后若还有项目则换页
                # 奇数项（0-based 偶数）在顶部 Frame；用 FrameBreak 进入下半页
                if i < n_proj - 1:
                    if i % 2 == 0:
                        story.append(FrameBreak())
                    else:
                        story.append(PageBreak())

            doc.build(story)
        finally:
            for path in temp_img_paths:
                if os.path.exists(path):
                    try:
                        os.unlink(path)
                    except Exception:
                        pass

    def _build_batch_project_card(self, proj, index, font_name, card_w, card_h, temp_img_paths):
        """
        构建单个项目卡片（半页）。
        布局：
          ┌─ 标题栏（项目名 + Cpk 徽章）─────────────────┐
          │  指标条（两行 8 格）                           │
          │  ┌─ 直方图 ──────┬─ 原始数据 ──────────────┐ │
          │  │               │                          │ │
          │  └───────────────┴──────────────────────────┘ │
          └──────────────────────────────────────────────┘
        内容按 card_h 预算尺寸，再包一层 KeepInFrame 兜底，避免溢出重叠。
        """
        s = proj['stats']
        data = proj['data']
        self.current_data = data
        self.current_stats = s

        def sf(v, fmt="{:.4f}"):
            if v is None:
                return "-"
            return fmt.format(v)

        cpk_val = s.get('Cpk')
        cpk_txt = f"{cpk_val:.3f}" if cpk_val is not None else "-"
        level = s.get('CPK_LEVEL', '-')
        ppm = s.get('PPM')
        sigma_lv = s.get('SigmaLevel', 0) or 0
        oos = s.get('OutOfSpecCount', 0)
        level_color = {
            '优秀': '#2563eb', '良好': '#16a34a', '一般': '#ca8a04',
            '较差': '#dc2626', '很差': '#dc2626'
        }.get(level, '#334155')

        # 唯一样式名，避免 reportlab 样式缓存冲突
        uid = f"p{index}"
        sty_title = ParagraphStyle(
            f'BT_{uid}', fontName=font_name, fontSize=11, leading=14,
            textColor=colors.black, alignment=TA_LEFT
        )
        sty_badge = ParagraphStyle(
            f'BB_{uid}', fontName=font_name, fontSize=8, leading=10,
            textColor=colors.white, alignment=TA_CENTER
        )
        sty_metric_lab = ParagraphStyle(
            f'BML_{uid}', fontName=font_name, fontSize=6, leading=7.5,
            textColor=colors.HexColor('#64748b'), alignment=TA_CENTER
        )
        sty_metric_val = ParagraphStyle(
            f'BMV_{uid}', fontName=font_name, fontSize=8, leading=10,
            textColor=colors.HexColor('#0f172a'), alignment=TA_CENTER
        )
        sty_sec = ParagraphStyle(
            f'BS_{uid}', fontName=font_name, fontSize=7, leading=9,
            textColor=colors.HexColor('#1d4ed8'), alignment=TA_LEFT
        )
        sty_note = ParagraphStyle(
            f'BN_{uid}', fontName=font_name, fontSize=5.5, leading=7,
            textColor=colors.HexColor('#94a3b8'), alignment=TA_CENTER
        )

        # ---- 高度预算（pt）----
        # 外层 Table 的 padding 会吃掉 2*pad，内部元素之和必须 ≤ card_h - 2*pad
        pad = 5
        header_h = 20
        sec_label_h = 11
        spacer_after_header = 3
        # 可用内容高度（扣除卡片上下内边距）
        content_budget = card_h - 2 * pad
        body_h = content_budget - header_h - spacer_after_header
        if body_h < 100:
            body_h = 100

        # 左右分栏：左图约 58%，右数据约 42%（图片横向拉伸）
        gutter = 6
        left_w = card_w * 0.58 - gutter / 2
        right_w = card_w * 0.42 - gutter / 2
        metrics_detail_h = 76
        chart_h = body_h - sec_label_h - metrics_detail_h - 4

        # ---- 1) 标题栏 ----
        sheet_info = f"  ·  {proj.get('sheet_name', '')}" if proj.get('sheet_name') else ""
        name_txt = str(proj.get('name', '未命名'))
        if len(name_txt) > 28:
            name_txt = name_txt[:27] + "…"
        title_para = Paragraph(f"<b>{index}. {name_txt}</b>{sheet_info}", sty_title)
        badge_para = Paragraph(
            f"<b>Cpk {cpk_txt}</b>  {level}",
            sty_badge
        )
        badge_w = 2.6 * cm
        header_tbl = Table(
            [[title_para, badge_para]],
            colWidths=[card_w - 2 * pad - badge_w, badge_w],
            rowHeights=[header_h]
        )
        header_tbl.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#eff6ff')),
            ('BACKGROUND', (1, 0), (1, 0), colors.HexColor(level_color)),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (1, 0), (1, 0), 'CENTER'),
            ('LEFTPADDING', (0, 0), (0, 0), 6),
            ('RIGHTPADDING', (0, 0), (0, 0), 4),
            ('LEFTPADDING', (1, 0), (1, 0), 4),
            ('RIGHTPADDING', (1, 0), (1, 0), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ('BOX', (0, 0), (-1, -1), 0.4, colors.HexColor('#93c5fd')),
        ]))

        # ---- 2) 质量指标详表（位于图片上方）----
        def metric_pair(label, value, val_color=None):
            vc = val_color or 'black'
            return Paragraph(
                f"<font color='#1e40af' size='9'><b>{label}: </b></font>"
                f"<font color='{vc}' size='8'><b>{value}</b></font>",
                ParagraphStyle(f'MP_{uid}', fontName=font_name, fontSize=8, leading=12, alignment=TA_LEFT)
            )

        md_w = left_w / 4.0
        md_data = [
            [
                metric_pair("LSL", sf(s['LSL'])),
                metric_pair("USL", sf(s['USL'])),
                metric_pair("N", f"{int(s['Count'])}"),
                metric_pair("超规", str(int(oos)), '#dc2626' if int(oos) > 0 else '#15803d'),
            ],
            [
                metric_pair("均值", sf(s['Mean'])),
                metric_pair("中位数", sf(s['Median'])),
                metric_pair("最小值", sf(s['Min'])),
                metric_pair("最大值", sf(s['Max'])),
            ],
            [
                metric_pair("标准差", sf(s['StdDev'])),
                metric_pair("极差", sf(s['Range'])),
                metric_pair("峰度", sf(s['Kurtosis'], "{:.3f}")),
                metric_pair("偏度", sf(s['Skewness'], "{:.3f}")),
            ],
            [
                metric_pair("CV", f"{s['CV']:.2f}%"),
                metric_pair("Cp", sf(s['Cp'])),
                metric_pair("Sigma", f"{sigma_lv:.2f}σ"),
                metric_pair("PPM", f"{int(ppm)}" if ppm is not None else "-"),
            ],
        ]
        md_tbl = Table(md_data, colWidths=[md_w] * 4, rowHeights=[metrics_detail_h / 4.0] * 4)
        md_tbl.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#f8fafc')),
            ('BOX', (0, 0), (-1, -1), 0.3, colors.HexColor('#cbd5e1')),
            ('INNERGRID', (0, 0), (-1, -1), 0.2, colors.HexColor('#e2e8f0')),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('TOPPADDING', (0, 0), (-1, -1), 1),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
            ('LEFTPADDING', (0, 0), (-1, -1), 1),
            ('RIGHTPADDING', (0, 0), (-1, -1), 1),
        ]))

        # ---- 4) 直方图 ----
        img_path = self._create_temp_chart_image(size='half')
        temp_img_paths.append(img_path)
        chart_img = Image(img_path, width=left_w - 2, height=chart_h - 2)
        chart_img.hAlign = 'CENTER'

        left_block = Table(
            [
                [Paragraph("<b>分布直方图</b>", sty_sec)],
                [md_tbl],
                [chart_img],
            ],
            colWidths=[left_w],
            rowHeights=[sec_label_h, metrics_detail_h, chart_h]
        )
        left_block.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('ALIGN', (0, 1), (0, 1), 'CENTER'),
            ('ALIGN', (0, 2), (0, 2), 'CENTER'),
            ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 3),
            ('TOPPADDING', (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
            ('BACKGROUND', (0, 2), (0, 2), colors.HexColor('#ffffff')),
            ('BOX', (0, 2), (0, 2), 0.3, colors.HexColor('#e2e8f0')),
        ]))

        # ---- 5) 原始数据表：最多 30 行 × 5 列 ----
        data_row_h = 8.5
        data_font = 5.8
        data_cols = 5
        max_data_rows = 30
        show_note = len(data) > max_data_rows * data_cols

        data_tbl = self._create_compact_data_table(
            data, font_name, right_w - 2,
            max_rows=max_data_rows, cols=data_cols,
            row_height=data_row_h, font_size=data_font,
        )

        right_block = Table(
            [
                [Paragraph("<b>原始数据</b>", sty_sec)],
                [data_tbl],
            ],
            colWidths=[right_w],
            rowHeights=[sec_label_h, body_h - sec_label_h]
        )
        right_block.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 3),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ('TOPPADDING', (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ]))

        body_tbl = Table(
            [[left_block, right_block]],
            colWidths=[left_w + gutter / 2, right_w + gutter / 2],
            rowHeights=[body_h]
        )
        body_tbl.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ('TOPPADDING', (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
            ('LINEAFTER', (0, 0), (0, 0), 0.5, colors.HexColor('#cbd5e1')),
        ]))

        # ---- 组装卡片 ----
        card_inner = Table(
            [
                [header_tbl],
                [Spacer(1, spacer_after_header)],
                [body_tbl],
            ],
            colWidths=[card_w - 2 * pad]
        )
        card_inner.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ('TOPPADDING', (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ]))

        card_outer = Table(
            [[card_inner]],
            colWidths=[card_w],
            rowHeights=[card_h]
        )
        card_outer.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#ffffff')),
            ('BOX', (0, 0), (-1, -1), 0.9, colors.HexColor('#64748b')),
            ('LEFTPADDING', (0, 0), (-1, -1), pad),
            ('RIGHTPADDING', (0, 0), (-1, -1), pad),
            ('TOPPADDING', (0, 0), (-1, -1), pad),
            ('BOTTOMPADDING', (0, 0), (-1, -1), pad),
        ]))

        # KeepInFrame 兜底：内容已按预算设计，正常不会 shrink；异常时缩小而非溢出重叠
        return KeepInFrame(
            card_w, card_h, [card_outer],
            mode='shrink', hAlign='LEFT', vAlign='TOP'
        )

    def _create_compact_data_table(self, data, font_name, width, max_count=120, cols=10,
                                  max_rows=None, row_height=9.2, font_size=6.0):
        """紧凑实测数据表。固定行高；优先按 max_rows 限制，避免撑破半页布局。"""
        d_list = data.tolist() if isinstance(data, np.ndarray) else list(data)
        if max_rows is not None:
            total_limit = cols * max_rows
        else:
            total_limit = max_count

        max_rows_eff = max(1, (total_limit + cols - 1) // cols)
        total_limit = min(total_limit, max_rows_eff * cols)
        subset = d_list[:total_limit]

        rows_data = []
        for i in range(0, len(subset), cols):
            row_slice = list(subset[i:i + cols])
            while len(row_slice) < cols:
                row_slice.append("")
            rows_data.append([
                f"{x:.3f}" if x != "" else "" for x in row_slice
            ])

        if not rows_data:
            return Paragraph(
                "无数据",
                ParagraphStyle('nd_empty', fontName=font_name, fontSize=7, leading=9)
            )

        n_rows = len(rows_data)
        col_w = width / cols
        # 行高至少能容纳字号 + 上下内边距，防止文字裁切/叠字
        safe_row_h = max(float(row_height), float(font_size) + 5.0)
        row_heights = [safe_row_h] * n_rows

        t = Table(rows_data, colWidths=[col_w] * cols, rowHeights=row_heights)
        style_cmds = [
            ('FONTNAME', (0, 0), (-1, -1), font_name),
            ('FONTSIZE', (0, 0), (-1, -1), font_size),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 1.0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 1.0),
            ('TOPPADDING', (0, 0), (-1, -1), 1.0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 1.0),
            ('GRID', (0, 0), (-1, -1), 0.25, colors.HexColor('#94a3b8')),
            ('BOX', (0, 0), (-1, -1), 0.45, colors.HexColor('#64748b')),
        ]
        for i_row in range(n_rows):
            bg = colors.HexColor('#f1f5f9') if i_row % 2 == 0 else colors.white
            style_cmds.append(('BACKGROUND', (0, i_row), (-1, i_row), bg))
        t.setStyle(TableStyle(style_cmds))

        if len(d_list) > total_limit:
            note = Paragraph(
                f"共 {len(d_list)} 条，显示前 {total_limit} 条",
                ParagraphStyle(
                    'data_note', fontName=font_name, fontSize=5.5,
                    alignment=TA_CENTER, leading=7,
                    textColor=colors.HexColor('#94a3b8')
                )
            )
            wrap = Table([[t], [Spacer(1, 1)], [note]], colWidths=[width])
            wrap.setStyle(TableStyle([
                ('LEFTPADDING', (0, 0), (-1, -1), 0),
                ('RIGHTPADDING', (0, 0), (-1, -1), 0),
                ('TOPPADDING', (0, 0), (-1, -1), 0),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ]))
            return wrap
        return t

    def _generate_single_pdf_logic(self, file_path):
        temp_img_paths = []
        try:
            doc = SimpleDocTemplate(
                file_path, pagesize=A4,
                rightMargin=1.0 * cm, leftMargin=1.0 * cm,
                topMargin=1.0 * cm, bottomMargin=1.0 * cm
            )
            story = []
            styles = getSampleStyleSheet()

            font_name = "Helvetica"
            try:
                if os.name == 'nt':
                    path = r"C:\Windows\Fonts\simhei.ttf"
                    if os.path.exists(path):
                        pdfmetrics.registerFont(TTFont('SimHei', path))
                        font_name = 'SimHei'
            except Exception:
                pass

            title_style = ParagraphStyle(
                'Title', parent=styles['Heading1'], fontName=font_name,
                fontSize=18, alignment=TA_CENTER, spaceAfter=5,
                textColor=colors.black, leading=22
            )
            sub_style = ParagraphStyle(
                'Sub', parent=styles['Normal'], fontName=font_name,
                fontSize=9, alignment=TA_CENTER, spaceAfter=8, textColor=colors.gray
            )
            head_style = ParagraphStyle(
                'Head', parent=styles['Heading3'], fontName=font_name,
                fontSize=12, spaceBefore=8, spaceAfter=4,
                textColor=colors.darkblue, leading=14
            )
            normal_style = ParagraphStyle(
                'Norm', parent=styles['Normal'], fontName=font_name,
                fontSize=9, leading=12, textColor=colors.black
            )

            story.append(Paragraph("CPK 过程能力分析报表", title_style))
            story.append(Paragraph(
                f"项目：{self.project_name} | 时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}",
                sub_style
            ))

            story.extend(self._create_stats_table_story(normal_style, font_name, compact=True))
            story.append(Spacer(1, 0.08 * inch))

            img_path = self._create_temp_chart_image(compact=True)
            temp_img_paths.append(img_path)
            story.append(Paragraph("分布直方图", head_style))
            img = Image(img_path, width=6.8 * inch, height=3.6 * inch)
            story.append(img)
            story.append(Spacer(1, 0.05 * inch))

            story.append(Paragraph("原始数据明细", head_style))
            story.extend(self._create_data_table_story(normal_style, font_name))

            doc.build(story)
        finally:
            for path in temp_img_paths:
                if os.path.exists(path):
                    try:
                        os.unlink(path)
                    except Exception:
                        pass

    def _create_stats_table_story(self, normal_style, font_name, compact=False):
        s = self.current_stats
        highlight_color = colors.Color(0.93, 0.96, 1.0)

        def safe_fmt(val, fmt_str="{:.4f}"):
            if val is None:
                return "-"
            return fmt_str.format(val)

        out_of_spec = s.get('OutOfSpecCount', 0)
        fs_main = 9 if compact else 10
        fs_small = 7.5 if compact else 9

        core_data = [
            [
                Paragraph("<b>USL</b><br/>" + safe_fmt(s['USL']), normal_style),
                Paragraph("<b>LSL</b><br/>" + safe_fmt(s['LSL']), normal_style),
                Paragraph("<b>Cpk</b><br/><font size=14 color='darkred'>" + safe_fmt(s['Cpk']) + "</font>", normal_style),
                Paragraph(
                    "<b>PPM</b><br/><font size=14 color='darkred'>" +
                    (f"{int(s['PPM'])}" if s['PPM'] is not None else "-") + "</font>",
                    normal_style
                )
            ],
            [
                Paragraph("均值：" + safe_fmt(s['Mean']), ParagraphStyle('Small', parent=normal_style, fontSize=fs_small)),
                Paragraph("中位数：" + safe_fmt(s['Median']), ParagraphStyle('Small', parent=normal_style, fontSize=fs_small)),
                Paragraph("标准差：" + safe_fmt(s['StdDev']), ParagraphStyle('Small', parent=normal_style, fontSize=fs_small)),
                Paragraph("CV：" + f"{s['CV']:.2f}%", ParagraphStyle('Small', parent=normal_style, fontSize=fs_small))
            ],
            [
                Paragraph("最大：" + safe_fmt(s['Max']), ParagraphStyle('Small', parent=normal_style, fontSize=fs_small)),
                Paragraph("最小：" + safe_fmt(s['Min']), ParagraphStyle('Small', parent=normal_style, fontSize=fs_small)),
                Paragraph("极差：" + safe_fmt(s['Range']), ParagraphStyle('Small', parent=normal_style, fontSize=fs_small)),
                Paragraph(
                    "偏度：" + f"{s['Skewness']:.2f}<br/>峰度：" + f"{s['Kurtosis']:.2f}",
                    ParagraphStyle('Small', parent=normal_style, fontSize=fs_small - 1)
                )
            ],
            [
                Paragraph("Cp：" + safe_fmt(s['Cp']), ParagraphStyle('Small', parent=normal_style, fontSize=fs_small)),
                Paragraph("Sigma：" + f"{s['SigmaLevel']:.2f}σ", ParagraphStyle('Small', parent=normal_style, fontSize=fs_small)),
                Paragraph("N：" + f"{int(s['Count'])}", ParagraphStyle('Small', parent=normal_style, fontSize=fs_small)),
                Paragraph("超规：<b>" + f"{int(out_of_spec)}" + "</b>", ParagraphStyle('Small', parent=normal_style, fontSize=fs_small))
            ]
        ]

        col_w = 4.0 * cm if compact else 4.2 * cm
        pad_big = 4 if compact else 8
        pad_small = 2 if compact else 4

        t_core = Table(core_data, colWidths=[col_w] * 4)
        t_core.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), highlight_color),
            ('BOX', (0, 0), (-1, 0), 1.0, colors.darkblue),
            ('INNERGRID', (0, 0), (-1, 0), 0.5, colors.white),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), pad_big),
            ('TOPPADDING', (0, 0), (-1, 0), pad_big),
            ('FONTNAME', (0, 0), (-1, 0), font_name),
            ('FONTSIZE', (0, 0), (-1, 0), fs_main),
            ('BACKGROUND', (0, 1), (-1, 3), colors.white),
            ('BOX', (0, 1), (-1, 3), 0.5, colors.lightgrey),
            ('ALIGN', (0, 1), (-1, 3), 'CENTER'),
            ('VALIGN', (0, 1), (-1, 3), 'MIDDLE'),
            ('BOTTOMPADDING', (0, 1), (-1, 3), pad_small),
            ('TOPPADDING', (0, 1), (-1, 3), pad_small),
            ('FONTNAME', (0, 1), (-1, 3), font_name),
            ('FONTSIZE', (0, 1), (-1, 3), fs_small),
            ('TEXTCOLOR', (0, 1), (-1, 3), colors.darkgrey),
        ]))
        return [t_core]

    def _create_temp_chart_image(self, compact=False, size=None):
        """生成直方图。size='half' 用于批量导出半页布局。"""
        if size == 'half':
            # 横向拉伸，适配左栏宽比例
            w, h = 6.8, 2.4
            dpi = 180
            title_fs, label_fs, tick_fs = 9, 7, 6
            lw = 1.6
            bins = 20
            tight_pad = 0.25
        elif compact:
            w, h = 7.0, 3.6
            dpi = 300
            title_fs, label_fs, tick_fs = 12, 9, 8
            lw = 2.5
            bins = 30
            tight_pad = 0.3
        else:
            w, h = 7.5, 4.0
            dpi = 300
            title_fs, label_fs, tick_fs = 12, 9, 8
            lw = 2.5
            bins = 30
            tight_pad = 0.3

        fig, ax = plt.subplots(figsize=(w, h), dpi=dpi)
        fig.patch.set_facecolor('white')
        ax.set_facecolor('white')

        data = self.current_data
        stats = self.current_stats
        if data is None or stats is None:
            raise ValueError("当前无有效数据可导出")

        mu, sigma = stats['Mean'], stats['StdDev']
        usl, lsl = stats['USL'], stats['LSL']
        font_cn = 'SimHei' if os.name == 'nt' else 'sans-serif'

        ax.hist(data, bins=bins, density=True,
                alpha=0.7, color='#3b82f6', edgecolor='white', linewidth=0.4)
        xmin, xmax = ax.get_xlim()
        base_span = 6 * sigma if sigma > 0 else 1.0
        plot_min = lsl - base_span * 0.2 if lsl is not None else min(xmin, mu - 4 * sigma)
        plot_max = usl + base_span * 0.2 if usl is not None else max(xmax, mu + 4 * sigma)

        x = np.linspace(plot_min, plot_max, 400)
        y = norm.pdf(x, mu, sigma)
        ax.plot(x, y, color='#dc2626', linewidth=lw)
        ax.fill_between(x, y, alpha=0.18, color='#dc2626')

        ymax = max(y) * 1.18 if len(y) > 0 else 1
        ax.set_ylim(0, ymax)

        # 规格线标签错开高度，避免 USL/LSL 重叠
        if usl is not None:
            ax.axvline(usl, c='#ef4444', ls='--', lw=1.4)
            ax.text(usl, ymax * 0.96, "USL", c='#ef4444', ha='center',
                    fontweight='bold', fontsize=tick_fs,
                    bbox=dict(boxstyle='round,pad=0.12', facecolor='white',
                              edgecolor='none', alpha=0.85))
        if lsl is not None:
            ax.axvline(lsl, c='#ef4444', ls='--', lw=1.4)
            ax.text(lsl, ymax * 0.82, "LSL", c='#ef4444', ha='center',
                    fontweight='bold', fontsize=tick_fs,
                    bbox=dict(boxstyle='round,pad=0.12', facecolor='white',
                              edgecolor='none', alpha=0.85))

        ax.axvline(mu, c='#16a34a', ls='-', lw=1.4, alpha=0.9)
        if size == 'half':
            ax.text(mu, ymax * 0.68, f"μ", c='#16a34a', ha='center',
                    fontsize=tick_fs, fontweight='bold',
                    bbox=dict(boxstyle='round,pad=0.1', facecolor='white',
                              edgecolor='none', alpha=0.8))

        cpk_v = stats.get('Cpk')
        title_str = f"Cpk = {cpk_v:.3f}" if cpk_v is not None else "Cpk = -"
        ax.set_title(title_str, fontsize=title_fs, fontweight='bold', pad=2, fontname=font_cn)
        ax.set_xlabel("测量值", fontsize=label_fs, color='black', labelpad=1, fontname=font_cn)
        ax.set_ylabel("密度", fontsize=label_fs, color='black', labelpad=1, fontname=font_cn)
        ax.tick_params(colors='black', labelsize=tick_fs)

        ax.grid(True, linestyle='--', alpha=0.25, color='gray')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_color('black')
        ax.spines['bottom'].set_color('black')

        fig.tight_layout(pad=tight_pad)

        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp_file:
            temp_img_path = tmp_file.name

        # half 模式不用 bbox_inches='tight'，保持画布比例稳定，避免 PDF 内再变形
        if size == 'half':
            fig.savefig(temp_img_path, dpi=dpi, facecolor='white')
        else:
            fig.savefig(temp_img_path, dpi=dpi, bbox_inches='tight', facecolor='white')
        plt.close(fig)
        return temp_img_path

    def _create_data_table_story(self, normal_style, font_name):
        d_list = self.current_data
        if isinstance(d_list, np.ndarray):
            d_list = d_list.tolist()

        cols = 10
        max_rows = 15
        rows_data = []

        total_limit = cols * max_rows
        subset = d_list[:total_limit] if len(d_list) > total_limit else d_list

        for i in range(0, len(subset), cols):
            row_slice = subset[i:i + cols]
            while len(row_slice) < cols:
                row_slice.append("")
            formatted_row = [f"{x:.3f}" if x != "" else "" for x in row_slice]
            rows_data.append(formatted_row)

        result_story = []
        if rows_data:
            col_width = (17.0 * cm) / cols
            t_data = Table(rows_data, colWidths=[col_width] * cols)

            data_table_style = [
                ('FONTNAME', (0, 0), (-1, -1), font_name),
                ('FONTSIZE', (0, 0), (-1, -1), 6.5),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 1),
                ('RIGHTPADDING', (0, 0), (-1, -1), 1),
                ('TOPPADDING', (0, 0), (-1, -1), 1),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
                ('GRID', (0, 0), (-1, -1), 0.25, colors.lightgrey),
            ]
            for i_row in range(len(rows_data)):
                bg_color = colors.Color(0.97, 0.97, 0.97) if i_row % 2 == 0 else colors.white
                data_table_style.append(('BACKGROUND', (0, i_row), (-1, i_row), bg_color))

            t_data.setStyle(TableStyle(data_table_style))
            result_story.append(t_data)

        if len(self.current_data) > total_limit:
            note_style = ParagraphStyle(
                'Note', parent=normal_style, fontSize=6,
                textColor=colors.gray, alignment=TA_CENTER, spaceBefore=2
            )
            result_story.append(Paragraph(
                f"... 共 {len(self.current_data)} 条，显示前 {total_limit} 条", note_style
            ))

        return result_story


if __name__ == "__main__":
    root = tk.Tk()
    app = CpkApp(root)
    root.mainloop()
