"""CPK 统计分析工具 Pro —— Fluent UI 控件库版本（PySide6）。"""
import math
import sys
import ctypes
import re
import os
import tempfile
import traceback
from datetime import datetime
from pathlib import Path

_KJ_PATH = r"D:\OneDrive\Application\CODE\0_组件库\控件库"
if _KJ_PATH not in sys.path:
    sys.path.insert(0, _KJ_PATH)

import matplotlib
matplotlib.use('Agg')
import numpy as np
import matplotlib.pyplot as plt
from scipy.stats import norm, skew, kurtosis

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

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False

from PySide6.QtCore import QPointF, QRectF, Qt, QRect, Signal
from PySide6.QtGui import (QColor, QFont, QFontMetrics, QGuiApplication,
                           QKeySequence, QPainter, QPainterPath, QPen,
                           QPolygonF, QShortcut)
from PySide6.QtWidgets import (QApplication, QFrame, QGridLayout, QHBoxLayout,
                               QLabel, QSizePolicy, QTreeWidgetItem,
                               QVBoxLayout, QWidget, QFileDialog)

from fluent.window import FluentApp, FluentWindow, apply_theme
from fluent.widgets import (FluentButton, FluentCard, FluentInputDialog,
                            FluentLineEdit, FluentMessageBox, FluentTabWidget,
                            FluentTextEdit, FluentTreeView)

plt.style.use('dark_background')
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False

C_BG = "#202020"
C_PANEL = "#1B1B1B"
C_CARD = "#2B2B2B"
C_BORDER = "#3B3B3B"
C_FG = "#FFFFFF"
C_FG_MUTED = "#9D9D9D"
C_FG_DIM = "#6E6E6E"
C_ACCENT = "#0067C0"
C_ACCENT_LIGHT = "#4CC2FF"
C_SUCCESS = "#00B050"
C_WARNING = "#FFB900"
C_DANGER = "#E74856"
C_GRID = "#3A3A3A"

LEVEL_COLORS = {
    "优秀": C_SUCCESS,
    "良好": "#7CBA5A",
    "一般": C_WARNING,
    "较差": "#FF6F00",
    "很差": C_DANGER,
}
LEVEL_BGS = {
    "优秀": "#123F1E",
    "良好": "#243018",
    "一般": "#3A2E0F",
    "较差": "#3A2410",
    "很差": "#3A1717",
}

FONT_UI = 'Microsoft YaHei UI'
FONT_MONO = 'Consolas'


def _nice_ticks(vmin, vmax, count=5):
    if vmax <= vmin:
        vmin, vmax = 0.0, 1.0
    span = vmax - vmin
    step_raw = span / count
    if step_raw <= 0:
        return [vmin]
    mag = 10 ** math.floor(math.log10(step_raw))
    norm = step_raw / mag
    if norm < 1.5:
        step = 1
    elif norm < 3:
        step = 2
    elif norm < 7:
        step = 5
    else:
        step = 10
    step *= mag
    start = math.ceil(vmin / step) * step
    ticks = []
    v = start
    while v <= vmax + step * 1e-9:
        ticks.append(v)
        v += step
    return ticks


def _fmt_num(v):
    if abs(v) >= 1000 or abs(v) < 0.01:
        return f"{v:g}"
    if v == int(v):
        return str(int(v))
    return f"{v:.1f}"


class HistogramWidget(QWidget):
    """纯 QPainter 自绘分布直方图：柱状频次 + 正态曲线 + 规格限 + 均值。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumSize(320, 240)
        self._stats = None
        self._title = ""
        self._subtitle = ""
        self._bars = []
        self._bar_edges = []
        self._curve = []
        self._x_lo = 0.0
        self._x_hi = 1.0
        self._y_max = 1.0
        self._markers = []

    def reset(self):
        self._stats = None
        self._bars = []
        self._bar_edges = []
        self._curve = []
        self._markers = []
        self._title = ""
        self._subtitle = ""
        self.update()

    def set_data(self, data, stats):
        self._stats = stats
        mu = stats['Mean']
        sigma = stats['StdDev']
        usl = stats['USL']
        lsl = stats['LSL']
        if sigma is None or sigma <= 1e-9:
            self.reset()
            return
        hist, edges = np.histogram(data, bins=30, density=True)
        self._bars = list(hist)
        self._bar_edges = list(edges)
        xmin, xmax = float(np.min(data)), float(np.max(data))
        base_span = 6.0 * sigma
        lo = (lsl - base_span * 0.2) if lsl is not None else min(xmin, mu - 4 * sigma)
        hi = (usl + base_span * 0.2) if usl is not None else max(xmax, mu + 4 * sigma)
        if hi <= lo:
            hi = lo + 1.0
        self._x_lo = float(lo)
        self._x_hi = float(hi)
        xs = np.linspace(lo, hi, 300)
        ys = norm.pdf(xs, mu, sigma)
        self._curve = [(float(x), float(y)) for x, y in zip(xs, ys)]
        self._y_max = max(float(np.max(hist)), 1e-9)
        if len(ys):
            self._y_max = max(self._y_max, float(np.max(ys)) * 1.25)
        self._markers = []
        if usl is not None:
            self._markers.append((float(usl), "USL"))
        if lsl is not None:
            self._markers.append((float(lsl), "LSL"))
        self._markers.append((float(mu), "μ"))
        cpk = stats.get('Cpk')
        level = stats.get('CPK_LEVEL', '')
        self._title = f"Cpk = {cpk:.3f}" if cpk is not None else "Cpk = -"
        self._subtitle = level
        self.update()

    def _tick_font(self):
        f = QFont(self.font())
        f.setPointSizeF(max(7.0, f.pointSizeF() - 0.5))
        return f

    def _draw_title(self, p):
        f = QFont(self.font())
        f.setBold(True)
        f.setPointSizeF(f.pointSizeF() + 1.5)
        p.setFont(f)
        p.setPen(QColor(C_FG))
        p.drawText(QRectF(0, 6, self.width(), 22),
                   Qt.AlignmentFlag.AlignHCenter, self._title)
        if self._subtitle:
            p.setFont(self._tick_font())
            p.setPen(QColor(C_FG_MUTED))
            p.drawText(QRectF(0, 24, self.width(), 16),
                       Qt.AlignmentFlag.AlignHCenter, self._subtitle)

    def _draw_axes(self, p, plot):
        f = self._tick_font()
        p.setFont(f)
        fm = QFontMetrics(f)
        for t in _nice_ticks(0, self._y_max, 5):
            y = plot.bottom() - (t / self._y_max) * plot.height()
            p.setPen(QPen(QColor(C_GRID), 1))
            p.drawLine(QPointF(plot.left(), y), QPointF(plot.right(), y))
            p.setPen(QColor(C_FG_MUTED))
            p.drawText(QRectF(0, y - 9, plot.left() - 8, 18),
                       Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter,
                       _fmt_num(t))
        for t in _nice_ticks(self._x_lo, self._x_hi, 5):
            x = plot.left() + (t - self._x_lo) / (self._x_hi - self._x_lo) * plot.width()
            p.setPen(QPen(QColor(C_GRID), 1))
            p.drawLine(QPointF(x, plot.bottom()), QPointF(x, plot.top()))
            p.setPen(QColor(C_FG_MUTED))
            p.drawText(QRectF(x - 44, plot.bottom() + 6, 88, 18),
                       Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop,
                       _fmt_num(t))
        _ = fm

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.fillRect(self.rect(), QColor(C_BG))
        if self._stats is None:
            f = QFont(self.font())
            f.setPointSizeF(f.pointSizeF() + 1)
            p.setFont(f)
            p.setPen(QColor(C_FG_DIM))
            p.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter,
                       "等待数据…\n在左侧输入测量值或导入 Excel")
            p.end()
            return
        r = self.rect()
        fm = QFontMetrics(self._tick_font())
        left = fm.horizontalAdvance("-888.8") + 30
        right = 14
        top = 44
        bottom = fm.height() + 30
        plot = QRectF(r.left() + left, r.top() + top,
                      r.width() - left - right, r.height() - top - bottom)
        if plot.width() < 40 or plot.height() < 40:
            p.end()
            return
        self._draw_title(p)
        self._draw_axes(p, plot)
        span = self._x_hi - self._x_lo
        for i, h in enumerate(self._bars):
            e0 = self._bar_edges[i]
            e1 = self._bar_edges[i + 1]
            x0 = plot.left() + (e0 - self._x_lo) / span * plot.width()
            x1 = plot.left() + (e1 - self._x_lo) / span * plot.width()
            y0 = plot.bottom() - (h / self._y_max) * plot.height()
            c = QColor(C_ACCENT_LIGHT)
            c.setAlpha(150)
            p.setPen(Qt.PenStyle.NoPen)
            p.setBrush(c)
            p.drawRoundedRect(QRectF(x0, y0, max(1.0, x1 - x0),
                                     plot.bottom() - y0), 1, 1)
        if len(self._curve) >= 2:
            pts = QPolygonF()
            for x, y in self._curve:
                px = plot.left() + (x - self._x_lo) / span * plot.width()
                py = plot.bottom() - (y / self._y_max) * plot.height()
                pts.append(QPointF(px, py))
            path = QPainterPath(pts.first())
            for pt in pts:
                path.lineTo(pt)
            path.lineTo(pts.last().x(), plot.bottom())
            path.lineTo(pts.first().x(), plot.bottom())
            path.closeSubpath()
            fc = QColor(C_SUCCESS)
            fc.setAlpha(45)
            p.fillPath(path, fc)
            pen = QPen(QColor(C_SUCCESS), 2.0)
            p.setPen(pen)
            p.setBrush(Qt.BrushStyle.NoBrush)
            p.drawPolyline(pts)
        for x, label in self._markers:
            px = plot.left() + (x - self._x_lo) / span * plot.width()
            if px < plot.left() or px > plot.right():
                continue
            is_spec = label in ("USL", "LSL")
            color = QColor(C_DANGER) if is_spec else QColor(C_WARNING)
            pen = QPen(color, 1.6)
            pen.setStyle(Qt.PenStyle.DashLine if is_spec else Qt.PenStyle.SolidLine)
            p.setPen(pen)
            p.drawLine(QPointF(px, plot.top()), QPointF(px, plot.bottom()))
            p.setFont(self._tick_font())
            p.setPen(color)
            label_y = plot.top() + 8 if label != "LSL" else plot.bottom() - 12
            p.drawText(QRectF(px - 40, label_y - 9, 80, 18),
                       Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter,
                       label)
        p.save()
        p.translate(14, plot.top() + plot.height() / 2)
        p.rotate(-90)
        p.setFont(self._tick_font())
        p.setPen(QColor(C_FG_MUTED))
        p.drawText(QRectF(-60, -10, 120, 20),
                   Qt.AlignmentFlag.AlignCenter, "概率密度")
        p.restore()
        p.setFont(self._tick_font())
        p.setPen(QColor(C_FG_MUTED))
        p.drawText(QRectF(plot.left(), plot.bottom() + 8, plot.width(), 18),
                   Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop,
                   "测量值")
        p.end()


class MetricChip(QFrame):
    """指标小卡片：标签 + 大号数值。"""

    def __init__(self, label='', value='-', parent=None):
        super().__init__(parent)
        self.setObjectName("MetricChip")
        self.setStyleSheet(f"#MetricChip {{ background-color: {C_CARD};"
                           f" border: 1px solid {C_BORDER}; border-radius: 8px; }}")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(8, 5, 8, 5)
        lay.setSpacing(1)
        l = QLabel(label)
        l.setStyleSheet("color: #9D9D9D; font-size: 10px; background: transparent;")
        lay.addWidget(l)
        self.value_lbl = QLabel(value)
        self._set_value_style(C_FG)
        lay.addWidget(self.value_lbl)
        self.setSizePolicy(QSizePolicy.Policy.Expanding,
                           QSizePolicy.Policy.Expanding)

    def _set_value_style(self, color):
        self.value_lbl.setStyleSheet(
            f"color: {color}; font-size: 13px; font-weight: 700;"
            " background: transparent;")

    def set_value(self, text, color=None):
        self.value_lbl.setText(text)
        self._set_value_style(color or C_FG)


class ClickableLabel(QLabel):
    """带悬停样式与点击信号的文本标签。"""

    clicked = Signal()

    def __init__(self, text='', idle='', hover='', parent=None):
        super().__init__(text, parent)
        self._idle = idle
        self._hover = hover
        self.setStyleSheet(idle)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

    def enterEvent(self, event):
        if self._hover:
            self.setStyleSheet(self._hover)
        super().enterEvent(event)

    def leaveEvent(self, event):
        if self._idle:
            self.setStyleSheet(self._idle)
        super().leaveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()
        super().mouseReleaseEvent(event)


def show_msg(parent, title, msg, is_error=True):
    if is_error:
        FluentMessageBox.error(parent, title, msg)
    else:
        FluentMessageBox.success(parent, title, msg)


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


class CpkApp(FluentWindow):
    def __init__(self):
        super().__init__(title="CPK 统计分析工具 Pro", width=1360, height=820)
        self.root = self
        self.setMinimumSize(1024, 640)

        self.app_dir = os.path.dirname(os.path.abspath(__file__))

        self.current_data = None
        self.current_stats = None
        self.current_usl = None
        self.current_lsl = None
        self.project_name = ""

        self.excel_projects = []
        self.current_excel_index = -1
        self._debug_tab_active = False

        self.setup_ui()
        QShortcut(QKeySequence("Ctrl+Shift+D"), self,
                  activated=self._show_debug_prompt)

    def _set_dark_titlebar(self):
        try:
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            hwnd = int(self.winId())
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE,
                ctypes.byref(ctypes.c_int(1)),
                ctypes.sizeof(ctypes.c_int(1)))
        except Exception:
            pass

    def _section_label(self, text):
        lbl = QLabel(text)
        lbl.setStyleSheet("color: #FFFFFF; font-size: 13px; font-weight: 700;")
        return lbl

    def _caption_label(self, text):
        lbl = QLabel(text)
        lbl.setStyleSheet("color: #6E6E6E; font-size: 10px;")
        return lbl

    def _danger_label(self, text):
        lbl = QLabel(text)
        lbl.setStyleSheet("color: #E74856; font-size: 11px;")
        return lbl

    def setup_ui(self):
        central = QWidget()
        root_lay = QVBoxLayout(central)
        root_lay.setContentsMargins(10, 10, 10, 10)
        root_lay.setSpacing(10)

        self._build_topbar()
        root_lay.addWidget(self._topbar)

        main = QHBoxLayout()
        main.setSpacing(10)

        left = QFrame()
        left.setObjectName("LeftPanel")
        left.setStyleSheet(
            f"#LeftPanel {{ background-color: {C_PANEL};"
            f" border: 1px solid {C_BORDER}; border-radius: 10px; }}")

        center = QWidget()

        right = QFrame()
        right.setObjectName("RightPanel")
        right.setStyleSheet(
            f"#RightPanel {{ background-color: {C_PANEL};"
            f" border: 1px solid {C_BORDER}; border-radius: 10px; }}")

        main.addWidget(left, 1)
        main.addWidget(center, 2)
        main.addWidget(right, 1)
        root_lay.addLayout(main, 1)

        self.setCentralWidget(central)

        self.init_left_panel(left)
        self.init_stats_panel(right)
        self.init_chart_panel(center)

    def _build_topbar(self):
        bar = QFrame()
        bar.setObjectName("TopBar")
        bar.setStyleSheet(
            f"#TopBar {{ background-color: {C_PANEL};"
            f" border: 1px solid {C_BORDER}; border-radius: 10px; }}")
        bar.setFixedHeight(44)
        lay = QHBoxLayout(bar)
        lay.setContentsMargins(14, 4, 14, 4)
        lay.setSpacing(10)

        logo = QLabel("CPK")
        logo.setStyleSheet(
            "background-color: #0067C0; color: #FFFFFF; font-weight: 700;"
            " padding: 4px 10px; border-radius: 6px;")
        lay.addWidget(logo)

        titles = QVBoxLayout()
        titles.setSpacing(0)
        t = QLabel("过程能力分析")
        t.setStyleSheet("color: #FFFFFF; font-size: 13px; font-weight: 700;")
        s = QLabel("Statistical Process Capability · Pro")
        s.setStyleSheet("color: #6E6E6E; font-size: 9px;")
        titles.addWidget(t)
        titles.addWidget(s)
        lay.addLayout(titles)

        lay.addStretch(1)

        self._status_chip = QLabel("就绪")
        self._status_chip.setStyleSheet(self._status_style('ok'))
        lay.addWidget(self._status_chip)

        idle = "color: #9D9D9D; padding: 4px 10px; background-color: #2B2B2B; border-radius: 6px; font-size: 11px;"
        hover = "color: #FFFFFF; padding: 4px 10px; background-color: #3F3F3F; border-radius: 6px; font-size: 11px;"
        about = ClickableLabel("关于", idle=idle, hover=hover)
        about.clicked.connect(self.show_about)
        lay.addWidget(about)

        self._topbar = bar

    def _status_style(self, kind):
        colors = {
            'ok': ('#123F1E', C_SUCCESS),
            'warn': ('#3A2E0F', C_WARNING),
            'err': ('#3A1717', C_DANGER),
            'info': ('#1A2C44', C_ACCENT_LIGHT),
        }
        bg, fg = colors.get(kind, colors['ok'])
        return (f"background-color: {bg}; color: {fg}; border-radius: 10px;"
                " padding: 4px 12px; font-size: 12px;")

    def set_status(self, text, kind='ok'):
        self._status_chip.setText(text)
        self._status_chip.setStyleSheet(self._status_style(kind))

    def init_left_panel(self, parent):
        lay = QVBoxLayout(parent)
        lay.setContentsMargins(12, 10, 12, 10)
        lay.setSpacing(8)

        proj_card = FluentCard()
        pl = QLabel("项目名称")
        pl.setStyleSheet("color: #9D9D9D; font-size: 12px;")
        self.inp_project = FluentLineEdit("未命名项目")
        proj_row = QWidget()
        pr = QHBoxLayout(proj_row)
        pr.setContentsMargins(0, 0, 0, 0)
        pr.setSpacing(10)
        pr.addWidget(pl)
        pr.addWidget(self.inp_project, 1)
        proj_card.addWidget(proj_row)
        lay.addWidget(proj_card)

        self.main_notebook = FluentTabWidget()
        t1 = QWidget()
        t3 = QWidget()
        self.main_notebook.addTab(t1, "数据分析")
        self.main_notebook.addTab(t3, "Excel 导入")
        self.setup_tab1(t1)
        self.setup_tab3(t3)
        self._notebook = self.main_notebook
        lay.addWidget(self.main_notebook, 1)

        if not PANDAS_AVAILABLE:
            lay.addWidget(self._danger_label("缺少 pandas/openpyxl，Excel 导入不可用"))

        export_card = FluentCard()
        el = QLabel("导出报告")
        el.setStyleSheet("color: #9D9D9D; font-size: 12px;")
        export_card.addWidget(el)
        self.btn_export = FluentButton("导出当前 PDF", variant="accent")
        self.btn_export.clicked.connect(self.export_report)
        self.btn_batch_export = FluentButton("合并全部 PDF")
        self.btn_batch_export.clicked.connect(self.export_merged_report)
        btn_row = QWidget()
        br = QHBoxLayout(btn_row)
        br.setContentsMargins(0, 0, 0, 0)
        br.setSpacing(8)
        br.addWidget(self.btn_export, 1)
        br.addWidget(self.btn_batch_export, 1)
        export_card.addWidget(btn_row)
        lay.addWidget(export_card)

        if not REPORTLAB_AVAILABLE:
            self.btn_export.setEnabled(False)
            self.btn_export.setText("缺少 reportlab")
            self.btn_batch_export.setEnabled(False)
            self.btn_batch_export.setText("缺少 reportlab")
            lay.addWidget(self._caption_label("pip install reportlab"))

        self.main_notebook.currentChanged.connect(self.on_tab_changed)

        self._debug_btn = ClickableLabel(
            "调试模式",
            idle="color: #6E6E6E; font-size: 11px; padding: 4px;",
            hover="color: #9D9D9D; font-size: 11px; padding: 4px;")
        self._debug_btn.clicked.connect(self._toggle_debug)
        lay.addWidget(self._debug_btn, 0, Qt.AlignmentFlag.AlignHCenter)

    def create_input(self, layout, label, default=""):
        return self._single_input(layout, label, default)

    def _single_input(self, layout, label, default=""):
        wrap = QWidget()
        vl = QVBoxLayout(wrap)
        vl.setContentsMargins(0, 0, 0, 0)
        vl.setSpacing(2)
        lbl = QLabel(label)
        lbl.setStyleSheet("color: #9D9D9D; font-size: 11px;")
        vl.addWidget(lbl)
        e = FluentLineEdit()
        e.setMinimumHeight(28)
        if default:
            e.setText(str(default))
        vl.addWidget(e)
        layout.addWidget(wrap, 1)
        return e

    def _pair_inputs(self, layout, spec1, spec2):
        wrap = QWidget()
        hl = QHBoxLayout(wrap)
        hl.setContentsMargins(0, 0, 0, 0)
        hl.setSpacing(8)
        e1 = self._single_input(hl, spec1[0], spec1[1] if len(spec1) > 1 else "")
        e2 = self._single_input(hl, spec2[0], spec2[1] if len(spec2) > 1 else "")
        layout.addWidget(wrap)
        return e1, e2

    def create_btn_bar(self, layout, cmd1, cmd2, lbl1):
        box = QHBoxLayout()
        box.setSpacing(8)
        b1 = FluentButton(lbl1, variant="accent")
        b1.setMinimumHeight(30)
        b1.clicked.connect(cmd1)
        b2 = FluentButton("清空")
        b2.setMinimumHeight(30)
        b2.clicked.connect(cmd2)
        box.addWidget(b1, 3)
        box.addWidget(b2, 1)
        layout.addLayout(box)

    def setup_tab1(self, f):
        lay = QVBoxLayout(f)
        lay.setContentsMargins(12, 8, 12, 8)
        lay.setSpacing(6)
        lay.addWidget(self._section_label("规格限"))
        self.inp_an_usl, self.inp_an_lsl = self._pair_inputs(
            lay, ("上限 USL",), ("下限 LSL",))
        lay.addSpacing(2)
        mh = QWidget()
        mhl = QHBoxLayout(mh)
        mhl.setContentsMargins(0, 0, 0, 0)
        mhl.setSpacing(8)
        mhl.addWidget(self._section_label("测量数据"))
        mhl.addWidget(self._caption_label("支持空格 / 换行 / 逗号分隔"))
        mhl.addStretch(1)
        lay.addWidget(mh)
        self.txt_data = FluentTextEdit(placeholder="输入测量数据…")
        lay.addWidget(self.txt_data, 1)
        self.create_btn_bar(lay, self.on_analyze, self.on_clear_tab1, "开始分析")

    def setup_tab2(self, f):
        lay = QVBoxLayout(f)
        lay.setContentsMargins(12, 8, 12, 8)
        lay.setSpacing(6)
        lay.addWidget(self._section_label("模拟参数"))
        self.inp_sim_usl, self.inp_sim_lsl = self._pair_inputs(
            lay, ("上限 USL",), ("下限 LSL",))
        self.inp_sim_cpk, self.inp_sim_mean = self._pair_inputs(
            lay, ("目标 Cpk", "1.33"), ("目标均值", "10.0"))
        self.inp_sim_cnt, self.inp_sim_prec = self._pair_inputs(
            lay, ("数量", "50"), ("小数精度", "3"))
        self.create_btn_bar(lay, self.on_simulate, self.on_clear_tab2, "生成数据")
        lay.addSpacing(2)
        lay.addWidget(self._section_label("结果预览"))
        self.txt_sim = FluentTextEdit()
        self.txt_sim.setReadOnly(True)
        self.txt_sim.setMinimumHeight(110)
        lay.addWidget(self.txt_sim, 1)
        copy_btn = FluentButton("复制结果")
        copy_btn.setMinimumHeight(30)
        copy_btn.clicked.connect(self.on_copy)
        lay.addWidget(copy_btn)

    def setup_tab3(self, f):
        lay = QVBoxLayout(f)
        lay.setContentsMargins(8, 8, 8, 8)
        lay.setSpacing(6)

        row = QHBoxLayout()
        row.setSpacing(8)
        imp = FluentButton("导入 Excel", variant="accent")
        imp.setMinimumHeight(30)
        imp.clicked.connect(self.load_excel_file)
        clr = FluentButton("清空列表", variant="danger")
        clr.setMinimumHeight(30)
        clr.clicked.connect(self.clear_excel_data)
        hint = QLabel("点击列表项查看分析")
        hint.setStyleSheet("color: #6E6E6E; font-size: 11px;")
        row.addWidget(imp)
        row.addWidget(clr)
        row.addStretch(1)
        row.addWidget(hint)
        lay.addLayout(row)

        self.tree_projects = FluentTreeView(headers=["项目", "子表", "Cpk", "等级"])
        self.tree_projects.header().setStyleSheet(
            "QHeaderView::section { background-color: #1B1B1B; color: #9D9D9D;"
            " border: none; border-bottom: 1px solid #3B3B3B; padding: 6px 8px;"
            " font-weight: 600; }")
        self.tree_projects.setColumnWidth(0, 120)
        self.tree_projects.setColumnWidth(1, 50)
        self.tree_projects.setColumnWidth(2, 50)
        self.tree_projects.setColumnWidth(3, 50)
        self.tree_projects.setRootIsDecorated(False)
        self.tree_projects.currentItemChanged.connect(
            lambda cur, prev: self.on_excel_item_select())
        lay.addWidget(self.tree_projects, 1)

    def init_stats_panel(self, parent):
        lay = QVBoxLayout(parent)
        lay.setContentsMargins(10, 10, 10, 10)
        lay.setSpacing(6)

        self.lbl_proj_display = QLabel("未命名项目")
        self.lbl_proj_display.setWordWrap(True)
        self.lbl_proj_display.setStyleSheet(
            "color: #4CC2FF; font-size: 14px; font-weight: 700;")
        lay.addWidget(self.lbl_proj_display)
        lay.addWidget(self._caption_label("详细质量指标"))

        hero = QFrame()
        hero.setObjectName("HeroCard")
        hero.setStyleSheet(
            f"#HeroCard {{ background-color: {C_CARD};"
            f" border: 1px solid {C_BORDER}; border-radius: 10px; }}")
        hl = QVBoxLayout(hero)
        hl.setContentsMargins(14, 8, 14, 8)
        hl.setSpacing(0)
        l0 = QLabel("Cpk")
        l0.setStyleSheet("color: #9D9D9D; font-size: 10px;")
        hl.addWidget(l0)
        hrow = QHBoxLayout()
        self.hero_cpk = QLabel("—")
        self.hero_cpk.setStyleSheet(
            "color: #FFFFFF; font-size: 26px; font-weight: 700;")
        hrow.addWidget(self.hero_cpk)
        hrow.addStretch(1)
        self.hero_level = QLabel("待分析")
        self._set_hero_level_style("待分析")
        hrow.addWidget(self.hero_level, 0, Qt.AlignmentFlag.AlignVCenter)
        hl.addLayout(hrow)
        lay.addWidget(hero)

        grid = QGridLayout()
        grid.setSpacing(5)
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
        self.stat_labels = {}
        for i, (key, label) in enumerate(fields):
            r, c = divmod(i, 3)
            chip = MetricChip(label, "-")
            grid.addWidget(chip, r, c)
            self.stat_labels[key] = chip
        for c in range(3):
            grid.setColumnStretch(c, 1)
        for i in range(6):
            grid.setRowStretch(i, 1)
        lay.addLayout(grid)

        self.stat_labels["Cpk"] = None
        self.stat_labels["CPK_LEVEL"] = None

    def _set_hero_level_style(self, text, fg=None, bg=None):
        self.hero_level.setText(text)
        self.hero_level.setStyleSheet(
            f"color: {fg or C_FG_MUTED}; background-color: {bg or C_CARD};"
            " border-radius: 10px; padding: 3px 10px;"
            " font-size: 11px; font-weight: 600;")

    def init_chart_panel(self, parent):
        lay = QVBoxLayout(parent)
        lay.setContentsMargins(14, 10, 14, 10)
        lay.setSpacing(6)
        hrow = QHBoxLayout()
        hrow.addWidget(self._section_label("分布直方图"))
        hrow.addWidget(self._caption_label("正态拟合 · 规格限 · 均值"))
        hrow.addStretch(1)
        lay.addLayout(hrow)
        self.hist_widget = HistogramWidget()
        lay.addWidget(self.hist_widget, 1)

    def _current_tab_text(self):
        return self.main_notebook.tabText(self.main_notebook.currentIndex())

    def on_tab_changed(self, index):
        selected_tab = self.main_notebook.tabText(index)
        if "Excel" in selected_tab:
            self.btn_export.setEnabled(
                bool(self.excel_projects) and self.current_excel_index != -1)
        else:
            self.btn_export.setEnabled(self.current_stats is not None)

    def get_val(self, entry, is_int=False, allow_empty=False):
        val_str = entry.text().strip()
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
            show_msg(self, "输入错误", "规格值必须是数字")
            return
        if usl is None and lsl is None:
            show_msg(self, "缺失规格", "请至少输入一个规格限")
            return

        raw = self.txt_data.toPlainText()
        nums = re.findall(r"[-+]?\d*\.?\d+|\d+", raw)
        try:
            data = np.array([float(x) for x in nums])
            if len(data) < 2:
                raise ValueError
        except Exception:
            show_msg(self, "数据错误", "请检查输入数据")
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
            show_msg(self, "输入错误", "请检查数值格式")
            return
        if usl is None and lsl is None:
            show_msg(self, "缺失规格", "请至少输入一个规格限")
            return

        data = CpkCalculator.simulate(cpk, mean, usl, lsl, cnt, max(0, min(prec, 10)))
        if data is None:
            show_msg(self, "错误", "无法生成数据")
            return

        fmt = f"{{:.{prec}f}}"
        self.txt_sim.clear()
        self.txt_sim.setPlainText("\n".join([fmt.format(x) for x in data]))
        self.process_result(data, usl, lsl)

    def process_result(self, data, usl, lsl, project_name=None):
        stats = CpkCalculator.calculate(data, usl, lsl)
        if "Error" in stats:
            show_msg(self, "计算错误", stats["Error"])
            self.set_status("计算失败", 'err')
            return

        self.current_data = data
        self.current_stats = stats
        self.current_usl = usl
        self.current_lsl = lsl
        self.update_stats_display(stats, project_name)
        self.draw_chart(data, stats)
        if "Excel" not in self._current_tab_text():
            self.btn_export.setEnabled(True)

    def draw_chart(self, data, stats):
        self.hist_widget.set_data(data, stats)

    def reset_chart(self):
        self.hist_widget.reset()
        for chip in self.stat_labels.values():
            if chip is not None:
                chip.set_value("-", C_FG_DIM)
        self.hero_cpk.setText("—")
        self.hero_cpk.setStyleSheet(
            "color: #FFFFFF; font-size: 26px; font-weight: 700;")
        self._set_hero_level_style("待分析")
        self.lbl_proj_display.setText("未命名项目")
        self.current_data = None
        self.current_stats = None
        self.set_status("就绪", 'ok')

    def update_stats_display(self, stats, project_name=None):
        if project_name:
            pname = project_name
        else:
            pname = self.inp_project.text().strip() or "未命名项目"

        self.lbl_proj_display.setText(pname)
        if not project_name:
            self.project_name = pname

        def fmt_val(v, precision=4):
            if v is None:
                return "N/A"
            return f"{v:.{precision}f}"

        cpk = stats.get('Cpk')
        level = stats.get('CPK_LEVEL', '')
        if cpk is not None:
            self.hero_cpk.setText(f"{cpk:.3f}")
            self.hero_cpk.setStyleSheet(
                f"color: {LEVEL_COLORS.get(level, C_FG)};"
                " font-size: 26px; font-weight: 700;")
            self._set_hero_level_style(level or "—",
                                       LEVEL_COLORS.get(level, C_FG_MUTED),
                                       LEVEL_BGS.get(level, C_CARD))
        else:
            self.hero_cpk.setText("—")
            self.hero_cpk.setStyleSheet(
                "color: #FFFFFF; font-size: 26px; font-weight: 700;")
            self._set_hero_level_style("—")

        for key, chip in self.stat_labels.items():
            if chip is None or key not in stats:
                continue
            val = stats[key]
            color = C_FG

            if key in ["Count", "OutOfSpecCount"]:
                txt = f"{int(val)}"
                if key == "OutOfSpecCount":
                    color = C_DANGER if int(val) > 0 else C_SUCCESS
            elif key == "PPM":
                txt = f"{int(val)}"
            elif key == "SigmaLevel":
                txt = f"{val:.2f}σ"
            elif key == "CV":
                txt = f"{val:.2f}%"
            elif key in ["USL", "LSL"] and val is None:
                txt = "未设置"
                color = C_FG_DIM
            elif key in ["Skewness", "Kurtosis"]:
                txt = fmt_val(val, 3)
                color = C_WARNING if abs(val) > 1.0 else C_SUCCESS
            elif key in ['Cp', 'Cpk', 'CPU', 'CPL']:
                txt = fmt_val(val, 3)
            else:
                txt = fmt_val(val, 4)

            if val is None and key not in ["USL", "LSL"]:
                color = C_FG_DIM

            chip.set_value(txt, color)

        self.set_status(f"Cpk {cpk:.3f} · {level}" if cpk is not None else "已更新", 'ok')

    def on_clear_tab1(self):
        self.inp_an_usl.clear()
        self.inp_an_lsl.clear()
        self.txt_data.clear()
        self.reset_chart()
        self.btn_export.setEnabled(False)

    def on_clear_tab2(self):
        for e in [self.inp_sim_usl, self.inp_sim_lsl, self.inp_sim_cpk,
                  self.inp_sim_mean, self.inp_sim_cnt, self.inp_sim_prec]:
            e.clear()
        self.txt_sim.clear()
        self.reset_chart()
        self.btn_export.setEnabled(False)

    def on_copy(self):
        QGuiApplication.clipboard().setText(self.txt_sim.toPlainText())
        show_msg(self, "复制成功", "内容已复制到剪贴板", False)

    def show_about(self):
        about_text = (
            "CPK 统计分析工具 V7.4 (Fluent UI)\n\n"
            "Copyright © 2025 Github:OUKUI All Rights Reserved."
        )
        show_msg(self, "关于软件", about_text, is_error=False)

    def _show_debug_prompt(self):
        if self._debug_tab_active:
            return
        pwd = FluentInputDialog.get_password(self, "调试模式", "请输入调试密码：")
        if pwd == "114514":
            t2 = QWidget()
            self.main_notebook.insertTab(1, t2, "模拟生成")
            self.setup_tab2(t2)
            self.main_notebook.setCurrentIndex(1)
            self._debug_tab_active = True
            self._debug_btn.setText("调试模式 ✓")
            self._debug_btn.setStyleSheet(
                "color: #00B050; font-size: 11px; padding: 4px;")
            show_msg(self, "调试模式", "模拟生成模块已激活", is_error=False)
            self.set_status("调试模式", 'info')

    def _toggle_debug(self, event=None):
        self._show_debug_prompt()

    # ==========================================
    # Excel 导入相关功能
    # ==========================================
    def load_excel_file(self):
        if not PANDAS_AVAILABLE:
            show_msg(self, "缺少依赖",
                     "未安装 pandas 或 openpyxl。\n请运行：pip install pandas openpyxl")
            return

        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择 Excel 文件", "", "Excel Files (*.xlsx *.xls)")
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
                show_msg(self, "导入失败", msg)
                self.set_status("导入失败", 'err')
                return

            self.excel_projects = all_projects
            self.refresh_excel_treeview()

            if self.excel_projects:
                self.tree_projects.setCurrentItem(self.tree_projects.topLevelItem(0))
                self.on_excel_item_select()

            msg = f"成功导入 {len(self.excel_projects)} 个项目"
            if is_comparison_mode:
                msg += f"（{len(sheet_names)} 个子表，交叉对比模式）"
            msg += "。"
            if error_logs:
                msg += f"\n跳过 {len(error_logs)} 个无效项目。"
            show_msg(self, "导入成功", msg, is_error=False)
            self.set_status(f"已导入 {len(self.excel_projects)} 项", 'ok')

        except Exception as e:
            show_msg(self, "导入失败", f"读取 Excel 文件时出错:\n{str(e)}")
            self.set_status("导入失败", 'err')

    def refresh_excel_treeview(self):
        self.tree_projects.clear()
        for i, proj in enumerate(self.excel_projects):
            cpk = proj['cpk_val']
            level = proj['level']
            sheet_name = proj.get('sheet_name', 'Sheet1')
            item = QTreeWidgetItem([
                str(proj['name']),
                str(sheet_name),
                f"{cpk:.3f}" if cpk is not None else "-",
                str(level),
            ])
            item.setData(0, Qt.ItemDataRole.UserRole, i)
            self.tree_projects.addTopLevelItem(item)

    def on_excel_item_select(self, event=None):
        item = self.tree_projects.currentItem()
        if item is None:
            return
        idx = item.data(0, Qt.ItemDataRole.UserRole)
        if idx is None:
            return
        idx = int(idx)
        if idx < 0 or idx >= len(self.excel_projects):
            return

        self.current_excel_index = idx
        project = self.excel_projects[idx]

        self.update_stats_display(project['stats'], project_name=project['name'])
        self.draw_chart(project['data'], project['stats'])
        self.root.update()

        if "Excel" in self._current_tab_text():
            self.btn_export.setEnabled(True)

    def clear_excel_data(self):
        self.excel_projects = []
        self.current_excel_index = -1
        self.tree_projects.clear()
        self.reset_chart()
        if "Excel" in self._current_tab_text():
            self.btn_export.setEnabled(False)
        self.set_status("列表已清空", 'info')

    def export_merged_report(self):
        if not self.excel_projects:
            show_msg(self, "无数据", "没有可导出的项目数据。请先导入 Excel。")
            return

        if not REPORTLAB_AVAILABLE:
            show_msg(self, "缺少依赖", "未安装 reportlab。")
            return

        default_filename = f"CPK_汇总报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"

        default_path = os.path.join(self.app_dir, default_filename)
        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存汇总报告 (所有项目将合并为此文件)",
            default_path, "PDF 文件 (*.pdf);;所有文件 (*.*)")

        if not file_path:
            return
        if not file_path.lower().endswith('.pdf'):
            file_path += '.pdf'

        try:
            self.btn_batch_export.setEnabled(False)
            self.btn_batch_export.setText("生成中...")
            self.set_status("导出中…", 'info')
            QApplication.processEvents()

            self._generate_merged_pdf_report(file_path, self.excel_projects)

            show_msg(self, "导出成功",
                     f"所有 {len(self.excel_projects)} 个项目已合并保存至:\n{file_path}",
                     is_error=False)
            self.set_status("导出完成", 'ok')
        except Exception as e:
            show_msg(self, "导出失败", f"错误:\n{str(e)}\n{traceback.format_exc()}")
            self.set_status("导出失败", 'err')
        finally:
            self.btn_batch_export.setEnabled(True)
            self.btn_batch_export.setText("合并全部 PDF")

    def export_report(self):
        if not REPORTLAB_AVAILABLE:
            show_msg(self, "缺少依赖", "未安装 reportlab。")
            return

        if "Excel" in self._current_tab_text():
            if not self.excel_projects or self.current_excel_index == -1:
                show_msg(self, "无数据", "请先在列表中选择一个项目。")
                return
            project = self.excel_projects[self.current_excel_index]
            data_to_export = project['data']
            stats_to_export = project['stats']
            usl_to_export = project['usl']
            lsl_to_export = project['lsl']
            name_to_export = project['name']
        elif self.current_stats is None:
            show_msg(self, "无数据", "请先进行分析或模拟生成数据。")
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

        default_path = os.path.join(self.app_dir, default_filename)
        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存 CPK 报告",
            default_path, "PDF 文件 (*.pdf);;所有文件 (*.*)")

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

            show_msg(self, "导出成功", f"报告已保存至:\n{file_path}", is_error=False)
            self.set_status("导出完成", 'ok')
        except Exception as e:
            show_msg(self, "导出失败", f"错误:\n{str(e)}\n{traceback.format_exc()}")
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

            frame_full = Frame(
                page_margin, page_margin, usable_w, usable_h,
                id='full',
                leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
                showBoundary=0,
            )
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

        pad = 5
        header_h = 20
        sec_label_h = 11
        spacer_after_header = 3
        content_budget = card_h - 2 * pad
        body_h = content_budget - header_h - spacer_after_header
        if body_h < 100:
            body_h = 100

        gutter = 6
        left_w = card_w * 0.58 - gutter / 2
        right_w = card_w * 0.42 - gutter / 2
        metrics_detail_h = 76
        chart_h = body_h - sec_label_h - metrics_detail_h - 4

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
    app = FluentApp()
    apply_theme(QApplication.instance())
    win = CpkApp()
    win.show()
    win._set_dark_titlebar()
    sys.exit(app.exec())
