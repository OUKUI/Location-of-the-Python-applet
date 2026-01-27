import tkinter as tk
from tkinter import ttk, messagebox
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from scipy import stats
import ctypes

# 开启高DPI感知
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except:
    pass

# 全局配色
BG_MAIN = "#121212"
BG_PANEL = "#1E1E1E"
FG_TEXT = "#E0E0E0"
ACCENT = "#00A2ED"
LINE_UCL = "#FF4444"
LINE_CL = "#44FF44"

class ProfessionalSPC:
    def __init__(self, root):
        self.root = root
        self.root.title("过程能力控制表 (Advanced SPC Dashboard)")
        self.root.state('zoomed')
        self.root.configure(bg=BG_MAIN)
        
        # SPC 系数 (n=9)
        self.A2, self.D3, self.D4 = 0.337, 0.184, 1.816
        
        self.rows, self.cols = 9, 40
        self.cells = []
        self.setup_ui()

    def setup_ui(self):
        # --- 1. 顶部规格区 ---
        header = tk.Frame(self.root, bg=BG_MAIN, pady=10)
        header.pack(fill=tk.X)
        
        params = [("USL", "4.5"), ("LSL", "1.0"), ("基准值", "2.75"), ("n", "9")]
        self.entries = {}
        for name, val in params:
            f = tk.Frame(header, bg=BG_MAIN)
            f.pack(side=tk.LEFT, padx=20)
            tk.Label(f, text=name, bg=BG_MAIN, fg=ACCENT, font=("微软雅黑", 10)).pack()
            e = tk.Entry(f, width=8, bg=BG_PANEL, fg="white", insertbackground="white", bd=0)
            e.insert(0, val)
            e.pack(pady=5)
            self.entries[name] = e

        # --- 2. 数据矩阵区 ---
        grid_frame = tk.Frame(self.root, bg=BG_MAIN)
        grid_frame.pack(fill=tk.X, padx=10)
        
        canvas = tk.Canvas(grid_frame, bg=BG_MAIN, height=220, highlightthickness=0)
        h_scroll = ttk.Scrollbar(grid_frame, orient="horizontal", command=canvas.xview)
        self.scroll_inner = tk.Frame(canvas, bg=BG_MAIN)
        
        canvas.create_window((0, 0), window=self.scroll_inner, anchor="nw")
        canvas.configure(xscrollcommand=h_scroll.set)
        
        for r in range(self.rows):
            row_data = []
            tk.Label(self.scroll_inner, text=f"#{r+1}", bg=BG_MAIN, fg="#666").grid(row=r, column=0)
            for c in range(self.cols):
                ent = tk.Entry(self.scroll_inner, width=5, bg="#252525", fg=FG_TEXT, bd=1, relief="flat")
                ent.grid(row=r, column=c+1, padx=1, pady=1)
                ent.bind("<Control-v>", self.handle_paste)
                row_data.append(ent)
            self.cells.append(row_data)
            
        canvas.pack(fill=tk.X)
        h_scroll.pack(fill=tk.X)
        self.scroll_inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        # --- 3. 图表展示区 ---
        chart_container = tk.Frame(self.root, bg=BG_MAIN)
        chart_container.pack(fill=tk.BOTH, expand=True, pady=10)

        self.fig = plt.figure(figsize=(16, 6), facecolor=BG_MAIN)
        # 布局：X控制图, R控制图, 直方图
        self.ax_x = self.fig.add_subplot(131)
        self.ax_r = self.fig.add_subplot(132)
        self.ax_h = self.fig.add_subplot(133)
        
        self.canvas_plt = FigureCanvasTkAgg(self.fig, master=chart_container)
        self.canvas_plt.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # --- 4. 底部状态栏 ---
        self.status = tk.Label(self.root, text="就绪 | 请输入数据并点击更新", bg=BG_PANEL, fg=FG_TEXT, font=("Consolas", 11), pady=5)
        self.status.pack(fill=tk.X)
        
        tk.Button(self.root, text=" 重新计算并生成控制图 ", bg=ACCENT, fg="white", font=("微软雅黑", 12, "bold"), 
                  command=self.analyze, bd=0, padx=20, pady=10).pack(pady=10)

    def handle_paste(self, event):
        try:
            raw = self.root.clipboard_get()
            lines = raw.split('\n')
            for i, line in enumerate(lines):
                if i >= self.rows: break
                vals = line.split('\t')
                for j, v in enumerate(vals):
                    if j >= self.cols: break
                    self.cells[i][j].delete(0, tk.END)
                    self.cells[i][j].insert(0, v.strip())
            return "break"
        except: pass

    def analyze(self):
        try:
            # 1. 数据收集 (按列计算均值和极差)
            col_means, col_ranges, all_data = [], [], []
            for c in range(self.cols):
                col_vals = []
                for r in range(self.rows):
                    v = self.cells[r][c].get()
                    if v: col_vals.append(float(v))
                
                if col_vals:
                    col_means.append(np.mean(col_vals))
                    col_ranges.append(np.max(col_vals) - np.min(col_vals))
                    all_data.extend(col_vals)

            if not all_data: return

            all_data = np.array(all_data)
            usl, lsl = float(self.entries["USL"].get()), float(self.entries["LSL"].get())
            
            # 2. 计算 SPC 指标
            grand_mean = np.mean(col_means)
            mean_range = np.mean(col_ranges)
            sigma = np.std(all_data, ddof=1)
            
            cpk = min((usl - grand_mean)/(3*sigma), (grand_mean - lsl)/(3*sigma))
            
            # X-bar 控制限
            ucl_x = grand_mean + self.A2 * mean_range
            lcl_x = grand_mean - self.A2 * mean_range
            
            # R 控制限
            ucl_r = self.D4 * mean_range
            lcl_r = self.D3 * mean_range

            # 3. 绘图逻辑
            self.render_charts(col_means, ucl_x, lcl_x, grand_mean, col_ranges, ucl_r, lcl_r, mean_range, all_data, usl, lsl, sigma)
            
            # 4. 更新状态
            self.status.config(text=f"均值: {grand_mean:.3f} | Sigma: {sigma:.3f} | CPK: {cpk:.3f} | {'需优化' if cpk < 1.33 else '合格'}")
            
        except Exception as e:
            messagebox.showerror("计算错误", str(e))

    def render_charts(self, mx, uclx, lclx, clx, mr, uclr, lclr, clr, data, usl, lsl, sigma):
        for ax in [self.ax_x, self.ax_r, self.ax_h]:
            ax.clear()
            ax.set_facecolor("#151515")
            ax.tick_params(colors="#888", labelsize=8)

        # X-Bar Chart
        self.ax_x.plot(mx, marker='o', color=ACCENT, markersize=4, label="均值线")
        self.ax_x.axhline(uclx, color=LINE_UCL, ls='--', lw=1, label="UCL")
        self.ax_x.axhline(clx, color=LINE_CL, ls='-', lw=1)
        self.ax_x.axhline(lclx, color=LINE_UCL, ls='--', lw=1, label="LCL")
        self.ax_x.set_title("X-bar 控制图", color=FG_TEXT, fontsize=10)

        # R Chart
        self.ax_r.plot(mr, marker='s', color="#FFA500", markersize=4, label="极差线")
        self.ax_r.axhline(uclr, color=LINE_UCL, ls='--', lw=1)
        self.ax_r.axhline(clr, color=LINE_CL, ls='-', lw=1)
        self.ax_r.axhline(lclr, color=LINE_UCL, ls='--', lw=1)
        self.ax_r.set_title("R 极差控制图", color=FG_TEXT, fontsize=10)

        # Histogram + Normal
        n, bins, _ = self.ax_h.hist(data, bins=15, density=True, color=ACCENT, alpha=0.4)
        mu = np.mean(data)
        x = np.linspace(lsl-0.5, usl+0.5, 100)
        self.ax_h.plot(x, stats.norm.pdf(x, mu, sigma), color="#F1C40F", lw=2)
        self.ax_h.axvline(usl, color="red", ls='-')
        self.ax_h.axvline(lsl, color="red", ls='-')
        self.ax_h.set_title("分布分布与直方图", color=FG_TEXT, fontsize=10)

        self.fig.tight_layout()
        self.canvas_plt.draw()

if __name__ == "__main__":
    root = tk.Tk()
    app = ProfessionalSPC(root)
    root.mainloop()