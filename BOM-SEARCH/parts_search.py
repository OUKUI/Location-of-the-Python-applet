#!/usr/bin/env python
# -*- coding: utf-8 -*-

import ctypes
import os
import re
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd
import openpyxl

FONT      = 'Microsoft YaHei UI'
FONT_SZ   = 13
FONT_SM   = 11
FONT_TITLE = 17
FONT_H1   = 14

def _init_dpi():
    if sys.platform != 'win32':
        return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

_init_dpi()

DARK_BG        = '#1e1e1e'
DARK_BG2       = '#252526'
DARK_BG3       = '#2d2d30'
DARK_BG4       = '#3c3c3c'
TEXT_PRIMARY   = '#cccccc'
TEXT_SECONDARY = '#9d9d9d'
ACCENT         = '#0078d4'
ACCENT_HOVER   = '#1e8cd4'
ACCENT_ACTIVE  = '#005a9e'
BORDER         = '#3e3e42'
TREE_SEL       = '#264f78'
MATCH_HL       = '#4a3000'


def _apply_dark_style(root):
    style = ttk.Style(root)
    root.configure(bg=DARK_BG)
    style.theme_use('clam')

    style.configure('.',
        background=DARK_BG, foreground=TEXT_PRIMARY,
        fieldbackground=DARK_BG3, bordercolor=BORDER,
        darkcolor=DARK_BG, lightcolor=DARK_BG,
        troughcolor=DARK_BG2, selectbackground=ACCENT,
        selectforeground='#ffffff', arrowcolor=TEXT_PRIMARY,
        font=(FONT, FONT_SZ),
    )
    style.configure('TFrame', background=DARK_BG)
    style.configure('TLabelframe', background=DARK_BG, bordercolor=BORDER)
    style.configure('TLabelframe.Label', background=DARK_BG,
                    foreground=TEXT_PRIMARY, font=(FONT, FONT_SZ))
    style.configure('TLabel', background=DARK_BG, foreground=TEXT_PRIMARY,
                    font=(FONT, FONT_SZ))
    style.configure('Small.TLabel', font=(FONT, FONT_SM))
    style.configure('Bold.TLabel', font=(FONT, FONT_SZ, 'bold'))
    style.configure('H1.TLabel', font=(FONT, FONT_H1, 'bold'))
    style.configure('Title.TLabel', font=(FONT, FONT_TITLE, 'bold'))
    style.configure('Section.TLabel', font=(FONT, FONT_SZ+1, 'bold'))

    style.configure('TButton',
        background=DARK_BG3, foreground=TEXT_PRIMARY,
        bordercolor=BORDER, relief='flat',
        padding=(14, 8), font=(FONT, FONT_SZ),
    )
    style.map('TButton',
        background=[('active', ACCENT_HOVER), ('pressed', ACCENT_ACTIVE),
                    ('disabled', DARK_BG2)],
        foreground=[('active', '#ffffff'), ('disabled', '#555555')],
    )
    style.configure('Accent.TButton',
        background=ACCENT, foreground='#ffffff',
        bordercolor=ACCENT, padding=(10, 5),
        font=(FONT, FONT_SZ, 'bold'),
    )
    style.map('Accent.TButton',
        background=[('active', ACCENT_HOVER), ('pressed', ACCENT_ACTIVE),
                    ('disabled', DARK_BG4)],
        foreground=[('disabled', '#888888')],
    )

    style.configure('TEntry',
        fieldbackground=DARK_BG3, foreground=TEXT_PRIMARY,
        insertcolor=TEXT_PRIMARY, bordercolor=BORDER,
        padding=6, font=(FONT, FONT_SZ),
    )
    style.map('TEntry',
        fieldbackground=[('readonly', DARK_BG2), ('focus', DARK_BG3)],
        bordercolor=[('focus', ACCENT)],
    )
    style.configure('TCombobox',
        fieldbackground=DARK_BG3, foreground=TEXT_PRIMARY,
        arrowcolor=TEXT_PRIMARY, bordercolor=BORDER,
        padding=4, font=(FONT, FONT_SZ),
    )
    style.map('TCombobox',
        fieldbackground=[('readonly', DARK_BG3), ('focus', DARK_BG3)],
        bordercolor=[('focus', ACCENT)],
    )
    root.option_add('*TCombobox*Listbox.background', DARK_BG3)
    root.option_add('*TCombobox*Listbox.foreground', TEXT_PRIMARY)
    root.option_add('*TCombobox*Listbox.selectBackground', ACCENT)
    root.option_add('*TCombobox*Listbox.selectForeground', '#ffffff')
    root.option_add('*TCombobox*Listbox.font', (FONT, FONT_SZ))

    style.configure('TCheckbutton',
        background=DARK_BG, foreground=TEXT_PRIMARY,
        font=(FONT, FONT_SZ),
    )
    style.map('TCheckbutton', background=[('active', DARK_BG)])

    style.configure('Vertical.TScrollbar',
        background=DARK_BG2, troughcolor=DARK_BG,
        bordercolor=DARK_BG, arrowcolor=TEXT_SECONDARY,
        gripcount=0, width=14,
    )
    style.map('Vertical.TScrollbar', background=[('active', DARK_BG4)])
    style.configure('Horizontal.TScrollbar',
        background=DARK_BG2, troughcolor=DARK_BG,
        bordercolor=DARK_BG, arrowcolor=TEXT_SECONDARY,
        gripcount=0, width=14,
    )

    style.configure('Treeview',
        background=DARK_BG3, foreground=TEXT_PRIMARY,
        fieldbackground=DARK_BG3, borderwidth=0,
        font=(FONT, FONT_SZ), rowheight=40,
    )
    style.map('Treeview',
        background=[('selected', TREE_SEL)],
        foreground=[('selected', '#ffffff')],
    )
    style.configure('Treeview.Heading',
        background=DARK_BG2, foreground=TEXT_PRIMARY,
        relief='flat', font=(FONT, FONT_SZ, 'bold'),
        padding=(8, 5),
    )
    style.map('Treeview.Heading', background=[('active', DARK_BG4)])

    style.configure('TProgressbar',
        background=ACCENT, troughcolor=DARK_BG2,
        borderwidth=0, thickness=6,
    )
    style.configure('TSeparator', background=BORDER)


# =====================================================================
class PartsSearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title('跨表零件号联合检索工具')
        self.root.minsize(1100, 650)
        self.root.state('zoomed')

        self.files = []
        self.file_cache = {}
        self.right_visible = False

        _apply_dark_style(root)
        self._build_ui()
        self.root.after(100, self._auto_import)

    # ==================================================================
    def _build_ui(self):
        main = ttk.Frame(self.root)
        main.pack(fill=tk.BOTH, expand=True)

        self._build_top_bar(main)

        body = ttk.Frame(main)
        body.pack(fill=tk.BOTH, expand=True)
        body.grid_rowconfigure(0, weight=1)
        body.grid_columnconfigure(2, weight=1)

        self.left_frame = ttk.Frame(body, width=250)
        self.left_frame.grid(row=0, column=0, sticky='ns')
        self.left_frame.grid_propagate(False)
        self._build_left_panel(self.left_frame)

        sep1 = ttk.Frame(body, width=1)
        sep1.configure(style='TSeparator')
        sep1.grid(row=0, column=1, sticky='ns')

        self.mid_right_container = ttk.Frame(body)
        self.mid_right_container.grid(row=0, column=2, sticky='nsew')

        self.center_frame = ttk.Frame(self.mid_right_container)
        self.center_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._build_center_panel(self.center_frame)

        self.right_frame = ttk.Frame(self.mid_right_container)
        self._build_right_panel(self.right_frame)

        self._build_status_bar(main)

        self._update_status()

    # -------------------- top bar --------------------
    def _build_top_bar(self, parent):
        f = ttk.Frame(parent, padding=(16, 12, 16, 8))
        f.pack(fill=tk.X)
        ttk.Label(f, text='跨表零件号联合检索工具',
                  style='Title.TLabel').pack(side=tk.LEFT)
        ttk.Label(f, text='BOM 多文件联合查询',
                  foreground=TEXT_SECONDARY, style='Small.TLabel').pack(
                      side=tk.LEFT, padx=(14, 0))
        ttk.Separator(parent, orient='horizontal').pack(fill=tk.X)

    # -------------------- left panel --------------------
    def _build_left_panel(self, parent):
        c = ttk.Frame(parent, padding=(12, 8, 8, 8))
        c.pack(fill=tk.BOTH, expand=True)

        ttk.Label(c, text='基础设置', style='Section.TLabel').pack(
            anchor=tk.W, pady=(0, 6))

        ttk.Label(c, text='文件管理', style='Bold.TLabel').pack(anchor=tk.W, pady=(6, 4))
        btn_f = ttk.Frame(c)
        btn_f.pack(fill=tk.X)
        ttk.Button(btn_f, text='+ 导入文件', command=self._import_files,
                   style='Accent.TButton').pack(side=tk.LEFT)
        ttk.Button(btn_f, text='清空', command=self._clear_files).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_f, text='移除', command=self._remove_selected_files).pack(side=tk.LEFT)

        tree_f = ttk.Frame(c)
        tree_f.pack(fill=tk.BOTH, expand=True, pady=(4, 0))
        self.file_tree = ttk.Treeview(tree_f, columns=('name',), show='headings',
                                       height=5, selectmode='extended')
        self.file_tree.heading('name', text='文件名')
        self.file_tree.column('name', width=240)
        self.file_tree.tag_configure('alt_row', background=DARK_BG2)
        fsb = ttk.Scrollbar(tree_f, orient=tk.VERTICAL, command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=fsb.set)
        self.file_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        fsb.pack(side=tk.RIGHT, fill=tk.Y)

        ttk.Separator(c, orient='horizontal').pack(fill=tk.X, pady=(10, 6))

        ttk.Label(c, text='搜索设置', style='Bold.TLabel').pack(anchor=tk.W)

        ttk.Label(c, text='搜索列', padding=(0, 8, 0, 2)).pack(anchor=tk.W)
        col_row = ttk.Frame(c)
        col_row.pack(fill=tk.X)
        self.col_var = tk.StringVar(value='G')
        cb = ttk.Combobox(col_row, textvariable=self.col_var,
                          values=[chr(i) for i in range(65, 91)],
                          width=5, state='readonly')
        cb.pack(side=tk.LEFT)
        ttk.Label(col_row, text='A-第1列  G-第7列',
                  foreground=TEXT_SECONDARY, style='Small.TLabel').pack(
                      side=tk.LEFT, padx=(6, 0))

        ttk.Label(c, text='零件号', padding=(0, 8, 0, 2)).pack(anchor=tk.W)
        self.part_entry = ttk.Entry(c)
        self.part_entry.pack(fill=tk.X)
        self.part_entry.bind('<Return>', lambda e: self._do_search())
        ttk.Label(c, text='逗号/分号/换行分隔多个零件号',
                  foreground=TEXT_SECONDARY, style='Small.TLabel').pack(anchor=tk.W, pady=(0, 8))

        self.ignore_case = tk.BooleanVar(value=True)
        ttk.Checkbutton(c, text='大小写不敏感',
                        variable=self.ignore_case).pack(anchor=tk.W)

        ttk.Button(c, text='搜索', command=self._do_search,
                   style='Accent.TButton').pack(fill=tk.X, pady=(10, 4))

        ttk.Separator(c, orient='horizontal').pack(fill=tk.X, pady=(8, 6))

        ttk.Label(c, text='操作', style='Bold.TLabel').pack(anchor=tk.W)
        self.export_btn = ttk.Button(c, text='导出结果 (Excel)',
                                      command=self._export_results, state=tk.DISABLED)
        self.export_btn.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(c, text='清空结果',
                    command=self._clear_results).pack(fill=tk.X, pady=(4, 0))

    # -------------------- center panel --------------------
    def _build_center_panel(self, parent):
        c = ttk.Frame(parent, padding=(0, 0, 0, 0))
        c.pack(fill=tk.BOTH, expand=True)

        hdr = ttk.Frame(c, padding=(0, 4, 0, 4))
        hdr.pack(fill=tk.X)
        ttk.Label(hdr, text='查询结果', style='Section.TLabel').pack(side=tk.LEFT)
        self.result_count_label = ttk.Label(hdr, text='', foreground=TEXT_SECONDARY,
                                            style='Small.TLabel')
        self.result_count_label.pack(side=tk.LEFT, padx=(10, 0))
        ttk.Label(hdr, text='双击查看完整表单',
                  foreground=TEXT_SECONDARY, style='Small.TLabel').pack(side=tk.RIGHT)

        tree_f = ttk.Frame(c)
        tree_f.pack(fill=tk.BOTH, expand=True)
        cols = ('零件号', '文件名', '工作表',
                '出现次数', '所在行', '匹配索引')
        self.result_tree = ttk.Treeview(tree_f, columns=cols, show='headings')
        col_w = [160, 220, 180, 90, 0, 0]
        for c, w in zip(cols, col_w):
            self.result_tree.heading(c, text=c)
            if w:
                self.result_tree.column(c, width=w, minwidth=60)
            else:
                self.result_tree.column(c, width=0, stretch=False)
        self.result_tree.tag_configure('alt_row', background=DARK_BG2)
        self.result_tree.bind('<Double-1>', self._on_result_double_click)
        self.result_tree.bind('<Delete>', lambda e: self._remove_selected_results())
        rsb = ttk.Scrollbar(tree_f, orient=tk.VERTICAL, command=self.result_tree.yview)
        self.result_tree.configure(yscrollcommand=rsb.set)
        self.result_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        rsb.pack(side=tk.RIGHT, fill=tk.Y)

    # -------------------- right panel --------------------
    def _build_right_panel(self, parent):
        c = ttk.Frame(parent, padding=(8, 4, 4, 4))
        c.pack(fill=tk.BOTH, expand=True)

        hdr = ttk.Frame(c)
        hdr.pack(fill=tk.X, pady=(0, 4))
        self.sheet_title = ttk.Label(hdr, text='工作表详情',
                                      style='Section.TLabel')
        self.sheet_title.pack(side=tk.LEFT)
        self.sheet_info = ttk.Label(hdr, text='', foreground=TEXT_SECONDARY,
                                     style='Small.TLabel')
        self.sheet_info.pack(side=tk.LEFT, padx=(10, 0))
        ttk.Button(hdr, text='✕', width=3,
                    command=self._hide_sheet_panel).pack(side=tk.RIGHT)

        self._sheet_hint_var = tk.StringVar(value='')

        ttk.Label(c, text='',
                  foreground='#e5c07b',
                  style='Small.TLabel',
                  textvariable=self._sheet_hint_var).pack(anchor=tk.W, pady=(0, 4))

        tree_f = ttk.Frame(c)
        tree_f.pack(fill=tk.BOTH, expand=True)
        tree_f.grid_rowconfigure(0, weight=1)
        tree_f.grid_columnconfigure(0, weight=1)

        self.sheet_tree = ttk.Treeview(tree_f, show='headings')
        self.sheet_tree.tag_configure('sh_header', background=DARK_BG2,
                                       foreground=TEXT_PRIMARY,
                                       font=(FONT, FONT_SZ, 'bold'))
        self.sheet_tree.tag_configure('sh_matched', background='#5a3a00',
                                       foreground='#ffce7b')
        self.sheet_tree.tag_configure('sh_alt', background=DARK_BG2)
        self.sheet_tree.tag_configure('sh_normal', background=DARK_BG3)

        vsb = ttk.Scrollbar(tree_f, orient=tk.VERTICAL, command=self.sheet_tree.yview)
        hsb = ttk.Scrollbar(tree_f, orient=tk.HORIZONTAL, command=self.sheet_tree.xview)
        self.sheet_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.sheet_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

    # -------------------- status bar --------------------
    def _build_status_bar(self, parent):
        btm = ttk.Frame(parent, padding=(16, 6, 16, 8))
        btm.pack(fill=tk.X, side=tk.BOTTOM)
        ttk.Separator(btm, orient='horizontal').pack(fill=tk.X, pady=(0, 4))
        self.status_var = tk.StringVar(value='尚未导入文件')
        ttk.Label(btm, textvariable=self.status_var,
                  foreground=TEXT_SECONDARY, style='Small.TLabel').pack(side=tk.LEFT)

    # ==================================================================
    #  file operations
    # ==================================================================
    def _auto_import(self):
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
        except NameError:
            script_dir = os.getcwd()
        exts = ('.xlsx', '.xlsm', '.xls')
        for entry in os.listdir(script_dir):
            fp = os.path.join(script_dir, entry)
            if os.path.isfile(fp) and entry.lower().endswith(exts):
                if fp not in self.files:
                    self.files.append(fp)
                    self.file_tree.insert('', tk.END, values=(entry,))
        if self.files:
            self._stripe(self.file_tree)
            self._update_status()

    def _import_files(self):
        paths = filedialog.askopenfilenames(
            title='选择 Excel 文件',
            filetypes=[('Excel 文件', '*.xlsx *.xlsm *.xls'), ('All Files', '*.*')])
        if not paths:
            return
        added = 0
        for p in paths:
            if p not in self.files:
                self.files.append(p)
                self.file_tree.insert('', tk.END, values=(os.path.basename(p),))
                added += 1
        if added:
            self._stripe(self.file_tree)
        else:
            messagebox.showinfo('提示', '所选文件均已导入')
        self._update_status()

    def _remove_selected_files(self):
        selected = self.file_tree.selection()
        if not selected:
            return
        for item in reversed(selected):
            idx = self.file_tree.index(item)
            if idx < len(self.files):
                del self.files[idx]
            self.file_tree.delete(item)
        self.file_cache.clear()
        self._stripe(self.file_tree)
        self._update_status()

    def _clear_files(self):
        self.files.clear()
        self.file_cache.clear()
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        self._update_status()

    def _clear_results(self):
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        self.result_count_label.configure(text='')
        self.export_btn.configure(state=tk.DISABLED)
        self._hide_sheet_panel()

    def _remove_selected_results(self):
        for item in self.result_tree.selection():
            self.result_tree.delete(item)
        self._stripe(self.result_tree)

    def _stripe(self, tree):
        for i, item in enumerate(tree.get_children()):
            tree.item(item, tags=('alt_row',) if i % 2 == 0 else ())

    def _update_status(self):
        if not self.files:
            self.status_var.set('尚未导入文件 - 点击“导入文件”开始')
        else:
            self.status_var.set('已导入 {} 个文件'.format(len(self.files)))

    # ==================================================================
    #  core
    # ==================================================================
    def _col_idx(self):
        return ord(self.col_var.get().upper()) - ord('A')

    @staticmethod
    def _norm(v):
        return '' if v is None else str(v).strip()

    def _match(self, cell_val, query):
        c = self._norm(cell_val)
        q = query
        if self.ignore_case.get():
            c = c.lower()
            q = q.lower()
        return c == q

    def _load_file(self, path):
        if path in self.file_cache:
            return self.file_cache[path]
        try:
            ext = os.path.splitext(path)[1].lower()
            if ext in ('.xlsx', '.xlsm'):
                sheets = pd.read_excel(path, sheet_name=None, dtype=str, engine='openpyxl')
            else:
                sheets = pd.read_excel(path, sheet_name=None, dtype=str, engine='xlrd')
        except Exception:
            try:
                sheets = {}
                wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
                for sn in wb.sheetnames:
                    ws = wb[sn]
                    data = []
                    for row in ws.iter_rows(values_only=True):
                        data.append([str(c) if c is not None else '' for c in row])
                    if data:
                        header = data[0]
                        body = data[1:]
                        ncols = max((len(r) for r in data), default=0)
                        for i, r in enumerate(body):
                            if len(r) < ncols:
                                body[i] = list(r) + [''] * (ncols - len(r))
                        if len(header) < ncols:
                            header = list(header) + ['Unnamed_{}'.format(j)
                                                     for j in range(len(header), ncols)]
                        sheets[sn] = pd.DataFrame(body, columns=header)
                    else:
                        sheets[sn] = pd.DataFrame()
                wb.close()
            except Exception:
                messagebox.showerror('错误',
                                     '无法读取文件: ' + os.path.basename(path))
                return None
        self.file_cache[path] = sheets
        return sheets

    # ==================================================================
    #  search
    # ==================================================================
    def _do_search(self):
        raw = self.part_entry.get().strip()
        if not raw:
            messagebox.showwarning('提示', '请输入要查询的零件号')
            return
        if not self.files:
            messagebox.showwarning('提示', '请先导入 Excel 文件')
            return

        queries = [q.strip() for q in re.split(r'[,;\n]+', raw) if q.strip()]
        self._clear_results()
        self._hide_sheet_panel()
        col_idx = self._col_idx()
        col_letter = chr(ord('A') + col_idx)

        self.status_var.set('正在搜索 {} 个零件号 (列 {}) ...'.format(
            len(queries), col_letter))
        self.root.update_idletasks()

        def worker():
            results = []
            for fpath in self.files:
                fname = os.path.basename(fpath)
                sheets = self._load_file(fpath)
                if sheets is None:
                    continue
                for sname, df in sheets.items():
                    if df.empty or col_idx >= df.shape[1]:
                        continue
                    col = df.iloc[:, col_idx]
                    for query in queries:
                        mask = col.apply(lambda x: self._match(x, query))
                        matches = df[mask]
                        cnt = len(matches)
                        if cnt:
                            excel_rows = ','.join(str(i + 2) for i in matches.index)
                            df_indices = ','.join(str(i) for i in matches.index)
                            results.append((query, fname, sname, str(cnt),
                                            excel_rows, df_indices))
            self.root.after(0, lambda: self._show_results(results, queries, col_letter))
        threading.Thread(target=worker, daemon=True).start()

    def _show_results(self, results, queries, col_letter):
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        self._hide_sheet_panel()

        if not results:
            self.status_var.set(
                '未找到匹配结果 - 搜索列: {}'.format(col_letter))
            self.result_count_label.configure(text='0 条记录')
            self.export_btn.configure(state=tk.DISABLED)
            return

        for r in results:
            self.result_tree.insert('', tk.END, values=r)
        self._stripe(self.result_tree)

        found = set(r[0] for r in results)
        missing = [q for q in queries if q not in found]
        total = len(results)
        self.result_count_label.configure(text='共 {} 条记录'.format(total))
        status = '共找到 {} 条记录 (搜索列: {})'.format(total, col_letter)
        if missing:
            status += ' | 未找到: {}'.format(', '.join(missing))
        self.status_var.set(status)
        self.export_btn.configure(state=tk.NORMAL)

    # ==================================================================
    #  right panel - full sheet preview
    # ==================================================================
    def _on_result_double_click(self, event):
        sel = self.result_tree.selection()
        if not sel:
            return
        vals = self.result_tree.item(sel[0])['values']
        if len(vals) < 6:
            return

        part_num = vals[0]
        fname    = vals[1]
        sheet    = vals[2]
        count    = vals[3]
        matched_indices_str = vals[5]

        matched_file = next((f for f in self.files
                             if os.path.basename(f) == fname), None)
        if not matched_file:
            return

        try:
            sheets = self._load_file(matched_file)
            if sheets is None or sheet not in sheets:
                return
            df = sheets[sheet]
            matched_indices = set(
                int(x.strip()) for x in matched_indices_str.split(',')
                if x.strip().isdigit()
            )
            self._show_sheet_panel(part_num, fname, sheet, df, matched_indices)
        except Exception as e:
            messagebox.showerror('错误', '无法加载工作表: ' + str(e))

    def _show_sheet_panel(self, part_num, fname, sheet, df, matched_indices):
        ncols = df.shape[1]
        scols = ['#'] + [chr(ord('A') + i) for i in range(ncols)]

        self.sheet_tree.configure(columns=scols)
        self.sheet_tree.heading('#', text='#')
        self.sheet_tree.column('#', width=45, minwidth=40, stretch=False)
        for c in scols[1:]:
            self.sheet_tree.heading(c, text=c)
            ci = ord(c) - ord('A')
            header_text = str(df.columns[ci]) if ci < len(df.columns) else ''
            w = max(70, min(220, len(header_text) * 14 + 30))
            self.sheet_tree.column(c, width=w, minwidth=50)

        for item in self.sheet_tree.get_children():
            self.sheet_tree.delete(item)

        header_vals = [''] + [str(df.columns[k]) if k < ncols else ''
                              for k in range(ncols)]
        self.sheet_tree.insert('', tk.END, values=header_vals, tags=('sh_header',))

        total_rows = len(df)
        for r_idx in range(total_rows):
            row_vals = [str(r_idx + 1)]
            row_vals += [str(df.iloc[r_idx, k]) if k < ncols else ''
                         for k in range(ncols)]
            if r_idx in matched_indices:
                tag = 'sh_matched'
            elif r_idx % 2 == 0:
                tag = 'sh_normal'
            else:
                tag = 'sh_alt'
            self.sheet_tree.insert('', tk.END, values=row_vals, tags=(tag,))

        self.sheet_title.configure(
            text='工作表: {}'.format(sheet))
        self.sheet_info.configure(
            text='{} | 匹配 {} 行 | 共 {} 行'.format(
                fname, len(matched_indices), total_rows))
        self._sheet_hint_var.set(
            '■ 棕色高亮 = 匹配行    搜索零件号: {}'.format(part_num))

        if not self.right_visible:
            self.right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
            self.right_visible = True

        self._sheet_data = (part_num, fname, sheet, df)

    def _hide_sheet_panel(self):
        if self.right_visible:
            self.right_frame.pack_forget()
            self.right_visible = False
        for item in self.sheet_tree.get_children():
            self.sheet_tree.delete(item)
        self.sheet_title.configure(text='工作表详情')
        self.sheet_info.configure(text='')
        self._sheet_hint_var.set('')

    # ==================================================================
    #  export
    # ==================================================================
    def _export_results(self):
        if not self.result_tree.get_children():
            return
        path = filedialog.asksaveasfilename(
            title='导出结果', defaultextension='.xlsx',
            filetypes=[('Excel 文件', '*.xlsx')])
        if not path:
            return
        try:
            rows = []
            for iid in self.result_tree.get_children():
                rows.append(self.result_tree.item(iid)['values'][:4])
            df = pd.DataFrame(rows, columns=['零件号', '文件名',
                                             '工作表', '出现次数'])
            df.to_excel(path, index=False)
            messagebox.showinfo('提示',
                                '结果已导出到:\n' + os.path.basename(path))
        except Exception as e:
            messagebox.showerror('错误', '导出失败: ' + str(e))


def main():
    root = tk.Tk()
    try:
        root.iconbitmap(default='')
    except Exception:
        pass
    app = PartsSearchApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
