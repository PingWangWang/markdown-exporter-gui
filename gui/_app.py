"""
Markdown Exporter GUI - 主应用类

包含 MarkdownExporterGUI 类，负责：
  - 窗口初始化与图标设置
  - 界面样式（颜色常量 + ttk 主题）
  - 界面构建（输入/输出区域、格式选择、日志、底部链接）
  - 文件选择、目录操作
  - 文件处理（多线程转换）
  - 对话框委托（关于、覆盖确认）
"""
import sys
import os
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path

from _version import APP_VERSION
from _dialogs import show_about, ask_overwrite


# 支持的输出格式
OUTPUT_FORMATS = {
    'DOCX': ('Word 文档', '.docx'),
    'PDF': ('PDF 文档', '.pdf'),
    'HTML': ('HTML 网页', '.html'),
    'PPTX': ('PowerPoint 演示文稿', '.pptx'),
    'XLSX': ('Excel 表格', '.xlsx'),
    'CSV': ('CSV 数据', '.csv'),
    'JSON': ('JSON 数据', '.json'),
    'XML': ('XML 数据', '.xml'),
    'LaTeX': ('LaTeX 文档', '.tex'),
    'IPYNB': ('Jupyter Notebook', '.ipynb'),
    'MD': ('Markdown 文件', '.md'),
}


class MarkdownExporterGUI:

    # ── 初始化 ────────────────────────────────────────────────────────────────

    def __init__(self, root, has_dnd=False):
        self.root = root
        self.has_dnd = has_dnd
        self.root.title(f"Markdown Exporter v{APP_VERSION}")

        window_width, window_height = 750, 560
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        self.root.geometry(f"{window_width}x{window_height}+{(sw-window_width)//2}+{(sh-window_height)//2}")
        self.root.resizable(False, False)

        self._set_window_icon()

        self.input_files     = []
        self.output_dir      = tk.StringVar()
        self.output_format   = tk.StringVar(value='DOCX')  # 默认输出格式
        self.is_processing   = False
        self.last_output_file = None  # 最后一次转换的输出路径（用于"打开文件夹并选中"）

        self.setup_styles()
        self.create_widgets()

        # 窗口完全显示后再次应用图标，确保任务栏图标生效
        self.root.after(100, self._set_window_icon)

    # ── 图标 ──────────────────────────────────────────────────────────────────

    def _get_icon_path(self):
        """返回 icad.ico 的路径，打包/开发环境均适用；找不到则返回 None"""
        meipass = getattr(sys, '_MEIPASS', None)
        # 尝试多个可能的图标位置
        possible_paths = [
            Path(meipass) / 'res' / 'icad.ico' if meipass else None,
            Path(__file__).parent.parent / 'res' / 'icad.ico',
        ]
        for p in possible_paths:
            if p and p.exists():
                return p
        return None

    def _set_window_icon(self):
        """设置窗口图标（标题栏 & 任务栏）"""
        try:
            if sys.platform == 'win32':
                import ctypes
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('MarkdownExporter.GUI.App')

            icon_path = self._get_icon_path()
            if icon_path:
                # 直接使用 .ico 文件
                self.root.iconbitmap(default=str(icon_path))
                try:
                    from PIL import Image, ImageTk
                    img = Image.open(str(icon_path))
                    # 调整为合适的大小
                    img32 = img.resize((32, 32), Image.LANCZOS)
                    self._taskbar_photo = ImageTk.PhotoImage(img32)
                    self.root.wm_iconphoto(True, self._taskbar_photo)
                except Exception:
                    pass
        except Exception:
            pass

    # ── 样式 ──────────────────────────────────────────────────────────────────

    def setup_styles(self):
        """定义配色常量并配置 ttk 主题"""
        self.C_BG         = '#F5F7FA'
        self.C_HEADER_BG  = '#4A90D9'
        self.C_HEADER_FG  = '#FFFFFF'
        self.C_PANEL_BG   = '#FFFFFF'
        self.C_LABEL_FG   = '#374151'
        self.C_ENTRY_BG   = '#EEF2FF'
        self.C_BTN_SEL    = '#4A90D9'
        self.C_BTN_SEL_A  = '#357ABD'
        self.C_BTN_RUN    = '#27AE60'
        self.C_BTN_RUN_A  = '#1E8449'
        self.C_BTN_OPEN   = '#E67E22'
        self.C_BTN_OPEN_A = '#CA6F1E'
        self.C_LOG_BG     = '#1E2533'
        self.C_LOG_FG     = '#D4E6F1'
        self.C_LINK       = '#2E86C1'
        self.C_BORDER     = '#D1D9E6'

        self.root.configure(bg=self.C_BG)
        s = ttk.Style()
        s.theme_use('clam')

        s.configure('.', background=self.C_BG, foreground=self.C_LABEL_FG,
                    font=('Microsoft YaHei UI', 9))
        s.configure('TFrame',       background=self.C_BG)
        s.configure('Panel.TFrame', background=self.C_PANEL_BG, relief='flat', borderwidth=1)
        s.configure('TLabel',       background=self.C_BG, foreground=self.C_LABEL_FG)
        s.configure('Field.TLabel', background=self.C_BG, foreground=self.C_LABEL_FG,
                    font=('Microsoft YaHei UI', 9))
        s.configure('Log.TLabel',   background=self.C_BG, foreground='#6B7280',
                    font=('Microsoft YaHei UI', 9))
        s.configure('Link.TLabel',  background=self.C_BG, foreground=self.C_LINK,
                    cursor='hand2', font=('Microsoft YaHei UI', 9, 'underline'))
        s.configure('TEntry', fieldbackground=self.C_ENTRY_BG, foreground='#1F2937',
                    bordercolor=self.C_BORDER, insertcolor=self.C_LABEL_FG)
        s.configure('TCombobox', fieldbackground=self.C_ENTRY_BG, foreground='#1F2937',
                    bordercolor=self.C_BORDER)

        s.configure('Select.TButton', background=self.C_BTN_SEL, foreground='#FFFFFF',
                    font=('Microsoft YaHei UI', 9, 'bold'),
                    borderwidth=0, focusthickness=0, padding=(8, 4))
        s.map('Select.TButton',
              background=[('active', self.C_BTN_SEL_A), ('disabled', '#A0AEC0')],
              foreground=[('disabled', '#E2E8F0')])

        s.configure('Run.TButton', background=self.C_BTN_RUN, foreground='#FFFFFF',
                    font=('Microsoft YaHei UI', 10, 'bold'),
                    borderwidth=0, focusthickness=0, padding=(12, 6))
        s.map('Run.TButton',
              background=[('active', self.C_BTN_RUN_A), ('disabled', '#A0AEC0')],
              foreground=[('disabled', '#E2E8F0')])

        s.configure('Open.TButton', background=self.C_BTN_OPEN, foreground='#FFFFFF',
                    font=('Microsoft YaHei UI', 10, 'bold'),
                    borderwidth=0, focusthickness=0, padding=(12, 6))
        s.map('Open.TButton',
              background=[('active', self.C_BTN_OPEN_A)],
              foreground=[])

    # ── 界面构建 ──────────────────────────────────────────────────────────────

    def create_widgets(self):
        """构建主界面所有控件"""
        mf = ttk.Frame(self.root, padding="14 10 14 6")
        mf.pack(fill=tk.BOTH, expand=True)
        mf.columnconfigure(1, weight=1)
        row = 0

        # 选择待处理文件
        ttk.Label(mf, text="选择 Markdown 文件:", style='Field.TLabel').grid(
            row=row, column=0, sticky=tk.NW, pady=4, padx=(0, 8))
        ff = ttk.Frame(mf);  ff.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=4)
        ff.columnconfigure(0, weight=1)
        # 文件列表框
        list_frame = tk.Frame(ff, bg=self.C_ENTRY_BG, highlightbackground=self.C_BORDER,
                              highlightthickness=1)
        list_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 6))
        list_frame.columnconfigure(0, weight=1)
        self.file_listbox = tk.Listbox(
            list_frame, height=4, selectmode=tk.EXTENDED,
            bg=self.C_ENTRY_BG, fg='#1F2937',
            selectbackground='#4A90D9', selectforeground='#FFFFFF',
            font=('Microsoft YaHei UI', 9), relief='flat', borderwidth=0,
            activestyle='none')
        self.file_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=4, pady=2)
        list_sb = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        list_sb.grid(row=0, column=1, sticky=(tk.N, tk.S), pady=2)
        self.file_listbox.configure(yscrollcommand=list_sb.set)
        # 按钮列
        btn_col = ttk.Frame(ff);  btn_col.grid(row=0, column=1, sticky=tk.N)
        ttk.Button(btn_col, text="添加文件", command=self.select_files,
                   style='Select.TButton', width=10).pack(pady=(0, 4))
        ttk.Button(btn_col, text="清空列表", command=self.clear_files,
                   style='Select.TButton', width=10).pack()
        # Delete 键删除选中项
        self.file_listbox.bind('<Delete>', lambda e: self.remove_selected_files())
        row += 1

        # 选择保存位置
        ttk.Label(mf, text="选择保存位置:", style='Field.TLabel').grid(
            row=row, column=0, sticky=tk.W, pady=4, padx=(0, 8))
        sf = ttk.Frame(mf);  sf.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=4)
        sf.columnconfigure(0, weight=1)
        ttk.Entry(sf, textvariable=self.output_dir, state='readonly').grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 6))
        ttk.Button(sf, text="保存位置", command=self.select_output_dir,
                   style='Select.TButton', width=10).grid(row=0, column=1)
        row += 1

        # 选择输出格式
        ttk.Label(mf, text="选择输出格式:", style='Field.TLabel').grid(
            row=row, column=0, sticky=tk.W, pady=4, padx=(0, 8))
        cf = ttk.Frame(mf);  cf.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=4)
        # 使用描述文本作为下拉框选项："Word 文档 (.docx)"
        format_list = [f"{desc} ({ext})" for desc, ext in OUTPUT_FORMATS.values()]
        self.format_combo = ttk.Combobox(cf, values=format_list, state='readonly', width=30)
        self.format_combo.set("Word 文档 (.docx)")  # 默认选中
        self.format_combo.grid(row=0, column=0, sticky=tk.W, padx=(0, 6))
        self.format_combo.bind('<<ComboboxSelected>>', self.on_format_change)
        row += 1

        # 分割线
        ttk.Separator(mf, orient='horizontal').grid(
            row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=6)
        row += 1

        # 操作按钮
        bf = ttk.Frame(mf);  bf.grid(row=row, column=0, columnspan=2, pady=4)
        self.process_button = ttk.Button(bf, text="▶  开始转换",
                                         command=self.start_processing,
                                         style='Run.TButton', width=14)
        self.process_button.pack(side=tk.LEFT, padx=6)
        ttk.Button(bf, text="📂  打开输出目录",
                   command=self.open_output_dir,
                   style='Open.TButton', width=14).pack(side=tk.LEFT, padx=6)
        row += 1

        # 日志区域
        ttk.Label(mf, text="处理日志:", style='Log.TLabel').grid(
            row=row, column=0, sticky=tk.NW, pady=(8, 2), padx=(0, 8))
        self.log_text = scrolledtext.ScrolledText(
            mf, height=7, wrap=tk.WORD, font=('Consolas', 9),
            bg=self.C_LOG_BG, fg=self.C_LOG_FG,
            insertbackground=self.C_LOG_FG,
            selectbackground='#2E86C1', selectforeground='#FFFFFF',
            relief='flat', borderwidth=0, state='disabled')
        self.log_text.grid(row=row, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(8, 2))
        mf.rowconfigure(row, weight=1)
        for tag, color in [('success', '#2ECC71'), ('error',    '#E74C3C'),
                           ('info',    '#5DADE2'), ('arrow',    '#F0B429'),
                           ('complete','#A9CCE3'), ('normal',   self.C_LOG_FG)]:
            self.log_text.tag_configure(tag, foreground=color)
        row += 1

        # 底部链接
        lf = ttk.Frame(mf)
        lf.grid(row=row, column=0, columnspan=2, pady=(4, 2), sticky=(tk.W, tk.E))
        lbl = ttk.Label(lf, text="查看项目说明及帮助文档 >>", style='Link.TLabel')
        lbl.pack(side=tk.LEFT)
        lbl.bind('<Button-1>', lambda e: self.show_about())
        ttk.Label(lf, text=f"v{APP_VERSION}", style='Log.TLabel').pack(side=tk.RIGHT)

        # 拖拽支持
        if self.has_dnd:
            self._register_drop_target()

    def _register_drop_target(self):
        """注册整个窗口为拖拽目标，接受 .md / .markdown 文件"""
        from tkinterdnd2 import DND_FILES
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self._on_drop)

    def _on_drop(self, event):
        """处理拖入的文件列表"""
        # tkinterdnd2 返回的路径格式：{path1} {path2} 或 path（含空格时用花括号包裹）
        raw = event.data
        import re
        # 解析路径：花括号包裹的整体路径 或 空格分隔的路径
        paths = re.findall(r'\{([^}]+)\}|([^\s]+)', raw)
        files = [p[0] or p[1] for p in paths]
        md_files = [f for f in files if f.lower().endswith(('.md', '.markdown'))]
        if not md_files:
            self.log_message('✗ 拖入的文件不含 .md / .markdown 文件，已忽略')
            return
        self._add_files(md_files)
        self.log_message(f'已拖入 {len(md_files)} 个文件')

    def on_format_change(self, event=None):
        """当输出格式改变时的回调"""
        pass  # 目前不需要特殊处理

    # ── 日志 ──────────────────────────────────────────────────────────────────

    def log_message(self, message):
        self.log_text.configure(state='normal')
        s = message.strip()
        if   s.startswith(('✓', '✅')):                      tag = 'success'
        elif s.startswith(('✗', '❌')):                      tag = 'error'
        elif s.startswith('[') and ']' in s:                  tag = 'info'
        elif s.startswith(('→', '  →')):                     tag = 'arrow'
        elif s.startswith(('处理完成', '开始处理')):          tag = 'complete'
        else:                                                  tag = 'normal'
        self.log_text.insert(tk.END, message + '\n', tag)
        self.log_text.see(tk.END)
        self.log_text.configure(state='disabled')

    # ── 文件选择 & 目录操作 ───────────────────────────────────────────────────

    def select_files(self):
        filetypes = [
            ('Markdown 文件', '*.md *.markdown'),
            ('所有文件', '*.*'),
        ]
        files = filedialog.askopenfilenames(title="选择 Markdown 文件", filetypes=filetypes)
        if not files:
            return
        self._add_files(list(files))

    def _add_files(self, files):
        """  将文件添加到列表（自动去重）"""
        existing = set(self.input_files)
        new_files = [f for f in files if f not in existing]
        for f in new_files:
            self.input_files.append(f)
            self.file_listbox.insert(tk.END, Path(f).name)
        if not self.output_dir.get() and self.input_files:
            self.output_dir.set(str(Path(self.input_files[0]).parent))

    def clear_files(self):
        self.input_files = []
        self.file_listbox.delete(0, tk.END)

    def remove_selected_files(self):
        selected = list(self.file_listbox.curselection())
        for i in reversed(selected):
            self.file_listbox.delete(i)
            del self.input_files[i]

    def select_output_dir(self):
        d = filedialog.askdirectory(title="选择保存位置")
        if d:
            self.output_dir.set(d)

    def open_output_dir(self):
        out = self.output_dir.get()
        if not out:
            messagebox.showwarning("警告", "请先选择保存位置！");  return
        if not os.path.exists(out):
            messagebox.showerror("错误", f"目录不存在：{out}");  return
        try:
            if sys.platform == 'win32':
                if self.last_output_file and os.path.exists(self.last_output_file):
                    subprocess.run(['explorer', '/select,', self.last_output_file])
                else:
                    os.startfile(out)
            elif sys.platform == 'darwin':
                if self.last_output_file and os.path.exists(self.last_output_file):
                    subprocess.run(['open', '-R', self.last_output_file])
                else:
                    subprocess.run(['open', out])
            else:
                os.system(f'xdg-open "{out}"')
        except Exception as e:
            messagebox.showerror("错误", f"无法打开目录：{e}")

    # ── 获取输出格式 ─────────────────────────────────────────────────────────

    def get_selected_format(self):
        """获取用户选择的输出格式代码"""
        selected = self.format_combo.get()
        for code, (desc, ext) in OUTPUT_FORMATS.items():
            if f"{desc} ({ext})" == selected:
                return code
        return 'DOCX'  # 默认

    # ── 文件处理 ──────────────────────────────────────────────────────────────

    def start_processing(self):
        if not self.input_files:
            messagebox.showwarning("警告", "请先选择要处理的文件！");  return
        if not self.output_dir.get():
            messagebox.showwarning("警告", "请选择保存位置！");  return
        
        output_format = self.get_selected_format()
        self.log_message(f"输出格式: {OUTPUT_FORMATS[output_format][0]}")
        
        self.process_button.configure(state='disabled')
        self.is_processing = True
        t = threading.Thread(target=self.process_files, daemon=True)
        t.start()

    def process_files(self):
        """后台线程：批量转换文件"""
        self._overwrite_all = False
        self._skip_all      = False
        try:
            total = len(self.input_files)
            output_format = self.get_selected_format()
            format_desc, format_ext = OUTPUT_FORMATS[output_format]
            
            self.log_message(f"开始处理 {total} 个文件...")
            self.log_message(f"目标格式: {format_desc}")
            
            for i, file_path in enumerate(self.input_files, 1):
                if not self.is_processing:
                    self.log_message("处理已取消");  break
                self.log_message(f"[{i}/{total}] 正在转换: {Path(file_path).name}")
                self.convert_file(file_path, output_format)
                stem = Path(file_path).stem
                self.log_message(f"✓ 转换成功: {stem}{format_ext}")
            self.log_message(f"\n处理完成！共处理 {total} 个文件。")
        except Exception as e:
            self.log_message(f"\n✗ 处理失败: {e}")
        finally:
            self.root.after(0, self.processing_complete)

    def convert_file(self, file_path, output_format):
        """转换单个文件并写入输出目录"""
        try:
            # 配置 pandoc 路径（支持打包后的环境）
            import os
            meipass = getattr(sys, '_MEIPASS', None)
            if meipass:
                # 打包后的环境，使用内置的 pandoc
                pandoc_exe = Path(meipass) / 'pypandoc' / 'files' / 'pandoc.exe'
                if pandoc_exe.exists():
                    os.environ['PYPANDOC_PANDOC'] = str(pandoc_exe)
                    # 强制 pypandoc 重新查找 pandoc
                    import pypandoc
                    pypandoc._pandoc_path = None  # 清除缓存
                    self.log_message(f"  使用内置 Pandoc: {pandoc_exe.name}")
            
            # 导入 md_exporter 的服务模块
            from md_exporter.services import (
                svc_md_to_docx, svc_md_to_pdf, svc_md_to_html,
                svc_md_to_pptx, svc_md_to_xlsx, svc_md_to_csv,
                svc_md_to_json, svc_md_to_xml, svc_md_to_latex,
                svc_md_to_ipynb, svc_md_to_md
            )
            
            # 读取 Markdown 文件内容
            md_text = Path(file_path).read_text(encoding='utf-8')
            stem = Path(file_path).stem
            output_file = Path(self.output_dir.get()) / f"{stem}{OUTPUT_FORMATS[output_format][1]}"
            
            self.log_message(f"  → 准备保存到: {output_file.name}")
            
            # 检查文件是否已存在
            if output_file.exists() and not self._ask_overwrite(output_file.name):
                self.log_message(f"  ✗ 已跳过: {output_file.name}");  return
            
            # 根据选择的格式调用相应的服务
            service_map = {
                'DOCX': lambda: svc_md_to_docx.convert_md_to_docx(md_text, output_file),
                'PDF': lambda: svc_md_to_pdf.convert_md_to_pdf(md_text, output_file),
                'HTML': lambda: svc_md_to_html.convert_md_to_html(md_text, output_file),
                'PPTX': lambda: svc_md_to_pptx.convert_md_to_pptx(md_text, output_file),
                'XLSX': lambda: svc_md_to_xlsx.convert_md_to_xlsx(md_text, output_file),
                'CSV': lambda: svc_md_to_csv.convert_md_to_csv(md_text, output_file),
                'JSON': lambda: svc_md_to_json.convert_md_to_json(md_text, output_file),
                'XML': lambda: svc_md_to_xml.convert_md_to_xml(md_text, output_file),
                'LaTeX': lambda: svc_md_to_latex.convert_md_to_latex(md_text, output_file),
                'IPYNB': lambda: svc_md_to_ipynb.convert_md_to_ipynb(md_text, output_file),
                'MD': lambda: svc_md_to_md.convert_md_to_md(md_text, output_file),
            }
            
            converter = service_map.get(output_format)
            if converter:
                converter()
                self.last_output_file = str(output_file)
            else:
                raise ValueError(f"不支持的输出格式: {output_format}")
                
        except ImportError as e:
            raise RuntimeError(f"模块导入失败: {e}\n请确保已安装 md-exporter 包")
        except Exception as e:
            raise RuntimeError(f"转换文件 {file_path} 失败: {e}")

    def processing_complete(self):
        self.is_processing = False
        self.process_button.configure(state='normal')

    # ── 对话框（委托给 _dialogs 模块）────────────────────────────────────────

    def _ask_overwrite(self, filename):
        return ask_overwrite(self, filename)

    def show_about(self):
        show_about(self)
