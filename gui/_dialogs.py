"""
Markdown Exporter GUI - 对话框组件

包含：
  - show_about(app)        关于窗口
  - ask_overwrite(app, filename)  文件覆盖确认窗口
  - is_file_locked(filepath)      检测文件是否被占用
  - ask_file_locked(app, filename) 文件被占用提示窗口
"""

import os
import platform
import subprocess
import sys
import threading
import tkinter as tk
import webbrowser

from _version import APP_VERSION


def is_file_locked(filepath):
    """检测文件是否被其他程序占用（打开）

    Returns:
        bool: True 表示文件被占用，False 表示未被占用
    """
    if not os.path.exists(filepath):
        return False

    try:
        # 尝试以独占模式重命名文件（Windows 和 Unix 都适用）
        # 如果文件被占用，重命名会失败
        test_path = filepath + ".lock_test"
        try:
            os.rename(filepath, test_path)
            # 重命名成功，立即改回来
            os.rename(test_path, filepath)
            return False
        except OSError:
            # 重命名失败，可能是文件被占用
            # 再尝试删除测试文件（如果存在）
            if os.path.exists(test_path):
                try:
                    os.remove(test_path)
                except Exception:
                    pass
            return True
    except Exception:
        # 其他异常，保守起见认为未被占用
        return False


def show_about(app):
    """显示关于信息（自定义风格）"""
    dlg = tk.Toplevel(app.root)
    dlg.overrideredirect(True)
    dlg.configure(bg=app.C_BG)
    dlg.resizable(False, False)

    # 标题栏
    header = tk.Frame(dlg, bg=app.C_HEADER_BG, height=46)
    header.pack(fill=tk.X)
    header.pack_propagate(False)
    tk.Label(
        header,
        text=f"关于  Markdown Exporter v{APP_VERSION}",
        bg=app.C_HEADER_BG,
        fg=app.C_HEADER_FG,
        font=("Microsoft YaHei UI", 12, "bold"),
    ).pack(side=tk.LEFT, padx=16, pady=8)

    # 内容区
    body = tk.Frame(dlg, bg=app.C_BG, padx=24, pady=16)
    body.pack(fill=tk.BOTH)

    sections = [
        (
            "项目来源",
            [
                f"版本: {APP_VERSION}",
                "作者: bowenliang123",
                "GitHub: https://github.com/bowenliang123/markdown-exporter",
            ],
        ),
        (
            "详细文档",
            [
                "点击查看 README.md 获取完整使用说明和示例",
            ],
        ),
    ]

    for title, items in sections:
        tk.Label(body, text=title, bg=app.C_BG, fg=app.C_HEADER_BG, font=("Microsoft YaHei UI", 10, "bold")).pack(
            anchor=tk.W, pady=(8, 2)
        )
        for item in items:
            # 检查是否为 URL 文本（包含 http:// 或 https://）
            if "http://" in item or "https://" in item:
                # 提取 URL
                url_start = item.find("http")
                prefix = item[:url_start].rstrip(": ").rstrip()
                url = item[url_start:]

                # 创建容器框架来保持整行内容一起
                item_frame = tk.Frame(body, bg=app.C_BG)
                item_frame.pack(fill=tk.X, anchor=tk.W, pady=1)

                # 创建前缀文本（如果有）
                if prefix:
                    tk.Label(
                        item_frame,
                        text=f"  • {prefix}: ",
                        bg=app.C_BG,
                        fg=app.C_LABEL_FG,
                        font=("Microsoft YaHei UI", 9),
                        justify="left",
                    ).pack(side=tk.LEFT, anchor=tk.W)
                else:
                    tk.Label(
                        item_frame,
                        text="  • ",
                        bg=app.C_BG,
                        fg=app.C_LABEL_FG,
                        font=("Microsoft YaHei UI", 9),
                        justify="left",
                    ).pack(side=tk.LEFT, anchor=tk.W)

                # 创建可点击的链接标签
                link_label = tk.Label(
                    item_frame,
                    text=url,
                    bg=app.C_BG,
                    fg="#1E90FF",
                    font=("Microsoft YaHei UI", 9, "underline"),
                    cursor="hand2",
                    justify="left",
                )
                link_label.pack(side=tk.LEFT, anchor=tk.W)
                link_label.bind("<Button-1>", lambda e, u=url: webbrowser.open(u))
                link_label.bind("<Enter>", lambda e: e.widget.config(fg="#4169E1"))
                link_label.bind("<Leave>", lambda e: e.widget.config(fg="#1E90FF"))
            # 检查是否为 README 链接
            elif "README.md" in item:
                item_frame = tk.Frame(body, bg=app.C_BG)
                item_frame.pack(fill=tk.X, anchor=tk.W, pady=1)
                
                tk.Label(
                    item_frame,
                    text="  • ",
                    bg=app.C_BG,
                    fg=app.C_LABEL_FG,
                    font=("Microsoft YaHei UI", 9),
                    justify="left",
                ).pack(side=tk.LEFT, anchor=tk.W)
                
                readme_link = tk.Label(
                    item_frame,
                    text="查看 README.md",
                    bg=app.C_BG,
                    fg="#1E90FF",
                    font=("Microsoft YaHei UI", 9, "underline"),
                    cursor="hand2",
                    justify="left",
                )
                readme_link.pack(side=tk.LEFT, anchor=tk.W)
                # 打开本地的 README.md 文件
                def open_readme(e):
                    try:
                        # 获取 README.md 的路径
                        if getattr(sys, 'frozen', False):
                            # PyInstaller 打包后的路径 - 文件在 _MEIPASS 临时目录
                            base_dir = sys._MEIPASS
                        else:
                            # 开发环境路径 - 项目根目录
                            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
                        
                        readme_path = os.path.join(base_dir, 'README.md')
                        
                        # 调试信息（可选）
                        # print(f"Looking for README at: {readme_path}")
                        # print(f"File exists: {os.path.exists(readme_path)}")
                        
                        if os.path.exists(readme_path):
                            # 根据操作系统打开文件
                            system = platform.system()
                            if system == 'Windows':
                                os.startfile(readme_path)
                            elif system == 'Darwin':  # macOS
                                subprocess.call(['open', readme_path])
                            else:  # Linux
                                subprocess.call(['xdg-open', readme_path])
                        else:
                            # 如果本地文件不存在，打开 GitHub URL
                            webbrowser.open("https://github.com/bowenliang123/markdown-exporter#readme")
                    except Exception as ex:
                        # 出错时回退到 GitHub URL
                        # print(f"Error opening README: {ex}")
                        webbrowser.open("https://github.com/bowenliang123/markdown-exporter#readme")
                
                readme_link.bind("<Button-1>", open_readme)
                readme_link.bind("<Enter>", lambda e: e.widget.config(fg="#4169E1"))
                readme_link.bind("<Leave>", lambda e: e.widget.config(fg="#1E90FF"))
            else:
                tk.Label(
                    body,
                    text=f"  • {item}",
                    bg=app.C_BG,
                    fg=app.C_LABEL_FG,
                    font=("Microsoft YaHei UI", 9),
                    justify="left",
                ).pack(anchor=tk.W, pady=1)

    # 底部
    tk.Frame(dlg, bg=app.C_BORDER, height=1).pack(fill=tk.X, pady=(8, 0))
    btn_frame = tk.Frame(dlg, bg=app.C_BG, pady=10)
    btn_frame.pack()
    ok_btn = tk.Button(
        btn_frame,
        text="确  定",
        width=10,
        bg=app.C_BTN_SEL,
        fg="#FFFFFF",
        relief="flat",
        font=("Microsoft YaHei UI", 9, "bold"),
        cursor="hand2",
        command=dlg.destroy,
    )
    ok_btn.pack()
    ok_btn.bind("<Enter>", lambda e: ok_btn.config(bg=app.C_BTN_SEL_A))
    ok_btn.bind("<Leave>", lambda e: ok_btn.config(bg=app.C_BTN_SEL))

    # 居中
    dlg.update_idletasks()
    w, h = dlg.winfo_width(), dlg.winfo_height()
    rx = app.root.winfo_x() + (app.root.winfo_width() - w) // 2
    ry = app.root.winfo_y() + (app.root.winfo_height() - h) // 2
    dlg.geometry(f"+{rx}+{ry}")
    dlg.grab_set()


def ask_overwrite(app, filename):
    """在主线程弹出文件覆盖确认对话框，返回 True（覆盖）或 False（跳过）"""
    if getattr(app, "_overwrite_all", False):
        return True
    if getattr(app, "_skip_all", False):
        return False

    result = [False]
    event = threading.Event()
    is_multi = len(app.input_files) > 1

    def _show():
        dlg = tk.Toplevel(app.root)
        dlg.overrideredirect(True)
        dlg.configure(bg=app.C_BG)
        dlg.resizable(False, False)

        # 标题栏
        header = tk.Frame(dlg, bg="#E67E22", height=36)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        tk.Label(header, text="文件已存在", bg="#E67E22", fg="#FFFFFF", font=("Microsoft YaHei UI", 10, "bold")).pack(
            side=tk.LEFT, padx=12, pady=6
        )

        # 内容区
        body = tk.Frame(dlg, bg=app.C_BG, padx=20, pady=16)
        body.pack(fill=tk.BOTH)
        tk.Label(
            body,
            text=f"「{filename}」已存在，是否覆盖？",
            bg=app.C_BG,
            fg=app.C_LABEL_FG,
            font=("Microsoft YaHei UI", 10),
            wraplength=340,
            justify="left",
        ).pack(anchor=tk.W)

        btn_frame = tk.Frame(dlg, bg=app.C_BG, pady=10)
        btn_frame.pack()

        def on_overwrite_one():
            result[0] = True
            dlg.destroy()

        def on_overwrite_all():
            app._overwrite_all = True
            result[0] = True
            dlg.destroy()

        def on_skip():
            result[0] = False
            dlg.destroy()

        def on_skip_all():
            app._skip_all = True
            result[0] = False
            dlg.destroy()

        def _btn(parent, text, bg, bg_hover, cmd, padx=6):
            b = tk.Button(
                parent,
                text=text,
                width=8,
                bg=bg,
                fg="#FFFFFF",
                relief="flat",
                font=("Microsoft YaHei UI", 9, "bold"),
                cursor="hand2",
                command=cmd,
            )
            b.pack(side=tk.LEFT, padx=padx)
            b.bind("<Enter>", lambda e: b.config(bg=bg_hover))
            b.bind("<Leave>", lambda e: b.config(bg=bg))
            return b

        if is_multi:
            _btn(btn_frame, "本次覆盖", app.C_BTN_RUN, app.C_BTN_RUN_A, on_overwrite_one)
            _btn(btn_frame, "全部覆盖", "#8E44AD", "#6C3483", on_overwrite_all)
            _btn(btn_frame, "本次跳过", app.C_BTN_SEL, app.C_BTN_SEL_A, on_skip)
            _btn(btn_frame, "全部跳过", "#7F8C8D", "#626567", on_skip_all)
        else:
            _btn(btn_frame, "覆  盖", app.C_BTN_RUN, app.C_BTN_RUN_A, on_overwrite_one, padx=8)
            _btn(btn_frame, "跳  过", app.C_BTN_SEL, app.C_BTN_SEL_A, on_skip, padx=8)

        # 居中
        dlg.update_idletasks()
        w, h = dlg.winfo_width(), dlg.winfo_height()
        rx = app.root.winfo_x() + (app.root.winfo_width() - w) // 2
        ry = app.root.winfo_y() + (app.root.winfo_height() - h) // 2
        dlg.geometry(f"+{rx}+{ry}")
        dlg.grab_set()
        dlg.wait_window()
        event.set()

    app.root.after(0, _show)
    event.wait()
    return result[0]


def ask_file_locked(app, filename):
    """在主线程弹出文件被占用提示对话框，返回 True（关闭文件并重试）或 False（跳过）"""
    result = [False]
    event = threading.Event()

    def _show():
        dlg = tk.Toplevel(app.root)
        dlg.overrideredirect(True)
        dlg.configure(bg=app.C_BG)
        dlg.resizable(False, False)

        # 标题栏
        header = tk.Frame(dlg, bg="#E74C3C", height=36)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        tk.Label(header, text="文件被占用", bg="#E74C3C", fg="#FFFFFF", font=("Microsoft YaHei UI", 10, "bold")).pack(
            side=tk.LEFT, padx=12, pady=6
        )

        # 内容区
        body = tk.Frame(dlg, bg=app.C_BG, padx=20, pady=16)
        body.pack(fill=tk.BOTH)

        tk.Label(
            body,
            text=f"「{filename}」正在被其他程序打开，\n无法覆盖保存。",
            bg=app.C_BG,
            fg=app.C_LABEL_FG,
            font=("Microsoft YaHei UI", 10),
            wraplength=340,
            justify="center",
        ).pack(anchor=tk.CENTER, pady=(0, 8))

        tk.Label(
            body,
            text="请关闭该文件后重试，或选择跳过。",
            bg=app.C_BG,
            fg="#7F8C8D",
            font=("Microsoft YaHei UI", 9),
            wraplength=340,
            justify="center",
        ).pack(anchor=tk.CENTER)

        btn_frame = tk.Frame(dlg, bg=app.C_BG, pady=10)
        btn_frame.pack()

        def on_retry():
            result[0] = True
            dlg.destroy()

        def on_skip():
            result[0] = False
            dlg.destroy()

        def _btn(parent, text, bg, bg_hover, cmd, padx=6):
            b = tk.Button(
                parent,
                text=text,
                width=8,
                bg=bg,
                fg="#FFFFFF",
                relief="flat",
                font=("Microsoft YaHei UI", 9, "bold"),
                cursor="hand2",
                command=cmd,
            )
            b.pack(side=tk.LEFT, padx=padx)
            b.bind("<Enter>", lambda e: b.config(bg=bg_hover))
            b.bind("<Leave>", lambda e: b.config(bg=bg))
            return b

        _btn(btn_frame, "关闭后重试", app.C_BTN_RUN, app.C_BTN_RUN_A, on_retry, padx=8)
        _btn(btn_frame, "跳  过", app.C_BTN_SEL, app.C_BTN_SEL_A, on_skip, padx=8)

        # 居中
        dlg.update_idletasks()
        w, h = dlg.winfo_width(), dlg.winfo_height()
        rx = app.root.winfo_x() + (app.root.winfo_width() - w) // 2
        ry = app.root.winfo_y() + (app.root.winfo_height() - h) // 2
        dlg.geometry(f"+{rx}+{ry}")
        dlg.grab_set()
        dlg.wait_window()
        event.set()

    app.root.after(0, _show)
    event.wait()
    return result[0]
