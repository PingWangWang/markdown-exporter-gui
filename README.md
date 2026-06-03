# Markdown Exporter

> 一个桌面 GUI 工具，将 Markdown 文件批量转换为 DOCX、PDF、HTML 格式。无需命令行，拖拽文件即可转换。

[![Version](https://img.shields.io/badge/version-3.6.9-blue)](https://github.com/pingwang1994/markdown-exporter-gui/releases)
[![License](https://img.shields.io/badge/license-Apache--2.0-green)](LICENSE)
[![Python](https://img.shields.io/badge/python-%3E%3D3.11-3776AB)](https://www.python.org/)

## 适用场景

- 需要把 Markdown 笔记转成 Word 文档发给同事
- 想快速生成排版整齐的 PDF 报告
- 批量转换多个 Markdown 文件，不想逐个手动操作
- 不熟悉命令行，希望有一个简单的图形界面

## 功能特性

- **批量转换**：一次选择多个 `.md` / `.markdown` 文件，一键全部转换
- **三种输出格式**：DOCX（Word）、PDF、HTML
- **拖拽支持**：直接把文件拖进窗口即可添加（需安装 tkinterdnd2）
- **自定义 DOCX 模板**：支持导入自己的 `.docx` 模板文件
- **Mermaid 图片转换**：可选将 Markdown 中的 Mermaid 图表渲染为图片嵌入文档
- **文件冲突处理**：输出文件已存在时可选择覆盖、跳过，支持"全部覆盖/全部跳过"
- **文件占用检测**：目标文件被其他程序打开时会提示，而不是直接报错
- **单文件 EXE 打包**：可打包为独立 exe，无需安装 Python 即可运行

## 快速开始

### 环境要求

- Python 3.11+
- uv（推荐）或 pip

### 安装依赖

```bash
# 安装 uv（如果没有）
pip install uv

# 同步依赖（自动创建虚拟环境）
uv sync
```

### 启动应用

```bash
uv run python run_gui.py
```

### 使用步骤

1. 点击「添加文件」选择 Markdown 文件（或直接拖拽文件到窗口）
2. 点击「保存位置」选择输出目录
3. 选择输出格式（DOCX / PDF / HTML）
4. 点击「开始转换」

## 项目结构

```
markdown-exporter-gui/
├── gui/                    # GUI 应用源码
│   ├── main.py            # 程序入口
│   ├── _app.py            # 主窗口类 MarkdownExporterGUI
│   ├── _dialogs.py        # 对话框（关于、文件覆盖确认、文件锁定检测）
│   └── _version.py        # 版本号
├── res/                    # 资源文件（图标）
├── build/                  # 打包相关
│   ├── build_exe.py       # PyInstaller 打包脚本
│   └── README_PACKAGING.md # 打包详细指南
├── pyproject.toml         # 项目配置与依赖定义
├── uv.lock                # 依赖版本锁
└── run_gui.py             # 启动脚本
```

## 打包为 EXE

```bash
uv sync
uv run python build/build_exe.py
```

输出：`dist/MarkdownExporter_v3.6.9_YYYYMMDD-HHMMSS.exe`（约 90-100 MB，单文件免安装）

详细说明见 [build/README_PACKAGING.md](build/README_PACKAGING.md)。

## 开发

```bash
# 代码检查
uv run ruff check gui/

# 自动修复
uv run ruff check --fix gui/

# 运行测试
uv run pytest
```

## 依赖说明

| 包 | 用途 |
|---|------|
| `markdown` | Markdown 解析 |
| `python-docx` | 生成 DOCX |
| `xhtml2pdf` / `reportlab` | 生成 PDF |
| `pypandoc-binary` | Markdown 转换引擎（内置 Pandoc） |
| `pillow` | 图片处理 |
| `pandas` | 数据处理（表格导出） |
| `tkinterdnd2` | 拖拽文件支持 |
| `requests` | 网络请求（Mermaid 图片渲染） |

## License

[Apache-2.0](LICENSE)

## 作者

pingwang1994
