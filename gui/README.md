# Markdown Exporter GUI

图形界面版本的 Markdown Exporter，提供简单易用的可视化操作界面。

## 功能特性

- 🎨 **友好的图形界面** - 无需命令行操作，直观易用
- 📁 **批量文件处理** - 一次转换多个 Markdown 文件
- 📊 **多种输出格式** - 支持 DOCX、PDF、HTML、PPTX、XLSX 等 11 种格式
- 🔒 **本地处理** - 所有转换在本地完成，保护您的隐私
- ⚡ **快速转换** - 基于成熟的 md-exporter 核心引擎

## 安装要求

### Python 环境

- Python 3.11 或更高版本
- tkinter（Python 标准库，Windows/macOS 已内置）

**Linux 用户需要额外安装 tkinter：**

```bash
# Ubuntu/Debian
sudo apt-get install python3-tk

# Fedora
sudo dnf install python3-tkinter

# CentOS/RHEL
sudo yum install python3-tkinter
```

### 依赖安装

```bash
# 安装项目依赖
pip install -r requirements.txt

# 或者使用 uv（推荐）
uv sync
```

## 使用方法

### 方式一：直接运行启动脚本

```bash
python run_gui.py
```

### 方式二：从 gui 目录运行

```bash
cd gui
python main.py
```

### 方式三：作为模块运行

```bash
python -m gui.main
```

## 使用步骤

1. **选择 Markdown 文件**

   - 点击"选择文件"按钮
   - 可以选择一个或多个 .md 文件

2. **选择保存位置**

   - 点击"保存位置"按钮
   - 选择转换后文件的输出目录

3. **选择输出格式**

   - 从下拉菜单中选择目标格式
   - 支持的格式：
     - Word 文档 (.docx)
     - PDF 文档 (.pdf)
     - HTML 网页 (.html)
     - PowerPoint (.pptx)
     - Excel 表格 (.xlsx)
     - CSV 数据 (.csv)
     - JSON 数据 (.json)
     - XML 数据 (.xml)
     - LaTeX 文档 (.tex)
     - Jupyter Notebook (.ipynb)
     - Markdown 文件 (.md)

4. **开始转换**

   - 点击"▶ 开始转换"按钮
   - 在日志区域查看转换进度

5. **查看结果**
   - 点击"📂 打开输出目录"按钮
   - 自动打开文件夹并选中最新生成的文件

## 界面说明

### 主要区域

- **文件选择区** - 选择要转换的 Markdown 文件
- **保存位置区** - 设置输出文件的保存目录
- **格式选择区** - 选择目标输出格式
- **操作按钮区** - 开始转换和打开输出目录
- **日志显示区** - 实时显示转换过程和结果

### 日志颜色说明

- 🟢 绿色 - 成功信息
- 🔴 红色 - 错误信息
- 🔵 蓝色 - 一般信息
- 🟡 黄色 - 箭头提示
- ⚪ 白色 - 普通文本

## 注意事项

1. **文件覆盖**

   - 如果输出文件已存在，会弹出确认对话框
   - 批量处理时可选择"全部覆盖"或"全部跳过"

2. **大文件处理**

   - 转换大文件时请耐心等待
   - 程序不会卡死，后台线程正在处理

3. **特殊格式要求**

   - PPTX 转换需要遵循 Pandoc 幻灯片语法
   - XLSX/CSV 转换主要针对 Markdown 表格

4. **模板支持**
   - DOCX 和 PPTX 支持自定义模板
   - 可通过命令行版本添加模板参数

## 故障排除

### 问题：无法启动 GUI

**解决方案：**

1. 确认 Python 版本 >= 3.11
2. 确认已安装所有依赖：`pip install -r requirements.txt`
3. Linux 用户确认已安装 tkinter

### 问题：导入 md_exporter 失败

**解决方案：**

```bash
# 确保在项目根目录运行
cd d:\Code\markdown-exporter-gui

# 重新安装依赖
pip install -e .
```

### 问题：转换失败

**解决方案：**

1. 检查日志区域的错误信息
2. 确认 Markdown 文件格式正确
3. 尝试使用命令行版本验证文件是否正常

## 与命令行版本的对比

| 特性       | GUI 版本   | 命令行版本 |
| ---------- | ---------- | ---------- |
| 易用性     | ⭐⭐⭐⭐⭐ | ⭐⭐⭐     |
| 批量处理   | ✅         | ✅         |
| 自定义参数 | 基础       | 完整       |
| 模板支持   | ❌         | ✅         |
| 自动化脚本 | ❌         | ✅         |
| 学习成本   | 低         | 中         |

## 开发说明

### 项目结构

```
gui/
├── main.py          # GUI 入口文件
├── _app.py          # 主应用类
├── _dialogs.py      # 对话框组件
└── _version.py      # 版本信息
```

### 添加新格式支持

1. 在 `_app.py` 的 `OUTPUT_FORMATS` 字典中添加新格式
2. 在 `convert_file` 方法的 `service_map` 中添加对应的服务调用
3. 确保相应的 service 模块已导入

### 打包为可执行文件

可以使用 PyInstaller 打包：

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name="MarkdownExporter" --icon=_assets/icon.png run_gui.py
```

## 相关链接

- GitHub: https://github.com/bowenliang123/markdown-exporter
- PyPI: https://pypi.org/project/md-exporter/
- Dify Market: https://marketplace.dify.ai/plugins/bowenliang123/md_exporter

## 许可证

本项目采用 Apache License 2.0 许可证。
