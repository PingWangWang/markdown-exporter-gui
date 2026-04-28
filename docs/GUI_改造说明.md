# Markdown Exporter GUI 改造完成说明

## 改造概述

已成功将 markitdown 的 GUI 界面改造为适配 markdown-exporter 项目的图形界面。

## 主要改动

### 1. 文件结构

```
gui/
├── main.py          # GUI 入口，启动 MarkdownExporterGUI
├── _app.py          # 主应用类，包含所有界面和转换逻辑
├── _dialogs.py      # 对话框组件（关于窗口、覆盖确认）
├── _version.py      # 版本信息（固定为 3.6.9）
└── README.md        # GUI 使用说明文档
```

### 2. 核心功能改造

#### \_version.py

- 简化版本检测逻辑
- 直接使用固定版本号 "3.6.9"
- 移除复杂的版本检测代码

#### \_app.py

主要改动：

- 类名从 `MarkItDownGUI` 改为 `MarkdownExporterGUI`
- 窗口标题改为 "Markdown Exporter v{version}"
- **输入文件类型**：从多种格式改为仅支持 Markdown (.md) 文件
- **输出格式**：新增下拉选择框，支持 11 种输出格式：
  - DOCX (Word 文档)
  - PDF (PDF 文档)
  - HTML (HTML 网页)
  - PPTX (PowerPoint)
  - XLSX (Excel 表格)
  - CSV (CSV 数据)
  - JSON (JSON 数据)
  - XML (XML 数据)
  - LaTeX (LaTeX 文档)
  - IPYNB (Jupyter Notebook)
  - MD (Markdown 文件)
- **转换逻辑**：集成 md_exporter 的服务模块
  - 导入所有 svc*md_to*\* 服务
  - 根据用户选择的格式调用对应的服务
  - 读取 Markdown 文件内容并转换为目标格式
- **移除图片处理选项**：原 markitdown 的图片处理方式不再需要
- **图标路径**：更新为 `_assets/icon.png`

#### \_dialogs.py

- 更新"关于"窗口内容
- 显示 markdown-exporter 项目信息
- 列出所有支持的输出格式
- 更新项目链接

#### main.py

- 更新注释和导入
- 使用 `MarkdownExporterGUI` 类

### 3. 新增文件

#### run_gui.py

- 项目根目录的 GUI 启动脚本
- 自动添加 gui 目录到 Python 路径
- 使用方法：`python run_gui.py`

#### gui/README.md

- 详细的 GUI 使用说明
- 安装要求和方法
- 使用步骤说明
- 故障排除指南
- 开发说明

### 4. 更新的配置

#### requirements.txt

- 添加 GUI 依赖说明
- 注明 tkinter 在 Linux 上需要单独安装

#### README.md

- 在使用方式表格中添加"图形界面 (GUI)"一行
- 提供 GUI 的安装和运行说明

## 技术实现细节

### 转换流程

1. 用户选择 Markdown 文件
2. 用户选择保存位置
3. 用户选择输出格式
4. 点击"开始转换"
5. 后台线程执行转换：
   - 读取 Markdown 文件内容
   - 根据格式调用对应的 service 模块
   - 写入输出文件
   - 更新日志显示

### 服务映射

```python
service_map = {
    'DOCX': svc_md_to_docx.convert_md_to_docx,
    'PDF': svc_md_to_pdf.convert_md_to_pdf,
    'HTML': svc_md_to_html.convert_md_to_html,
    'PPTX': svc_md_to_pptx.convert_md_to_pptx,
    'XLSX': svc_md_to_xlsx.convert_md_to_xlsx,
    'CSV': svc_md_to_csv.convert_md_to_csv,
    'JSON': svc_md_to_json.convert_md_to_json,
    'XML': svc_md_to_xml.convert_md_to_xml,
    'LaTeX': svc_md_to_latex.convert_md_to_latex,
    'IPYNB': svc_md_to_ipynb.convert_md_to_ipynb,
    'MD': svc_md_to_md.convert_md_to_md,
}
```

### 多线程处理

- 使用 `threading.Thread` 在后台执行转换
- 避免界面卡死
- 通过 `root.after()` 安全地更新 UI

## 使用前提条件

1. **Python 版本**：3.11 或更高
2. **依赖安装**：
   ```bash
   pip install -r requirements.txt
   ```
3. **tkinter**：
   - Windows/macOS：已内置
   - Linux：需要安装 `python3-tk`

## 运行方法

```bash
# 方法一：使用启动脚本
python run_gui.py

# 方法二：直接从 gui 目录运行
cd gui
python main.py

# 方法三：作为模块运行
python -m gui.main
```

## 与原版对比

| 特性     | markitdown GUI              | markdown-exporter GUI |
| -------- | --------------------------- | --------------------- |
| 输入格式 | 多种（PDF、Word、Excel 等） | 仅 Markdown           |
| 输出格式 | 仅 Markdown                 | 11 种格式             |
| 转换方向 | 多格式 → Markdown           | Markdown → 多格式     |
| 图片处理 | 支持多种模式                | 由底层服务处理        |
| 核心引擎 | markitdown                  | md-exporter           |

## 注意事项

1. **模板支持**：当前 GUI 版本不支持自定义模板（DOCX/PPTX），如需使用模板请使用命令行版本
2. **高级参数**：GUI 仅提供基本功能，高级参数（如 force-text、strip-wrapper 等）需使用命令行
3. **批量处理**：支持同时转换多个 Markdown 文件
4. **文件覆盖**：批量处理时可选择"全部覆盖"或"全部跳过"

## 测试状态

✅ 语法检查通过（Python 3.14.3）
✅ 版本信息加载成功
✅ 模块导入正常

⏳ 完整功能测试需要安装依赖后运行

## 后续优化建议

1. **添加模板选择**：为 DOCX 和 PPTX 添加模板文件选择功能
2. **进度条**：添加可视化进度条显示转换进度
3. **预览功能**：添加 Markdown 内容预览
4. **最近使用**：记录最近使用的文件和目录
5. **拖拽支持**：支持拖拽文件到窗口
6. **打包发布**：使用 PyInstaller 打包为独立可执行文件

## 总结

GUI 改造已完成，核心功能已实现：

- ✅ 界面适配 markdown-exporter
- ✅ 集成所有转换服务
- ✅ 支持 11 种输出格式
- ✅ 批量文件处理
- ✅ 友好的用户界面
- ✅ 完整的文档说明

用户现在可以通过图形界面轻松地将 Markdown 文件转换为各种格式，无需使用命令行。
