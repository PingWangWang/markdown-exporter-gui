# Markdown Exporter GUI 打包指南

## 📋 环境要求

- Windows 10/11 x64
- Python 3.11 或以上（建议 3.12+），需加入 PATH
- 网络畅通（首次安装依赖时需要）

---

## 📦 第一步：安装依赖

在新电脑上首次打包前，需安装以下所有 Python 库。

### 1. 打包工具（必须）

```bash
pip install pyinstaller>=6.0.0
```

### 2. Markdown Exporter 核心及其依赖

使用项目提供的 gui_requirements.txt：

```bash
cd d:\Code\markdown-exporter-gui
pip install -r gui_requirements.txt
```

或者手动安装：

```bash
pip install markdown~=3.10.2
pip install pandas[excel,html,xml]~=3.0.1
pip install xhtml2pdf~=0.2.17
pip install pillow~=12.1.0
pip install pypandoc-binary~=1.16.2
```

### 一键安装（复制整段执行）

```bash
pip install ^
  pyinstaller ^
  markdown~=3.10.2 ^
  pandas[excel,html,xml]~=3.0.1 ^
  xhtml2pdf~=0.2.17 ^
  pillow~=12.1.0 ^
  pypandoc-binary~=1.16.2
```

---

## 🚀 第二步：执行打包

```bash
cd d:\Code\markdown-exporter-gui
py build\build_exe.py
```

> **注意**：若终端中文乱码，请先执行：
>
> ```powershell
> $OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
> ```

---

## 📂 第三步：获取产物

打包完成后，exe 在项目根目录的 `dist/` 下：

```
d:\Code\markdown-exporter-gui\
├── build\                 ← 临时构建文件（打包后可删除）
└── dist\
    └── MarkdownExporter_v3.6.9_20250427-xxxxxx.exe  ← 直接发给对方即可
```

**单文件模式，对方无需安装 Python，直接运行 exe 即可。**

---

## ⚠️ 常见问题

| 现象                          | 原因                     | 解决方法                                                                                    |
| ----------------------------- | ------------------------ | ------------------------------------------------------------------------------------------- |
| 终端中文乱码                  | PowerShell 默认 GBK 编码 | 执行 `$OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8`           |
| 对方电脑运行报 VCRUNTIME 错误 | 缺少 Visual C++ 运行库   | 对方安装 [Visual C++ Redistributable 2019+](https://aka.ms/vs/17/release/vc_redist.x64.exe) |
| 找不到 pandoc                 | 系统未安装 Pandoc        | 从 https://pandoc.org/installing.html 下载并安装                                            |
| 图标文件不存在                | res/icad.ico 文件缺失    | 确保项目根目录下有 res/icad.ico 文件                                                        |
| 转换失败                      | 缺少必要的转换器依赖     | 检查是否安装了所有 gui_requirements.txt 中的依赖                                            |

---

## 📝 打包说明

### 包含的功能

- ✅ Markdown → DOCX（Word 文档）
- ✅ Markdown → PDF（PDF 文档）
- ✅ Markdown → HTML（网页）
- ✅ Markdown → PPTX（PowerPoint）
- ✅ Markdown → XLSX（Excel）
- ✅ Markdown → CSV
- ✅ Markdown → JSON
- ✅ Markdown → XML
- ✅ Markdown → LaTeX
- ✅ Markdown → IPYNB（Jupyter Notebook）
- ✅ Markdown → MD（保存为文件）

### 注意事项

1. **Pandoc 依赖**：DOCX、PPTX、HTML 等格式需要系统安装 Pandoc

   - 下载地址：https://pandoc.org/installing.html
   - 打包时会自动包含 pypandoc，但目标机器仍需安装 Pandoc

2. **文件大小**：单文件模式打包后约 50-100MB（包含所有依赖）

3. **模板支持**：当前 GUI 版本暂不支持自定义 DOCX/PPTX 模板，如需使用请使用命令行版本

4. **图片处理**：GUI 自动处理图片转换，无需额外配置
