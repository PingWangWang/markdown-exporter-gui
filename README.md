<div align="center">
  <img src="https://raw.githubusercontent.com/bowenliang123/markdown-exporter/main/_assets/icon.png" alt="Markdown Exporter Logo" width="200">
</div>
<p align="center">
  <a href="https://github.com/bowenliang123/markdown-exporter" target="_blank">
      <img alt="Github" src="https://img.shields.io/badge/bowenliang123-markdown--exporter-lightgray?logo=github"></a>
  <a href="https://marketplace.dify.ai/plugin/bowenliang123/md_exporter" target="_blank">
      <img alt="Github" src="https://img.shields.io/badge/Dify-md__exporter-blue"></a>
  <a href="https://clawhub.ai/bowenliang123/markdown-exporter" target="_blank">
      <img alt="Github" src="https://img.shields.io/badge/🦞OpenClaw-markdown--exporter-red"></a>
  <a href="https://pypi.org/project/md-exporter/" target="_blank">
      <img alt="Github" src="https://img.shields.io/badge/PyPI-md--exporter-yellow?logo=python"></a>
</p>

# Markdown Exporter

### 一个用于将 Markdown 导出为强大文档的 Agent Skill 和 Dify 插件

- Author: [bowenliang123](https://github.com/bowenliang123)
- GitHub Repo: [markdown-exporter](https://github.com/bowenliang123/markdown-exporter)

**Markdown Exporter** 可以用作：

| 使用方式           | 平台和安装方法                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| ------------------ | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Dify 插件          | **平台**: [Dify](https://github.com/langgenius/dify) <br/> **安装方法**: <br/> - 从 [Dify 市场](https://marketplace.dify.ai/plugins/bowenliang123/md_exporter) 安装                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      |
| Agent Skills       | **平台**: 任何支持 [Agent Skills](https://agentskills.io) 的平台 <br/> - **IDEs/CLIs**: [Claude Code](https://code.claude.com/docs/en/skills), [Trae](https://docs.trae.ai/ide/skills), [Codebuddy](https://copilot.tencent.com/docs/cli/skills) 等 <br/> - **Agent 框架**: [LangChain DeepAgents](https://www.blog.langchain.com/using-skills-with-deep-agents/), [AgentScope](https://doc.agentscope.io/tutorial/task_agent_skill.html) 等 <br/><br/> **安装方法**: <br/> - **本地导入**: 下载并导入 [源代码 zip](https://github.com/bowenliang123/markdown-exporter/archive/refs/heads/main.zip) <br/> - **远程安装**: 在 agent CLIs 中运行 `/plugin marketplace add bowenliang123/markdown-exporter` |
| OpenClaw Skills 🦞 | **平台**: [OpenClaw](https://docs.openclaw.ai/tools/skills#clawhub-install-%2B-sync) <br/> - 从 [ClawHub](https://clawhub.ai/bowenliang123/markdown-exporter) 安装: `npx clawhub@latest install markdown-exporter`                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| 命令行界面 (CLI)   | **平台**: Python<br/> - 从 [PyPI](https://pypi.org/project/md-exporter/) 安装: `pip install md-exporter`<br/> - 运行: `markdown-exporter -h` 查看使用说明                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| 图形界面 (GUI)     | **平台**: Python + tkinter<br/> - 克隆或下载本项目<br/> - 安装依赖: `pip install -r requirements.txt`<br/> - 运行: `python run_gui.py`                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   |

---

## ✨ 什么是 Markdown Exporter？

**Markdown Exporter** 是一个强大的工具集，作为 Agent Skill 或 Dify 插件，可以将您的 Markdown 文本转换为各种专业格式。无论您需要创建精美的报告、出色的演示文稿、组织有序的电子表格还是代码文件——这个工具都能满足您的需求。

支持 **15+ 种输出格式**，Markdown Exporter 架起了简单文本编辑和专业文档创建之间的桥梁，同时保持了 Markdown 语法的简洁和优雅。

### 🎯 您会喜欢它的原因

- **🚀 极速转换** – 在毫秒内将 Markdown 导出为多种格式
- **🎨 可定制** – 使用自定义 DOCX 和 PPTX 模板以匹配您的品牌
- **🔒 100% 隐私** – 所有处理都在本地进行，数据永远不会离开您的环境
- **📊 多功能** – 从文档到电子表格，从演示文稿到代码文件
- **🌐 多语言支持** – 非常适合国际团队和内容

---

## 🛠️ 介绍和使用指南

### 工具和支持的格式

| 工具 | 输入 | 输出 |
| ---- | ---- | ---- |

  <tr>
    <td><code>md_to_docx</code></td>
    <td rowspan="7">📝 Markdown text</td>
    <td>📄 Word document (.docx)</td>
  </tr>
  <tr>
    <td><code>md_to_html</code></td>
    <td>🌐 HTML file (.html)</td>
  </tr>
  <tr>
    <td><code>md_to_html_text</code></td>
    <td>🌐 HTML text string</td>
  </tr>
  <tr>
    <td><code>md_to_pdf</code></td>
    <td>📑 PDF file (.pdf)</td>
  </tr>
  <tr>
    <td><code>md_to_md</code></td>
    <td>📝 Markdown file (.md)</td>
  </tr>
  <tr>
    <td><code>md_to_ipynb</code></td>
    <td>📓 Jupyter Notebook (.ipynb)</td>
  </tr>
  <tr>
    <td><code>md_to_pptx</code></td>
    <td>
      <div>
        📝 Markdown slides
      </div>
      <div>
      in <a href="https://pandoc.org/MANUAL.html#slide-shows">Pandoc style </a>
      </div>
    </td>
    <td>🎯 PowerPoint (.pptx)</td>
  </tr>
  <tr>
    <td><code>md_to_xlsx</code></td>
    <td rowspan="5">📋<a href="https://www.markdownguide.org/extended-syntax/#tables"> Markdown tables </a> </td>
    <td>📊 Excel spreadsheet (.xlsx)</td>
  </tr>
  <tr>
    <td><code>md_to_csv</code></td>
    <td>📋 CSV file (.csv)</td>
  </tr>
  <tr>
    <td><code>md_to_json</code></td>
    <td>📦 JSON/JSONL file (.json)</td>
  </tr>
  <tr>
    <td><code>md_to_xml</code></td>
    <td>🏷️ XML file (.xml)</td>
  </tr>
  <tr>
    <td><code>md_to_latex</code></td>
    <td>📝 LaTeX file (.tex)</td>
  </tr>
  <tr>
    <td><code>md_to_codeblock</code></td>
    <td>💻 <a href="https://www.markdownguide.org/extended-syntax/#fenced-code-blocks"> Code blocks in Markdown </a> </td>
    <td>📁 Code files by language (.py, .js, .sh, etc.)</td>
  </tr>
</table>

---

## 📖 作为 Dify 插件使用

![使用演示](https://raw.githubusercontent.com/bowenliang123/markdown-exporter/main/_assets/screenshots/usage_md_to_docx.png)

只需输入您的 Markdown 文本，选择所需的输出格式，然后点击导出。就是这么简单！

---

## 🎨 Dify 工具用法

### 📄 Markdown → DOCX

创建具有精美格式的 Word 文档。

> **✨ 专业提示：使用模板自定义样式**
>
> `md_to_docx` 工具支持自定义 DOCX 模板文件，让您完全控制文档的外观。
>
> **您可以自定义的内容：**
>
> - 标题样式（字体、大小、颜色）
> - 段落格式（间距、缩进）
> - 表格样式和边框
> - 列表样式和项目符号
> - 以及更多！
>
> 查看 [默认 docx 模板](https://github.com/bowenliang123/markdown-exporter/blob/main/md_exporter/assets/template/docx_template.docx) 或创建您自己的模板。了解如何操作请访问 [自定义或创建新样式](https://support.microsoft.com/en-us/office/customize-or-create-new-styles-d38d6e47-f6fc-48eb-a607-1eb120dec563)。

![DOCX Example](https://raw.githubusercontent.com/bowenliang123/markdown-exporter/main/_assets/screenshots/md_to_docx_1.png)

---

### 📊 Markdown → XLSX

将您的 Markdown 表格转换为精美的 Excel 电子表格，自动调整列宽并保留数据类型。

**输入：**

```markdown
| 名称    | 年龄 | 城市      |
| ------- | ---- | --------- |
| Alice   | 30   | New York  |
| Bowen   | 25   | Guangzhou |
| Charlie | 35   | Tokyo     |
| David   | 40   | Miami     |
```

**输出：**
![XLSX 示例](https://raw.githubusercontent.com/bowenliang123/markdown-exporter/main/_assets/screenshots/md_to_xlsx_1.png)

---

### 🎯 Markdown → PPTX

自动将您的 Markdown 转换为令人惊叹的 PowerPoint 演示文稿。

> **✨ 语法要求**

> **支持的功能：**
>
> - ✅ 标题幻灯片
> - ✅ 列布局
> - ✅ 表格
> - ✅ 超链接
> - ✅ 以及更多！
>
> > **🎨 自定义模板：**
> >
> > 使用带有幻灯片母版的自定义 PPTX 模板以匹配您品牌的视觉识别。[了解如何操作](https://support.microsoft.com/en-us/office/customize-a-slide-master-036d317b-3251-4237-8ddc-22f4668e2b56)。获取 [默认 pptx 模板](https://github.com/bowenliang123/markdown-exporter/blob/main/md_exporter/assets/template/pptx_template.pptx)。

输入的 Markdown 必须遵循 [Pandoc 幻灯片](https://pandoc.org/MANUAL.html#slide-shows) 中的语法和指导。

**输入示例：**

````markdown
---
title: Markdown Exporter
author: Bowen Liang
---

# Introduction

## Welcome Slide

Welcome to our Markdown Exporter!

::: notes
Remember to greet the audience warmly.
:::

---

# Section 1: Basic Layouts

## Title and Content

- This is a basic slide with bullet points
- It uses the "Title and Content" layout
- Perfect for simple content presentation

## Two Column Layout

::::: columns
::: column
Left column content:

- Point 1
- Point 2
  :::
  ::: column
  Right column content:
- Point A
- Point B
  :::
  :::::

## Comparison Layout

::::: columns
::: column
Text followed by an image:

![Test Image](https://avatars.githubusercontent.com/u/127165244?s=48&v=4)
:::
::: column

- This triggers the "Comparison" layout
- Useful for side-by-side comparisons
  :::
  :::::

## Content with Caption

Here's some explanatory text about the image below.

![Test Image](https://avatars.githubusercontent.com/u/127165244?s=48&v=4 "fig:Test Image")

---

# Section 2: Advanced Features

## Code Block

Here's a Python code block:

```python
def greet(name):
    return f"Hello, {name}!"

print(greet("World"))
```

## Table Example

| Column 1 | Column 2 | Column 3 |
| -------- | -------- | -------- |
| Row 1    | Data     | More     |
| Row 2    | Info     | Stuff    |

## Incremental List

::: incremental

- This point appears first
- Then this one
- And finally this one
  :::

## {background-image="https://avatars.githubusercontent.com/u/127165244?s=48&v=4"}

::: notes
This is a slide with a background image and speaker notes only.
The "Blank" layout will be used.
:::

# Conclusion

## Thank You

Thank you for viewing this kitchen sink presentation!

::: notes
Remember to thank the audience and invite questions.
:::
````

**输出：**
![PPTX 示例](https://raw.githubusercontent.com/bowenliang123/markdown-exporter/main/_assets/screenshots/md_to_pptx_1.png)

---

### 🌐 Markdown → HTML

将您的 Markdown 转换为干净、语义化的 HTML，非常适合网页使用。

![HTML Example](https://raw.githubusercontent.com/bowenliang123/markdown-exporter/main/_assets/screenshots/md_to_html_1.png)

---

### 📑 Markdown → PDF

生成适合打印或分享的专业 PDF 文档。

![PDF Example](https://raw.githubusercontent.com/bowenliang123/markdown-exporter/main/_assets/screenshots/md_to_pdf_1.png)

---

### 🏷️ Markdown → Jupyter Notebook

将您的 Markdown 转换为 Jupyter Notebook `.ipynb` 格式。

**输入示例：**

````markdown
# Example Jupyter Notebook

This is a simplified test markdown file that will be converted to an IPYNB notebook with multiple cells.

## Introduction

This notebook demonstrates the conversion of markdown to IPYNB format.

- It includes markdown formatting
- It has code cells in different languages
- It shows how tables are handled

```python
# Python Code Cell
print("Hello, world!")
x = 10
y = 20
print(f"Sum: {x + y}")
```

## Data Table

Here's a sample table:

| Name  | Score | Occupation |
| ----- | ----- | ---------- |
| Alice | 80    | Engineer   |
| Bowen | 90    | Designer   |

## Conclusion

This concludes the simplified test notebook.
````

**输出文件：**
![IPYNB Example](https://raw.githubusercontent.com/bowenliang123/markdown-exporter/main/_assets/screenshots/md_to_ipynb_1.png)

---

### 💻 Markdown → 代码块文件

轻松从 Markdown 中提取代码块并将它们保存为单独的文件，保留语法高亮和格式。

#### 支持的编程语言和文件扩展名

| Language   | File Extension | Language | File Extension |
| ---------- | -------------- | -------- | -------------- |
| Python     | `.py`          | CSS      | `.css`         |
| JavaScript | `.js`          | YAML     | `.yaml`        |
| HTML       | `.html`        | Ruby     | `.rb`          |
| Bash       | `.sh`          | Java     | `.java`        |
| JSON       | `.json`        | PHP      | `.php`         |
| XML        | `.xml`         | Markdown | `.md`          |
| SVG        | `.svg`         |          |                |

![Code Block Example 1](https://raw.githubusercontent.com/bowenliang123/markdown-exporter/main/_assets/screenshots/usage_md_to_codeblock_2.png)

**专业提示：** 启用压缩功能将所有提取的文件捆绑到一个 ZIP 归档中，便于共享和组织！

![Code Block Example 2](https://raw.githubusercontent.com/bowenliang123/markdown-exporter/main/_assets/screenshots/usage_md_to_codeblock_3.png)
![Code Block Example 3](https://raw.githubusercontent.com/bowenliang123/markdown-exporter/main/_assets/screenshots/usage_md_to_codeblock_4.png)

### 📋 Markdown → CSV

将您的 Markdown 表格导出为通用的 CSV 格式。

![CSV Example](https://raw.githubusercontent.com/bowenliang123/markdown-exporter/main/_assets/screenshots/md_to_csv_1.png)

---

### 📦 Markdown → JSON / JSONL

将您的表格转换为结构化数据格式。

**JSONL 样式（默认）**

- 每行一个 JSON 对象
- 非常适合流式传输和日志记录

![JSONL Example](https://raw.githubusercontent.com/bowenliang123/markdown-exporter/main/_assets/screenshots/md_to_json_2.png)

**JSON 数组样式**

- 所有对象在一个数组中
- 非常适合 API 响应

![JSON Example](https://raw.githubusercontent.com/bowenliang123/markdown-exporter/main/_assets/screenshots/md_to_json_1.png)

---

### 🏷️ Markdown → XML

将您的数据转换为 XML 格式。

![XML Example](https://raw.githubusercontent.com/bowenliang123/markdown-exporter/main/_assets/screenshots/md_to_xml_1.png)

---

### 📝 Markdown → LaTeX

为学术和技术文档生成 LaTeX 源代码。

**LaTeX 输出：**
![LaTeX Example 1](https://raw.githubusercontent.com/bowenliang123/markdown-exporter/main/_assets/screenshots/md_to_latex_1.png)

**编译后的 PDF：**
![LaTeX Example 2](https://raw.githubusercontent.com/bowenliang123/markdown-exporter/main/_assets/screenshots/md_to_latex_2.png)

---

### 📝 Markdown → Markdown

将您的 Markdown 内容保存为 `.md` 文件以供将来使用。

---

## 命令行界面 (CLI) 用法

Markdown Exporter 提供了一个强大的命令行界面，让您可以直接从终端访问其所有功能。

### 安装

```bash
# with pip
pip install md-exporter

# with uv
uv tool install md-exporter
```

### 基本用法

使用 `markdown-exporter` 命令访问所有工具：

```bash
markdown-exporter <subcommand> <args> [options]
```

### 工具使用指南

#### md_to_csv - 将 Markdown 表格转换为 CSV

```bash
markdown-exporter md_to_csv <input> <output> [options]
```

- **参数**: `input` (Markdown 文件路径), `output` (CSV 文件路径)
- **选项**: `--strip-wrapper` (如果存在则移除代码块包装器)

#### md_to_pdf - 将 Markdown 转换为 PDF

```bash
markdown-exporter md_to_pdf <input> <output> [options]
```

- **参数**: `input` (Markdown 文件路径), `output` (PDF 文件路径)
- **选项**: `--strip-wrapper` (如果存在则移除代码块包装器)

#### md_to_docx - 将 Markdown 转换为 DOCX

```bash
markdown-exporter md_to_docx <input> <output> [options]
```

- **参数**: `input` (Markdown 文件路径), `output` (DOCX 文件路径)
- **选项**: `--template` (DOCX 模板文件路径), `--strip-wrapper` (如果存在则移除代码块包装器)

#### md_to_xlsx - 将 Markdown 表格转换为 XLSX

```bash
markdown-exporter md_to_xlsx <input> <output> [options]
```

- **参数**: `input` (Markdown 文件路径), `output` (XLSX 文件路径)
- **选项**: `--force-text` (将单元格值转换为文本类型), `--strip-wrapper` (如果存在则移除代码块包装器)

#### md_to_pptx - 将 Markdown 转换为 PPTX

```bash
markdown-exporter md_to_pptx <input> <output> [options]
```

- **参数**: `input` (Markdown 文件路径), `output` (PPTX 文件路径)
- **选项**: `--template` (PPTX 模板文件路径)

#### md_to_codeblock - 提取代码块到文件

```bash
markdown-exporter md_to_codeblock <input> <output> [options]
```

- **参数**: `input` (Markdown 文件路径), `output` (输出目录或 ZIP 文件路径)
- **选项**: `--compress` (将所有代码块压缩为 ZIP 文件)

#### md_to_json - 将 Markdown 表格转换为 JSON

```bash
markdown-exporter md_to_json <input> <output> [options]
```

- **参数**: `input` (Markdown 文件路径), `output` (JSON 文件路径)
- **选项**: `--style` (JSON 输出样式: jsonl 或 json_array), `--strip-wrapper` (如果存在则移除代码块包装器)

#### md_to_xml - 将 Markdown 转换为 XML

```bash
markdown-exporter md_to_xml <input> <output> [options]
```

- **参数**: `input` (Markdown 文件路径), `output` (XML 文件路径)
- **选项**: `--strip-wrapper` (如果存在则移除代码块包装器)

#### md_to_latex - 将 Markdown 表格转换为 LaTeX

```bash
markdown-exporter md_to_latex <input> <output> [options]
```

- **参数**: `input` (Markdown 文件路径), `output` (LaTeX 文件路径)
- **选项**: `--strip-wrapper` (如果存在则移除代码块包装器)

#### md_to_html - 将 Markdown 转换为 HTML

```bash
markdown-exporter md_to_html <input> <output> [options]
```

- **参数**: `input` (Markdown 文件路径), `output` (HTML 文件路径)
- **选项**: `--strip-wrapper` (如果存在则移除代码块包装器)

#### md_to_html_text - 将 Markdown 转换为 HTML 文本

```bash
markdown-exporter md_to_html_text <input>
```

- **参数**: `input` (Markdown 文件路径)

#### md_to_md - 将 Markdown 转换为 MD 文件

```bash
markdown-exporter md_to_md <input> <output>
```

- **参数**: `input` (Markdown 文件路径), `output` (MD 文件路径)

#### md_to_ipynb - 将 Markdown 转换为 IPYNB

```bash
markdown-exporter md_to_ipynb <input> <output> [options]
```

- **参数**: `input` (Markdown 文件路径), `output` (IPYNB 文件路径)
- **选项**: `--strip-wrapper` (如果存在则移除代码块包装器)

### 重要说明

- 所有命令仅支持文件路径作为输入
- 包会自动处理所有依赖管理
- 您可以在系统的任何位置运行命令，无需导航到项目目录
- 使用 `markdown-exporter <subcommand> -h` 查看每个子命令的详细帮助

---

## 📢 发布版本

发布版本可在以下位置获取：

- [GitHub 仓库发布](https://github.com/bowenliang123/markdown-exporter/releases)
- [Dify 市场发布](https://marketplace.dify.ai/plugins/bowenliang123/md_exporter)

### 更新日志

- 3.6.9

  - 通过删除 README 文档的截图图像，将打包的 Dify 插件文件大小减少了 95%。

- 3.6.7

  - 修复了 Python 分发包中 docx 和 pptx 模板文件的路径错误问题

- 3.6.6

  - 重构代码结构，确保 Agent Skill、Dify 插件和 CLI 正确共享核心文件转换逻辑
  - 使项目可作为 Agent Skill 🦞 在 OpenClaw 上安装和使用
  - 重构 Python 打包，使 `markdown-exporter` 作为独立的 CLI 工具，安装 Python 包 `md-exporter`
  - 移除 `md_to_linked_image` 工具

- 3.6.0

  - 通过移除对标题头中空格和空行的强制要求，提高 `md_to_docx`、`md_to_pptx` 和 `md_to_ipynb` 工具的转换成功率
  - 通过运行预热方法加快 pandoc 的首次调用速度

- 3.5.1

  - 在 `md_to_ipynb` 工具中通过预处理 markdown 输入正确处理代码单元格

- 3.5.0

  - 引入 `md_to_ipynb` 工具，用于将 Markdown 文本转换为 Jupyter Notebook (.ipynb) 格式
  - `md_to_ipynb` 工具自动将 markdown 内容分割为单独的笔记本单元格
  - 更新 logo 图标（首次）

- 3.4.0

  - [重大变更] 将 `md_to_pptx` 工具迁移到使用 `pandoc` 进行转换，以获得稳定的功能和减少依赖占用
  - `md_to_pptx` 工具的 Markdown 输入现在必须遵循 [Pandoc 幻灯片](https://pandoc.org/MANUAL.html#slide-shows) 中的 markdown 语法和指导

- 3.3.0

  - 将 `md_to_pptx` 从 6.1.1 更新到 6.2.1
  - 移除 `md_to_mermaid` 工具，通过消除 Node.js 运行时依赖来减少安装时间和占用空间

- 3.2.0

  - 引入 `md_to_mermaid` 工具，用于将 Markdown 中的 Mermaid 图表代码块转换为 PNG 图像
  - 重构 Agent Skill 脚本和入口 shell 脚本

- 3.0.0

  - `md_exporter` 现已准备好用于 Agent Skills 并可独立运行
    - 添加 `SKILL.md` 用于 Agent Skills 描述
    - 添加 `/scripts` 作为所有工具的代码脚本，作为 Agent Skills 执行入口
    - 将核心文件转换逻辑提取到 `/scripts/lib`，由 Agent Skills 脚本和 Dify 插件共享
    - 添加 `pyproject.toml` 作为 Python 项目描述和依赖管理
  - 添加涵盖 Claude Skills 脚本入口点所有用法的自动化测试
  - 将 `md2pptx` 从 6.1 升级到 6.1.1 以修复表格标题 bug

- 2.3.0

  - 通过在 `md_to_xlsx` 工具中跳过第一个表格之前的无关字符来解决 XLSX 生成问题

- 2.2.0

  - 修复 GitHub Actions 中的 CI 问题

- 2.1.0

  - 通过将 `md2pptx` 更新到 6.1 来修复 `md_to_pptx` 工具中的远程图像获取问题
  - 通过更新默认 DOCX 模板文件来修复 `md_to_docx` 工具中缺失的表格边框
  - 拦截 `md_to_pptx` 工具中的 `run-python` 宏使用

- 2.0.0

  - [重大变更] 将 `md_to_docx` 工具迁移到使用 `pandoc` 进行转换
  - `md_to_docx` 工具的主要改进：
    - 支持使用模板 DOCX 文件自定义样式，允许自定义标题、段落等的字体、字号和颜色样式
    - 优化生成的 DOCX 文件大小
    - 更好地支持多语言内容
  - 移除 `md_to_rst` 和 `md_to_epub` 工具
  - 在 `md_to_xlsx` 工具中添加 `force_text_value` 选项以控制是否强制将所有单元格值转换为文本字符串

- 1.12.0

  - 修复 `md_to_pptx` 工具中的可选 PPTX 模板文件处理
  - 修复 `md_to_pptx` 工具中命令组装中的文件路径错误消息
  - 将 `md2pptx` 升级到 6.0

- 1.10.2

  - 将 `md2pptx` 从 5.4.4 升级到 5.4.5
  - 将 `python-docx` 从 1.1 升级到 1.2.0
  - 将 PDF 生成功能限制提高到 500MB

- 1.10.0

  - 在 `md_to_json` 工具中支持 JSONL 输出样式，使用每行一个对象的 JSON Lines 格式
  - 将 `md_to_json` 工具的默认输出样式更改为 JSONL
  - 参数描述中的小文档更新

- 1.9.0

  - 在 `md_to_xlsx` 工具中使用 Markdown 文本中的标题支持自定义工作表名称
  - 在 `md_to_xlsx` 工具中强制将列类型转换为字符串，以防止 Microsoft Excel 中的数据精度丢失和显示问题
  - 在 `md_to_xlsx` 工具中自动调整列宽

- 1.8.0

  - 解决在 Microsoft Excel 中打开包含非 ASCII 字符（例如中文、日文、表情符号字符）的 CSV 文件时的乱码问题

- 1.7.0

  - 在 `md_to_pptx` 工具中支持自定义 PPTX 模板文件
  - 在 `md_to_csv`、`md_to_latex` 和 `md_to_xlsx` 工具中支持从多个表格生成文件

- 1.6.0

  - 引入 `md_to_html_text` 工具，用于将 Markdown 文本转换为 HTML 文本
  - 标准化 `md_to_docx` 工具生成的 DOCX 文件中标题和正文段落的字体

- 1.5.0

  - 通过跳过 CJK 字符的字体设置，改善 `md_to_pdf` 工具中纯英文 markdown 文本输入的 PDF 显示效果
  - 引入 `md_to_epub` 工具，用于将 Markdown 文本转换为 EPUB 电子书文件
  - 在 `md_to_png` 工具中支持将所有 PNG 文件压缩为单个 zip 文件
  - 在 `md_to_pdf` 工具中将 PDF 文件的容量限制提高到 100MB
  - 移除显式的超时配置 MAX_REQUEST_TIMEOUT

- 1.4.100

  - 庆祝 [Dify](https://github.com/langgenius/dify) 达到 10 万 GitHub stars 里程碑的特别版本
  - 添加 `md_to_png` 工具，用于将 Markdown 文本转换为 PNG 图像文件

- 1.3.0

  - 更新 SDK 版本

- 1.2.0

  - 在 `md_to_linked_image` 工具中支持将图像压缩为单个 zip 文件

- 1.1.0

  - 默认在所有工具中启用换行符规范化，将所有 `\n` 替换为 `\n`
  - 移除输入 Markdown 文本中的推理内容的 `<think>` 标签
  - 修复 `md_to_csv`、`md_to_json` 和 `md_to_latex` 工具中缺少的自定义输出文件名支持

- 1.0.1

  - 移除自定义输出文件名中冗余的 URL 安全转换

- 1.0.0

  - 支持自定义输出文件名

- 0.5.0

  - 引入 `md_to_linked_image` 工具，用于从 Markdown 文本中的链接提取图像文件

- 0.4.3

  - 通过在全球字体中包含中文字符时设置为宋体，改善 `md_to_docx` 工具中文本段落的中文字符显示

- 0.4.2

  - 在 `md_to_codeblock` 工具中支持 Java、PHP 和 Ruby 文件导出

- 0.4.1

  - 在 `md_to_codeblock` 工具中支持 YAML 文件导出

- 0.4.0

  - 在 `md_to_codeblock` 工具中支持将 Markdown 代码块导出为单个 zip 文件

- 0.3.0

  - 修复在自托管 Dify 插件守护服务上运行时 `md_to_pptx` 工具中的库导入错误

- 0.2.0

  - 引入 `md_to_codeblock` 工具，用于将 Markdown 中的代码块提取为 Python、JSON、JS、BASH、SVG、HTML、XML 和 MARKDOWN 文件
  - 引入 `md_to_rst` 工具，用于将 Markdown 转换为 reStructuredText (.rst) 格式

- 0.1.x

  - 引入 `md_to_pptx` 工具，用于将 Markdown 转换为 PowerPoint (.pptx) 格式

- 0.0.x
  - 发布到 Dify 市场
  - 支持将 Markdown 导出为 DOCX、PPTX、XLSX、PDF、HTML、MD、CSV、JSON、XML 和 LaTeX 文件

---

## 🤝 贡献

欢迎贡献！请随时在我们的 [GitHub 仓库](https://github.com/bowenliang123/markdown-exporter) 提交问题或拉取请求。

### 代码风格

我们使用 `ruff` 来确保代码一致性。运行以下命令自动修复代码风格问题：

```bash
uv run ruff check --fix --select UP .
```

或使用提供的脚本：

```bash
dev/reformat.sh
```

---

## 📜 许可证

本项目采用 **Apache License 2.0** 许可证。

---

## 🔒 隐私

隐私很重要。有关更多详细信息，请参阅 [隐私政策](./PRIVACY.md)。本插件：

- **不收集**任何数据
- **本地**处理所有内容
- **不发送**任何信息给第三方服务

所有文件转换完全在本地环境中进行。

---

## 🙏 致谢

本项目站在巨人的肩膀上。我们感谢这些优秀的开源项目：

| Project                                               | License              |
| ----------------------------------------------------- | -------------------- |
| [pypandoc](https://github.com/JessicaTegner/pypandoc) | MIT License          |
| [pandas](https://github.com/pandas-dev/pandas)        | BSD 3-Clause License |
| [xhtml2pdf](https://github.com/xhtml2pdf/xhtml2pdf)   | Apache License 2.0   |
