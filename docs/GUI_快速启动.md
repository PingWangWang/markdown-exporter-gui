# Markdown Exporter GUI 快速启动指南

## 🚀 快速开始（3 步）

### 步骤 1：确认 Python 版本

```bash
python --version
# 或
py --version
```

需要 Python 3.11 或更高版本。

### 步骤 2：安装依赖

```bash
pip install -r requirements.txt
```

或使用 uv（更快）：

```bash
uv sync
```

**Linux 用户额外步骤：**

```bash
sudo apt-get install python3-tk
```

### 步骤 3：启动 GUI

```bash
python run_gui.py
```

## 📖 使用流程

1. **选择文件** → 点击"选择文件"按钮，选择一个或多个 .md 文件
2. **选择位置** → 点击"保存位置"按钮，选择输出目录
3. **选择格式** → 从下拉菜单选择目标格式（DOCX、PDF、HTML 等）
4. **开始转换** → 点击"▶ 开始转换"按钮
5. **查看结果** → 点击"📂 打开输出目录"查看生成的文件

## 💡 提示

- ✅ 支持批量转换多个文件
- ✅ 转换在后台进行，界面不会卡死
- ✅ 日志区域实时显示转换进度
- ✅ 文件已存在时会询问是否覆盖

## ❓ 常见问题

### Q: 提示找不到 tkinter？

**A:** Linux 用户需要安装：`sudo apt-get install python3-tk`

### Q: 提示找不到 md_exporter 模块？

**A:** 确保已安装依赖：`pip install -r requirements.txt`

### Q: 如何转换为特定格式？

**A:** 在"选择输出格式"下拉框中选择即可

### Q: 支持自定义模板吗？

**A:** 当前 GUI 版本暂不支持，请使用命令行版本

## 📚 更多信息

详细文档请查看：[gui/README.md](gui/README.md)

项目主页：[GitHub - markdown-exporter](https://github.com/bowenliang123/markdown-exporter)
