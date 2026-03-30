# MD2DOCX - Markdown 转 Word 桌面工具

![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)
![PyQt6](https://img.shields.io/badge/GUI-PyQt6-green.svg)
![python-docx](https://img.shields.io/badge/Word-python--docx-orange.svg)

**MD2DOCX** 是一个基于 Python 和 PyQt6 开发的轻量级桌面应用程序。它可以方便地将 `.md` (Markdown) 文件快速转换为带有标准 Word 样式的 `.docx` 文档，支持一键拖拽和两种不同的标题级别映射。

---

## ✨ 核心特性

- **图形化界面 (GUI)**：基于 PyQt6 的现代化界面，支持拖拽文件或点击选择文件。
- **丰富的 Markdown 元素支持**：依靠 `mistune` 强大的 AST 解析能力，精准转换以下元素：
  - 标题 (H1~H6)
  - 文本格式（**加粗**、*斜体*）
  - 列表（无序列表、有序列表及嵌套）
  - 代码块（带底色和边框格式化的等宽字体）和行内代码
  - 表格（自动加粗表头）
  - 引用说明块（左侧带有提示颜色边条及底色）
  - 分隔线
  - 超链接（蓝色下划线）
  - 本地图片嵌入
- **两种不同的标题排版样式**：
  - **样式 A**：文档顶部的 `# 标题` 被识别为 Word 内置的 **「标题 (Title)」** 样式，其余标题往下顺延（`##` -> `Heading 1`）。
  - **样式 B**：标准的映射方式，`# 标题` 被识别为 **「标题 1 (Heading 1)」** 样式（`#` -> `Heading 1`, `##` -> `Heading 2`）。
- **一键打包部署**：自带打包脚本，无需配置即可生成独立的免安装 `.exe`。

---

## 📂 项目结构

```text
md2docx/
├── main.py                  # 程序的启动入口
├── build.py                 # 一键打包 .exe 脚本
├── requirements.txt         # 所有的 Python 依赖项
├── test_sample.md           # 包含各种 Markdown 语法的综合测试示例文档
├── test_convert.py          # 无 GUI 的纯命令行转换测试脚本
├── converter/               # 转换引擎核心模块
│   ├── md_parser.py         # Markdown AST 语法树解析器
│   └── docx_writer.py       # DOCX 文档生成及格式映射引擎
├── styles/
│   └── docx_styles.py       # 全局样式配置（字体、字号配置与标题映射规则）
└── ui/
    └── main_window.py       # PyQt6 主窗口视图与交互逻辑
```

---

## 🛠️ 安装与运行指南

### 1. 环境依赖

确保您的系统已安装了 Python（建议 3.9 或更高版本）。

克隆或下载本仓库后，在命令行中执行以下命令安装依赖：

```bash
pip install -r requirements.txt
```

*主要依赖项：`PyQt6`, `python-docx`, `mistune`*

### 2. 通过源代码启动应用程序

安装完依赖后，在项目根目录下直接运行：

```bash
python main.py
```

### 3. 如何使用

1. 打开应用程序。
2. 将你需要转换的 Markdown (*.md*) 文件直接拖入软件中央的虚线框内，或者点击下方按钮选择文件。
3. 根据您的排版需求，在底部的单选框里选择转换样式（样式 A 或 样式 B）。
4. 点击 **🚀 开始转换** 按钮，选择你想保存生成的 Word 文档的位置。
5. 等待状态栏提示“转换成功”，可以在目标文件夹中查收生成的文档了！

---

## 📦 打包为独立的可执行文件 (EXE)

如果您希望将本工具分享给没有安装 Python 环境的人使用，可以使用项目中自带的构建脚本打包出无弹窗的单文件 Windows `exe`。

在项目根目录下运行：

```bash
python build.py
```

程序将自动：

1. 检测并安装 `pyinstaller`（如果没有安装的话）
2. 清除之前的历史打包缓存 (`build/`, `dist/`)
3. 读取项目源代码进行依赖图分析，并将所有依赖项压入一个文件
4. 执行完毕后，你将在项目目录下的 **`dist/MD2DOCX.exe`** 拿到最终的免安装程序。

---

## 🧩 二次开发说明

- **字体与样式**：如果想修改生成的 Word 默认中文字体、文本字号或段落行高，请直接修改 `styles/docx_styles.py` 中的 `FONT_CONFIG` 变量。
- **添加新 Markdown 语法**：
  1. 需要确保在 `converter/md_parser.py` 的 Mistune 选项中开启了相关的 `plugins`。
  2. 随后在 `converter/docx_writer.py` 的 `_process_node()` 函数增加对应新类型 AST 节点的拦截和生成逻辑。
