# MD2DOCX

> 🖋️ 轻量级 Markdown → Word 桌面转换工具，面向中文公文排版场景

![Python](https://img.shields.io/badge/Python-3.9+-3776AB?logo=python&logoColor=white)
![PyQt6](https://img.shields.io/badge/GUI-PyQt6-41CD52?logo=qt&logoColor=white)
![python-docx](https://img.shields.io/badge/Word-python--docx-2B579A?logo=microsoftword&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS-lightgrey)

**MD2DOCX** 是一款基于 Python + PyQt6 的跨平台桌面应用。将 `.md` 文件一键转化为带标准 Word 样式的 `.docx` 文档——仿宋正文、黑体标题、首行缩进、页脚页码，开箱即用，无需手动排版。

---

## ✨ 核心特性

### 📝 Markdown 解析

基于 [Mistune](https://github.com/lepture/mistune) AST 引擎，精准转换以下元素：

| 元素 | 说明 |
|---|---|
| 标题 | H1 ~ H6，支持两套 Word 大纲映射 |
| 文本格式 | **加粗**、*斜体* |
| 列表 | 有序 / 无序，多层嵌套 |
| 代码 | 代码块（底色 + 边框 + 等宽字体）、行内代码 |
| 表格 | 自动加粗表头 |
| 引用块 | 左侧蓝色边条 + 淡蓝底色 |
| 超链接 | 蓝色下划线 |
| 图片 | 本地图片嵌入 |
| 分隔线 | 水平分割 |

### 📄 双输入模式

- **文件转换** — 拖拽或选择 `.md` 文件
- **文本转换** — 直接粘贴 / 手敲 Markdown 原文，即时转换

### 🎯 智能排版容错

即使用户粘贴的文本缺少严格空行，也能正确排版：

- 标题与正文间缺少空行 → 自动隔离
- 段落内软回车（`^l`） → 自动切分为独立 Word 段落（`^p`）

### 🎨 专业的排版样式

所有排版参数统一管理在 `styles/docx_styles.py`，可视化调整：

| 样式 | 字体 | 字号 | 行距 | 对齐 | 段前 / 段后 |
|---|---|---|---|---|---|
| 正文 (Normal) | 仿宋_GB2312 | 小四 (12pt) | 单倍 | 两端对齐 | 0 / 0 |
| 正文文本 (Body Text) | 仿宋_GB2312 | 小四 (12pt) | 1.2 倍 | 两端对齐 | 0.5 行 / 0 |
| 一级标题 (`#`) | 黑体 + 加粗 | 18pt | 单倍 | 居中 | 0 / 16 磅 |
| 其余标题 (`##` ~ `######`) | 宋体 | 按级递减 | 单倍 | 左对齐 | 1 行 / 0.5 行 |
| 代码块 | Consolas | 10pt | — | — | 6pt / 6pt |

> 首行缩进 2 字符 (24pt)，文档网格对齐已全局禁用。

### 🔀 两种标题级别映射

| Markdown | 样式 A（Title 模式） | 样式 B（Heading 模式） |
|---|---|---|
| `#` | Title (18pt) | Heading 1 (18pt) |
| `##` | Heading 1 (16pt) | Heading 2 (16pt) |
| `###` | Heading 2 (14pt) | Heading 3 (14pt) |
| `####` | Heading 3 (12pt) | Heading 4 (12pt) |
| `#####` | Heading 4 (11pt) | Heading 5 (11pt) |
| `######` | Heading 5 (10.5pt) | Heading 6 (10.5pt) |

> 两种模式下同级标题视觉大小完全一致，区别仅在 Word 大纲层级。

### 🔧 自动化能力

- **字体自动安装** — 首次启动自动检测并静默安装仿宋_GB2312 / 宋体 / 黑体（Windows 写注册表，macOS 写 `~/Library/Fonts`，无需管理员权限）
- **页脚页码** — 自动生成「第 X 页 共 Y 页」居中页码
- **智能保存** — 自动预填 `.docx` 文件名，弹出保存对话框自由选择归档位置
- **快速预览** — 转换成功后一键调用系统 Office 打开生成的文档
- **文件占用友好提示** — 检测到文件被 Word 占用时弹出提示而非崩溃

---

## 🚀 快速开始

### 方式一：极速脚本启动（推荐）

已安装 Python 环境的用户，无需打包即可使用：

**💻 Windows**
```
双击 run-win.bat
```

**🍎 macOS**
```bash
chmod +x run-mac.command   # 首次运行前赋权
双击 run-mac.command
```

> 脚本会自动通过 [uv](https://github.com/astral-sh/uv) 同步依赖并启动 GUI。

### 方式二：手动运行

```bash
# 1. 安装依赖
uv sync

# 2. 启动应用
uv run python main.py
```

### 方式三：打包为独立可执行文件

```bash
python build.py
```

打包脚本会自动：

1. 检测并安装 `pyinstaller`
2. 清除旧的打包缓存 (`build/`, `dist/`)
3. 将字体资源及所有依赖打包为单文件
4. 根据当前操作系统生成对应产物：
   - 💻 Windows → `dist/MD2DOCX.exe`
   - 🍎 macOS → `dist/MD2DOCX.app`

> ⚠️ 由于二进制不可跨平台，需在目标系统上分别执行 `python build.py` 打包。

---

## 📖 使用方法

**文件转换**

1. 停留在「📄 文件转换」标签页
2. 将 `.md` 文件拖入中央区域，或点击按钮选择文件

**文本转换**

1. 切换到「✍️ 文本转换」标签页
2. 在文本框内粘贴 Markdown 内容

**共通步骤**

1. 在底部选择转换样式（样式 A 或样式 B）
2. 点击 **🚀 开始转换**
3. 在弹出的对话框中选择保存位置
4. 转换完成后，可点击 **打开文档** 按钮直接预览结果

---

## 📂 项目结构

```
md2docx/
├── main.py                  # 应用入口
├── pyproject.toml           # 项目依赖配置 (uv)
├── build.py                 # 跨平台打包脚本
├── run-win.bat              # Windows 一键启动
├── run-mac.command           # macOS 一键启动
│
├── ui/
│   └── main_window.py       # PyQt6 主窗口 & 交互逻辑
│
└── .agents/skills/md2docx-converter/  # 核心转换逻辑 (Skill)
    ├── converter/           # 转换引擎 (AST解析 & DOCX生成)
    ├── styles/              # 全局样式配置
    ├── utils/               # 字体自动检测与安装
    ├── font/                # 内嵌字体资源
    └── scripts/convert.py   # CLI 转换入口
```

---

## 🧩 二次开发

### 样式与排版

所有样式参数集中管理在 `styles/docx_styles.py`：

| 配置项 | 说明 |
|---|---|
| `FONT_CONFIG` | 正文 / 标题 / 一级标题 / 代码的字体配置 |
| `NORMAL_*` | 基础正文 (Normal) 的行距、段前段后 |
| `BODY_*` | 正文文本 (Body Text) 的行距、缩进、段前段后 |
| `HEADING_*` | 标题样式的行距、段前段后、居中等级 |
| `HEADING_FONT_SIZES_A/B` | 两种模式的标题字号映射表 |
| `OL_*` / `UL_*` | 有序列表 / 无序列表的缩进与间距 |
| `FOOTER_*` | 页脚页码的字体和字号 |
| `CODE_BLOCK_*` | 代码块的背景色和边框色 |

> ⚠️ **重要**：`docx_styles.py` 是样式的唯一真实来源。`docx_writer.py` 的渲染逻辑必须始终引用其导出变量，**不允许硬编码格式参数**。

### 添加新 Markdown 语法

1. 在 `converter/md_parser.py` 的 Mistune 配置中启用对应的 `plugins`
2. 在 `converter/docx_writer.py` 的 `_process_node()` 中拦截新类型 AST 节点并实现生成逻辑

---

## 📋 依赖项

| 包 | 用途 |
|---|---|
| [PyQt6](https://pypi.org/project/PyQt6/) | GUI 框架 |
| [python-docx](https://pypi.org/project/python-docx/) ≥ 1.2.0 | Word 文档生成 |
| [mistune](https://pypi.org/project/mistune/) ≥ 3.0.0 | Markdown AST 解析 |
| [pyinstaller](https://pypi.org/project/pyinstaller/) | 打包为可执行文件（可选） |

---

## 📄 License

MIT
