<p align="center">
  <h1 align="center">📝 MD2DOCX</h1>
  <p align="center">
    <strong>轻量级 Markdown → Word 桌面转换工具</strong><br/>
    面向中文公文排版场景，开箱即用
  </p>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.9+-3776AB?logo=python&logoColor=white" alt="Python"/>
  <img src="https://img.shields.io/badge/GUI-PyQt6-41CD52?logo=qt&logoColor=white" alt="PyQt6"/>
  <img src="https://img.shields.io/badge/Word-python--docx-2B579A?logo=microsoftword&logoColor=white" alt="python-docx"/>
  <img src="https://img.shields.io/badge/Platform-Windows%20%7C%20macOS-lightgrey" alt="Platform"/>
  <img src="https://img.shields.io/badge/Version-1.0.0-blue" alt="Version"/>
  <img src="https://img.shields.io/badge/License-MIT-green" alt="License"/>
</p>

---

**MD2DOCX** 是一款基于 Python + PyQt6 的**跨平台桌面应用**。将 Markdown 文件或文本一键转化为带标准 Word 样式的 `.docx` 文档——仿宋正文、黑体标题、首行缩进、页脚页码，开箱即用，无需手动排版。

同时提供 **CLI 命令行接口**，可作为 AI Agent Skill 被自动化工具链直接调用。

---

## 📑 目录

- [核心特性](#-核心特性)
- [快速开始](#-快速开始)
- [使用方法](#-使用方法)
- [项目架构](#-项目架构)
- [样式与排版系统](#-样式与排版系统)
- [CLI 命令行接口](#-cli-命令行接口)
- [二次开发指南](#-二次开发指南)
- [依赖项](#-依赖项)
- [常见问题](#-常见问题)
- [License](#-license)

---

## ✨ 核心特性

### 📝 Markdown 元素全覆盖

基于 [Mistune v3](https://github.com/lepture/mistune) AST 引擎，精准解析并转换以下 Markdown 元素：

| 元素     | 说明                         | 排版效果                              |
| -------- | ---------------------------- | ------------------------------------- |
| 标题     | `#` ~ `######`（H1–H6） | 支持两套 Word 大纲映射，字号自动分级  |
| 文本格式 | **加粗**、*斜体*     | 原生 Word Run 格式化                  |
| 有序列表 | `1. 2. 3.`                 | `List Paragraph` 样式，支持多层嵌套 |
| 无序列表 | `- * +`                    | `List Bullet` 样式，支持多层嵌套    |
| 代码块   | ` ```lang ``` `            | 灰色底色 + 边框 + Consolas 等宽字体   |
| 行内代码 | `` `code` ``                 | 等宽字体内嵌正文                      |
| 表格     | `                            | col1                                  |
| 引用块   | `> quote`                  | 左侧蓝色边条 + 淡蓝底色               |
| 超链接   | `[text](url)`              | 蓝色下划线                            |
| 图片     | `![alt](path)`             | 本地图片嵌入                          |
| 分隔线   | `---` / `***`            | 水平分割                              |

### 📄 双输入模式

| 模式                   | 说明                                             |
| ---------------------- | ------------------------------------------------ |
| 📄**文件转换**   | 拖拽或点击选择 `.md` 文件，支持一键转换        |
| ✍️**文本转换** | 直接在编辑框中粘贴或手敲 Markdown 原文，即时转换 |

### 🎯 智能排版容错

即使用户粘贴的文本格式不规范，也能正确排版：

- **标题隔离** — 标题与正文间缺少空行时，自动补充空行后解析
- **段落切分** — 段落内的软回车（`^l` / Shift+Enter）自动切分为独立的 Word 段落（`^p`）
- **紧凑列表** — 无空行的列表项（Tight List）中的 `block_text` 节点正确处理，不丢失内容

### 🔀 两种标题级别映射

提供两种 Word 大纲映射模式，适配不同应用场景：

| Markdown   | 样式 A（Title 模式）     | 样式 B（Heading 模式）       |
| ---------- | ------------------------ | ---------------------------- |
| `#`      | Title (18pt, 黑体, 居中) | Heading 1 (18pt, 黑体, 居中) |
| `##`     | Heading 1 (16pt, 宋体)   | Heading 2 (16pt, 宋体)       |
| `###`    | Heading 2 (14pt, 宋体)   | Heading 3 (14pt, 宋体)       |
| `####`   | Heading 3 (12pt, 宋体)   | Heading 4 (12pt, 宋体)       |
| `#####`  | Heading 4 (11pt, 宋体)   | Heading 5 (11pt, 宋体)       |
| `######` | Heading 5 (10.5pt, 宋体) | Heading 6 (10.5pt, 宋体)     |

> **说明**：两种模式下同级标题的视觉大小完全一致，区别仅在于 Word 大纲层级。
>
> - **样式 A** 适合：`#` 作为文档总标题（Title），`##` 起为正文章节
> - **样式 B** 适合：`#` 直接作为一级章节标题（Heading 1）

### 🔧 自动化能力

| 功能                       | 说明                                                                            |
| -------------------------- | ------------------------------------------------------------------------------- |
| 🔤**字体自动安装**   | 首次启动自动检测并静默安装所需字体（仿宋_GB2312 / 宋体 / 黑体），无需管理员权限 |
| 📄**页脚页码**       | 自动生成「第 X 页 共 Y 页」居中页码（宋体 9pt）                                 |
| 💾**智能保存**       | 自动预填 `.docx` 文件名，弹出保存对话框自由选择归档位置                       |
| 👁️**快速预览**     | 转换成功后一键调用系统默认程序打开文档                                          |
| ⚠️**占用友好提示** | 检测到文件被 Word 占用时弹出提示，而非直接崩溃                                  |

### 🌍 跨平台支持

| 平台                | 字体安装                   | 文件预览             | 打包产物 | 快速启动                       |
| ------------------- | -------------------------- | -------------------- | -------- | ------------------------------ |
| 💻**Windows** | 写入用户注册表             | `os.startfile`     | `.exe` | `启动 MD2DOCX.vbs`（无黑窗） |
| 🍎**macOS**   | 复制至 `~/Library/Fonts` | `subprocess: open` | `.app` | `run-mac.command`            |

---

## 🚀 快速开始

### 前置要求

- **Python** ≥ 3.9
- **[uv](https://github.com/astral-sh/uv)** — 推荐的 Python 依赖管理器（启动脚本会自动检测）

### 方式一：极速脚本启动（推荐 ⭐）

无需打包，双击即可运行，脚本会自动通过 `uv` 同步依赖并启动 GUI。

**💻 Windows**

```
双击「启动 MD2DOCX.vbs」
```

> 💡 使用 VBS 启动器可**完全隐藏命令行黑窗口**，直接弹出 GUI 界面。
> 如需查看运行日志或调试，可改用 `run-win.bat`。

**🍎 macOS**

```bash
# 首次运行需赋予执行权限
chmod +x run-mac.command

# 之后双击 run-mac.command 即可
# macOS 版会自动安装 uv（如果未安装）
```

### 方式二：手动运行

```bash
# 1. 安装依赖
uv sync

# 2. 启动应用
uv run python main.py
```

### 方式三：打包为独立可执行文件

```bash
uv run python build.py
```

打包脚本会自动：

1. 检测并安装 `pyinstaller`（如未安装）
2. 清除旧的打包缓存（`build/`、`dist/`）
3. 将字体资源及所有依赖打包为单文件
4. 根据当前操作系统生成对应产物：
   - 💻 Windows → `dist/MD2DOCX.exe`
   - 🍎 macOS → `dist/MD2DOCX.app`

> ⚠️ 由于二进制不可跨平台，需在目标系统上分别执行 `python build.py` 打包。

---

## 📖 使用方法

### 📄 文件转换

1. 停留在「📄 文件转换」标签页
2. 将 `.md` 文件**拖入**中央区域，或点击 **选择文件** 按钮
3. 在底部选择转换样式（样式 A 或样式 B）
4. 点击 **🚀 开始转换**
5. 在弹出的对话框中选择保存位置
6. 转换完成后，可点击 **打开文档** 按钮直接预览结果

### ✍️ 文本转换

1. 切换到「✍️ 文本转换」标签页
2. 在文本框内粘贴或输入 Markdown 内容
3. 在底部选择转换样式（样式 A 或样式 B）
4. 点击 **🚀 开始转换**
5. 在弹出的对话框中选择保存位置
6. 转换完成后，可点击 **打开文档** 按钮直接预览结果

> 💡 **提示**：点击 **清除选择** 按钮可重置已选文件或输入内容。

---

## 📂 项目架构

```
md2docx/
├── main.py                          # 🚀 应用入口（字体检查 + GUI 启动）
├── pyproject.toml                   # 📦 项目依赖配置 (uv)
├── build.py                         # 🔨 跨平台打包脚本 (PyInstaller)
├── 启动 MD2DOCX.vbs                 # 💻 Windows 无黑窗启动器（推荐）
├── run-win.bat                      # 💻 Windows 一键启动（调试用）
├── run-mac.command                  # 🍎 macOS 一键启动
│
├── ui/
│   ├── __init__.py
│   └── main_window.py               # 🖥️ PyQt6 主窗口（拖拽/选择/文本输入/转换/预览）
│
├── .agents/skills/md2docx-converter/ # 🧠 核心转换引擎（独立可移植 Skill）
│   ├── SKILL.md                     # Skill 使用说明（供 AI Agent 调用）
│   ├── converter/
│   │   ├── md_parser.py             # Markdown → AST 解析器（Mistune 封装）
│   │   └── docx_writer.py           # AST → DOCX 渲染引擎
│   ├── styles/
│   │   └── docx_styles.py           # 📐 全局样式配置（字体/字号/间距/映射）
│   ├── utils/
│   │   └── font_installer.py        # 🔤 字体自动检测与安装
│   ├── font/                        # 📁 内嵌字体资源
│   │   ├── 仿宋_GB2312.ttf
│   │   ├── simsun.ttc               # 宋体
│   │   └── simhei.ttf               # 黑体
│   ├── scripts/
│   │   └── convert.py               # ⌨️ CLI 转换入口
│   ├── examples/
│   │   └── sample.md                # 📝 测试样例
│   └── requirements.txt             # Skill 独立依赖清单
│
├── test_convert.py                  # 🧪 转换测试脚本
├── test_sample.md                   # 🧪 测试用 Markdown 样例
├── DEVELOPMENT_LOG.md               # 📋 开发日志
└── uv.lock                          # 🔒 依赖锁定文件
```

### 架构设计说明

项目采用**核心引擎外置 Skill 架构**：

```
┌─────────────────────────────────┐
│        main.py (入口)           │
│   sys.path.insert(SKILL_DIR)    │
├────────────┬────────────────────┤
│  ui/       │  .agents/skills/   │
│ (GUI 层)   │  md2docx-converter │
│            │  (核心转换引擎)     │
│ PyQt6      │  ├ md_parser.py    │
│ 主窗口     │  ├ docx_writer.py  │
│ 拖拽/选择  │  ├ docx_styles.py  │
│ 文本输入   │  └ font_installer  │
└────────────┴────────────────────┘
```

- **GUI 层**（`ui/`）：负责用户交互（文件拖拽、文本输入、样式选择、转换触发）
- **核心引擎**（`.agents/skills/md2docx-converter/`）：Markdown 解析 → AST → DOCX 生成的完整管线，独立可移植
- **桥接方式**：`main.py` 通过 `sys.path.insert` 将 Skill 目录注入模块搜索路径

> 💡 这种架构使得核心转换逻辑可以直接拷贝到其他项目复用，同时也支持 AI Agent 通过 CLI 脚本直接调用。

---

## 🎨 样式与排版系统

所有排版参数统一管理在 `styles/docx_styles.py`，是样式的**唯一真实来源**。

### 字体配置

| 用途             | 中文字体    | 英文字体    | 字号        |
| ---------------- | ----------- | ----------- | ----------- |
| 正文             | 仿宋_GB2312 | 仿宋_GB2312 | 小四 (12pt) |
| 标题 (##~######) | 宋体        | 宋体        | 按级递减    |
| 一级标题 (#)     | 黑体 + 加粗 | 黑体 + 加粗 | 18pt        |
| 代码块           | Consolas    | Consolas    | 10pt        |
| 页脚             | 宋体        | 宋体        | 9pt         |

### 段落排版

| 样式                       | 行距   | 段前   | 段后   | 对齐            | 首行缩进      |
| -------------------------- | ------ | ------ | ------ | --------------- | ------------- |
| Normal（基础正文）         | 单倍   | 0      | 0      | 两端对齐        | —            |
| Body Text（正文文本）      | 1.2 倍 | 0.5 行 | 0      | 两端对齐        | 2 字符 (24pt) |
| Heading（标题）            | 单倍   | 1 行   | 0.5 行 | 左对齐 / 居中¹ | —            |
| List Paragraph（有序列表） | 1.1 倍 | 0.5 行 | 0      | 两端对齐        | 悬挂 0.7cm    |
| List Bullet（无序列表）    | 1.1 倍 | 0.5 行 | 0      | 两端对齐        | 悬挂 0.41cm   |
| 代码块                     | —     | 6pt    | 6pt    | —              | —            |

> ¹ 一级标题 (`#`) 居中对齐，其余标题左对齐。

### 全局优化

- ✅ 文档网格对齐已全局禁用（消除行间距被网格吸附控制的视觉差）
- ✅ 右缩进自动调整已全局禁用
- ✅ 代码块底色 `#F5F5F5` + 边框色 `#DDDDDD`

---

## ⌨️ CLI 命令行接口

除 GUI 外，项目还提供独立的命令行转换工具，不依赖 PyQt6：

```bash
# 文件模式
python .agents/skills/md2docx-converter/scripts/convert.py \
  --input input.md --output output.docx --style A

# 文本模式
python .agents/skills/md2docx-converter/scripts/convert.py \
  --text "# 标题\n正文内容" --output output.docx

# 标准输入模式
echo "# 标题" | python .agents/skills/md2docx-converter/scripts/convert.py \
  --stdin --output output.docx
```

### CLI 参数

| 参数                  | 缩写   | 必填   | 说明                            |
| --------------------- | ------ | ------ | ------------------------------- |
| `--input`           | `-i` | 三选一 | 输入的 `.md` 文件路径         |
| `--text`            | `-t` | 三选一 | 直接传入 Markdown 文本内容      |
| `--stdin`           | —     | 三选一 | 从标准输入读取 Markdown         |
| `--output`          | `-o` | ✅     | 输出的 `.docx` 文件路径       |
| `--style`           | `-s` | ❌     | 样式模式：`A`（默认）或 `B` |
| `--skip-font-check` | —     | ❌     | 跳过字体检测与自动安装          |

### 退出码

| 退出码 | 说明                               |
| ------ | ---------------------------------- |
| `0`  | 转换成功                           |
| `1`  | 一般性错误（输入无效、转换失败等） |
| `2`  | 文件被占用（Permission denied）    |

---

## 🧩 二次开发指南

### 修改排版样式

所有样式参数集中管理在 `styles/docx_styles.py`：

| 配置项                        | 说明                                          |
| ----------------------------- | --------------------------------------------- |
| `FONT_CONFIG`               | 正文 / 标题 / 一级标题 / 代码的字体名称和字号 |
| `NORMAL_*`                  | 基础正文 (Normal) 的行距、段前、段后          |
| `BODY_*`                    | 正文文本 (Body Text) 的行距、缩进、段前、段后 |
| `HEADING_*`                 | 标题样式的行距、段前、段后、居中等级          |
| `HEADING_FONT_SIZES_A / _B` | 两种模式的标题字号映射表                      |
| `STYLE_A_HEADING_MAP / _B`  | 两种模式的标题级别 → Word 样式名映射         |
| `OL_*`                      | 有序列表缩进与间距参数                        |
| `UL_*`                      | 无序列表缩进与间距参数                        |
| `FOOTER_*`                  | 页脚页码的字体和字号                          |
| `CODE_BLOCK_*`              | 代码块的背景色和边框色                        |

> [!IMPORTANT]
> `docx_styles.py` 是样式的**唯一真实来源**。`docx_writer.py` 的渲染逻辑必须始终引用其导出变量，**不允许硬编码格式参数**。修改样式时应先修改 `docx_styles.py`，再确保 `docx_writer.py` 正确引用。

### 添加新 Markdown 语法

1. 在 `converter/md_parser.py` 的 `mistune.create_markdown()` 调用中，通过 `plugins` 参数启用对应的 Mistune 插件
2. 在 `converter/docx_writer.py` 的 `_process_node()` 方法中，拦截新类型的 AST 节点并实现 DOCX 生成逻辑

### 添加新字体

1. 将字体文件放入 `.agents/skills/md2docx-converter/font/` 目录
2. 在 `utils/font_installer.py` 的 `FONT_LIST` 中添加字体信息：
   ```python
   FONT_LIST = [
       ("字体文件名.ttf", "注册表显示名 (TrueType)"),
       # ...
   ]
   ```
3. 在 `docx_styles.py` 中引用新字体

### 转换管线流程

```
Markdown 文本
    │
    ▼
┌────────────────────────────┐
│  md_parser.parse_markdown  │  正则预处理 → Mistune AST 解析
└────────────┬───────────────┘
             │ AST 节点列表
             ▼
┌────────────────────────────┐
│  DocxWriter._process_node  │  遍历 AST，递归渲染
│  ├─ _add_heading           │  标题
│  ├─ _add_paragraph         │  段落（含软回车切分）
│  ├─ _add_list              │  有序 / 无序列表
│  ├─ _add_code_block        │  代码块
│  ├─ _add_table             │  表格
│  ├─ _add_block_quote       │  引用块
│  ├─ _add_horizontal_rule   │  分隔线
│  └─ _add_inline_content    │  行内元素（加粗/斜体/代码/链接/图片）
└────────────┬───────────────┘
             │
             ▼
┌────────────────────────────┐
│  DocxWriter._setup_footer  │  页脚页码
│  DocxWriter.convert()      │  保存 .docx
└────────────────────────────┘
```

---

## 📋 依赖项

### 运行时依赖

| 包                                                | 版本要求 | 用途                       |
| ------------------------------------------------- | -------- | -------------------------- |
| [PyQt6](https://pypi.org/project/PyQt6/)             | —       | GUI 界面框架               |
| [python-docx](https://pypi.org/project/python-docx/) | ≥ 1.2.0 | Word 文档（`.docx`）生成 |
| [mistune](https://pypi.org/project/mistune/)         | ≥ 3.0.0 | Markdown AST 解析引擎      |

### 开发 / 可选依赖

| 包                                                | 用途                                    |
| ------------------------------------------------- | --------------------------------------- |
| [pyinstaller](https://pypi.org/project/pyinstaller/) | 打包为可执行文件（`.exe` / `.app`） |
| [uv](https://github.com/astral-sh/uv)                | 推荐的 Python 依赖管理器                |

---

## ❓ 常见问题

<details>
<summary><strong>Q: 首次运行提示缺少字体怎么办？</strong></summary>

程序会在首次启动时自动检测并安装所需字体（仿宋_GB2312、宋体、黑体）。安装至当前用户的字体目录，无需管理员权限。如果自动安装失败，控制台会输出警告信息，请手动安装对应字体。

</details>

<details>
<summary><strong>Q: 转换时提示 "Permission denied" 怎么办？</strong></summary>

这通常是因为目标 `.docx` 文件正在被 Word 或其他程序打开。请关闭占用该文件的程序，然后重试。程序会弹出友好提示引导操作。

</details>

<details>
<summary><strong>Q: 粘贴的 Markdown 格式不规范（缺少空行），转换结果正常吗？</strong></summary>

正常。程序内置了智能预处理，会自动在标题前后补充空行，并将软回车拆分为独立的 Word 段落。

</details>

<details>
<summary><strong>Q: 样式 A 和样式 B 应该怎么选？</strong></summary>

- **样式 A**（Title 模式）：`#` 映射为 Word 的 "Title" 样式，适合 `#` 作为文档总标题的场景
- **样式 B**（Heading 模式）：`#` 映射为 Word 的 "Heading 1" 样式，适合 `#` 直接作为一级章节标题的场景

两种模式下同级标题的视觉效果（字体、字号）完全一致，区别仅在于 Word 大纲层级。

</details>

<details>
<summary><strong>Q: 可以在没有 GUI 的服务器上使用吗？</strong></summary>

可以。使用 CLI 命令行接口即可，不依赖 PyQt6。参见 [CLI 命令行接口](#-cli-命令行接口) 章节。但需确保服务器上安装了所需字体，或使用 `--skip-font-check` 跳过检测。

</details>

<details>
<summary><strong>Q: 如何将转换引擎移植到其他项目？</strong></summary>

直接拷贝 `.agents/skills/md2docx-converter/` 整个目录到目标项目即可。该目录是完全自包含的，包含转换引擎、样式配置、字体资源和 CLI 入口。目标项目只需安装 `python-docx` 和 `mistune` 两个依赖。

</details>

---

## 📄 License

MIT
