# MD2DOCX

<p align="center">
  <strong>Markdown → Word 一键转换桌面工具</strong>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.9+-3776AB?logo=python&logoColor=white" alt="Python">
  <img src="https://img.shields.io/badge/GUI-PyQt6-41CD52?logo=qt&logoColor=white" alt="PyQt6">
  <img src="https://img.shields.io/badge/Word-python--docx-2B579A?logo=microsoftword&logoColor=white" alt="python-docx">
  <img src="https://img.shields.io/badge/Parser-mistune_3-FF6B6B" alt="mistune">
</p>

---

**MD2DOCX** 是一款基于 Python 和 PyQt6 的轻量级桌面应用，专为**中文公文 / 正式文档**场景设计。只需拖入 `.md` 文件，即可生成带有标准 Word 样式的 `.docx` 文档——自动排版、自动页码、自动字体安装，开箱即用。

## ✨ 功能亮点

### 📝 Markdown 全要素解析

基于 `mistune` AST 引擎，精准转换以下元素：

- **文本格式** — 加粗、斜体
- **标题** — H1 ~ H6，支持两种 Word 样式映射
- **列表** — 有序 / 无序 / 多层嵌套
- **代码** — 行内代码 & 代码块（等宽字体 + 底色边框）
- **表格** — 自动加粗表头
- **引用块** — 左侧蓝色边条 + 淡蓝底色
- **超链接** — 蓝色下划线可点击
- **本地图片** — 自动嵌入
- **分隔线**

### 🎨 专业中文排版

所有排版参数集中管理在 `styles/docx_styles.py`，开箱即符合公文规范：

| 元素 | 字体 | 字号 | 特殊说明 |
|---|---|---|---|
| 正文 | 仿宋_GB2312 | 小四 (12pt) | 两端对齐、首行缩进 2 字符 |
| 一级标题 | 黑体 + 加粗 | 18pt | 居中、段后 16 磅 |
| 其余标题 | 宋体 | 按级递减 | 段前 1 行、段后 0.5 行 |
| 代码块 | Consolas | 10pt | 灰底边框 |
| 页脚 | — | — | 「第 X 页 共 Y 页」居中页码 |

### 📄 两种标题映射模式

| Markdown | 样式 A (Title 模式) | 样式 B (Heading 模式) |
|---|---|---|
| `#` | Title (18pt) | Heading 1 (18pt) |
| `##` | Heading 1 (16pt) | Heading 2 (16pt) |
| `###` | Heading 2 (14pt) | Heading 3 (14pt) |
| `####` | Heading 3 (12pt) | Heading 4 (12pt) |
| `#####` | Heading 4 (11pt) | Heading 5 (11pt) |
| `######` | Heading 5 (10.5pt) | Heading 6 (10.5pt) |

> 两种模式视觉效果完全一致，区别仅在 Word 大纲层级。

### 🔧 自动化特性

| 特性 | 说明 |
|---|---|
| 🔤 字体自动安装 | 首次启动静默安装仿宋_GB2312（免管理员权限） |
| 📄 智能保存 | 输出文件自动保存到源文件同目录 |
| 👀 快速预览 | 转换完成后一键打开 Word 文档 |
| ⚠️ 占用检测 | 目标文件被占用时友好提示，不会崩溃 |

---

## 🚀 快速开始

### 环境要求

- Python 3.9+
- Windows（字体安装功能依赖 Windows 注册表）

### 安装依赖

```bash
pip install -r requirements.txt
```

### 启动应用

```bash
python main.py
```

### 使用流程

1. **拖入文件** — 将 `.md` 文件拖到窗口中央，或点击 📂 按钮选择
2. **选择样式** — 样式 A（Title 模式）或 样式 B（Heading 模式）
3. **一键转换** — 点击 🚀 开始转换
4. **预览结果** — 在弹出对话框中点击「打开文件」

---

## 📦 打包为 EXE

```bash
python build.py
```

自动完成：检测并安装 PyInstaller → 清除缓存 → 打包字体资源 → 生成 `dist/MD2DOCX.exe` 单文件程序。

---

## 📂 项目结构

```text
md2docx/
├── main.py                  # 程序入口（含字体自动安装）
├── build.py                 # 一键打包脚本
├── requirements.txt         # Python 依赖
├── converter/
│   ├── md_parser.py         # Markdown AST 解析器 (mistune)
│   └── docx_writer.py       # DOCX 生成 & 格式映射引擎
├── styles/
│   └── docx_styles.py       # 全局样式配置（唯一真实来源）
├── ui/
│   └── main_window.py       # PyQt6 主窗口交互逻辑
├── utils/
│   └── font_installer.py    # 字体自动检测与安装
└── font/
    └── 仿宋_GB2312.ttf      # 内嵌字体（打包时分发）
```

---

## 🧩 二次开发

### 调整排版样式

所有样式参数集中在 [`styles/docx_styles.py`](styles/docx_styles.py) 中：

| 配置项 | 控制范围 |
|---|---|
| `FONT_CONFIG` | 正文 / 标题 / 一级标题 / 代码的字体与字号 |
| `NORMAL_*` | 基础正文 (Normal) 的行距与间距 |
| `BODY_*` | 正文文本 (Body Text) 的行距、缩进、间距 |
| `HEADING_*` | 标题样式的行距、间距、居中等级 |
| `HEADING_FONT_SIZES_A/B` | 两种模式的标题字号映射 |
| `CODE_BLOCK_*` | 代码块背景色与边框色 |

> **⚠️ 重要规则**：`docx_styles.py` 是样式的唯一真实来源。`docx_writer.py` 必须始终引用其导出变量，禁止硬编码格式参数。

### 扩展新语法支持

1. 在 `converter/md_parser.py` 的 Mistune 配置中启用对应 `plugins`
2. 在 `converter/docx_writer.py` 的 `_process_node()` 中处理新 AST 节点

---

## 🛠️ 技术栈

| 组件 | 技术 | 用途 |
|---|---|---|
| GUI 框架 | PyQt6 | 桌面窗口与交互 |
| 文档生成 | python-docx ≥ 1.2.0 | Word 文档创建与样式控制 |
| Markdown 解析 | mistune ≥ 3.0.0 | AST 语法树解析 |
| 打包工具 | PyInstaller | 生成独立 EXE |
