---
name: md2docx-converter
description: 将 Markdown 内容转换为格式化的 Word (.docx) 文档，支持中文公文排版样式
---

# MD2DOCX Converter Skill

将 Markdown 内容快速转换为带有标准中文公文排版样式的 Word 文档。

## 触发条件

当用户明确要求或暗示需要以下操作时，应主动使用此 Skill：

- 将 Markdown 内容 **导出 / 转换为 Word 文档**
- 将 AI 生成的报告、方案、文档**输出为 .docx 格式**
- 任何涉及 "转 Word"、"导出 docx"、"生成文档" 的请求

## Skill 目录结构

```
md2docx-converter/
├── SKILL.md              # 本文件
├── requirements.txt      # Python 依赖
├── scripts/
│   └── convert.py        # CLI 转换入口
├── converter/            # 转换引擎
│   ├── md_parser.py      # Markdown AST 解析器
│   └── docx_writer.py    # DOCX 文档生成引擎
├── styles/
│   └── docx_styles.py    # 全局排版样式配置
├── utils/
│   └── font_installer.py # 字体自动安装
├── font/                 # 内嵌字体（仿宋_GB2312、宋体、黑体）
└── examples/
    └── sample.md         # 测试样例
```

## 前置依赖

Skill 运行需要以下 Python 包，已列于 `requirements.txt` 中：

- `python-docx>=1.2.0`
- `mistune>=3.0.0`

**首次使用前**，请确认依赖已安装。若未安装，执行：

```bash
pip install -r <SKILL_DIR>/requirements.txt
```

## 转换命令

**定位 Skill 目录**：首先确定 Skill 在当前工作区中的实际路径。通常位于：

```
<工作区根目录>/.agents/skills/md2docx-converter/
```

**命令格式**：

```bash
# 方式一：从 .md 文件转换
python <SKILL_DIR>/scripts/convert.py --input <md文件路径> --output <docx路径> --style <A|B>

# 方式二：直接传入 Markdown 文本
python <SKILL_DIR>/scripts/convert.py --text "<markdown文本>" --output <docx路径> --style <A|B>
```

**参数说明**：

| 参数 | 必填 | 说明 |
|---|---|---|
| `--input` / `-i` | 二选一 | 输入的 `.md` 文件路径 |
| `--text` / `-t` | 二选一 | 直接传入 Markdown 文本内容 |
| `--output` / `-o` | ✅ | 输出的 `.docx` 文件路径 |
| `--style` / `-s` | ❌ | 样式模式：`A`（默认）或 `B` |
| `--skip-font-check` | ❌ | 跳过字体检测与自动安装 |

**退出码**：`0` = 成功，`1` = 一般错误，`2` = 文件占用

## 样式说明

默认使用**样式 A**。你可以根据上下文建议用户选择：

- **样式 A（Title 模式）**：`#` 映射为 Word 的 Title 样式。适合**单篇独立文档**，如报告、方案、通知。
- **样式 B（Heading 模式）**：`#` 映射为 Word 的 Heading 1 样式。适合**多章节长文档**，如手册、规范、论文。

如果用户指定了样式，按用户要求执行。

## 执行步骤

### 步骤 1：准备 Markdown 内容

- 如果内容已在对话中生成（如 AI 撰写的报告），将其写入一个临时 `.md` 文件。
- 如果用户提供了现有的 `.md` 文件路径，直接使用。

### 步骤 2：确定输出路径

根据上下文**自动选择**合适的输出路径：

1. 用户明确指定了输出位置 → 使用用户指定的路径
2. 有活动工作区 → 保存到工作区根目录
3. 兜底 → 保存到用户桌面

文件名应有意义，基于内容标题或用户需求命名（如 `项目方案.docx`、`会议纪要.docx`）。

### 步骤 3：定位 Skill 并执行转换

1. 确定当前工作区中 Skill 的安装路径（搜索 `.agents/skills/md2docx-converter/scripts/convert.py`）
2. 使用 `run_command` 调用转换脚本

如果 Markdown 内容较短（≤ 2000 字符），可以用 `--text` 参数直接传入；否则写入临时 `.md` 文件后用 `--input` 传入。

### 步骤 4：报告结果

- **成功**：告知用户文档已生成，给出完整文件路径。
- **失败**：显示错误信息，退出码 2 表示文件被占用，建议关闭 Word 后重试。

## 排版能力

- 正文：仿宋_GB2312，小四（12pt），首行缩进 2 字符，两端对齐
- 一级标题：黑体加粗，18pt，居中
- 其余标题：宋体，按级递减字号
- 页脚：自动生成「第 X 页 共 Y 页」居中页码
- 支持：表格、代码块（底色+边框）、引用块（蓝色边条）、有序/无序列表、超链接、图片嵌入

## 移植指南

将整个 `md2docx-converter/` 文件夹拷贝到目标项目的 `.agents/skills/` 下即可使用：

```bash
cp -r md2docx-converter/ <目标项目>/.agents/skills/
```

移植后需确认：
1. 目标环境已安装 Python 3.9+ 和 `requirements.txt` 中的依赖
2. 首次运行时脚本会自动安装缺失字体（仿宋_GB2312、宋体、黑体）

## 维护说明

Skill 内的 `converter/`、`styles/`、`utils/` 是转换逻辑的**唯一来源**。如需修改排版参数或转换逻辑，请直接编辑 Skill 目录内的对应文件。

> **重要**：`styles/docx_styles.py` 是样式的唯一真实来源。`converter/docx_writer.py` 的渲染逻辑必须始终引用其导出变量，不允许硬编码格式参数。
