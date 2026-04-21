# 上下文检查点

> 最后更新：2026-04-21 22:28

## 🎯 当前焦点

项目核心功能稳定。刚完成 README.md 重写和 Windows 启动体验优化（隐藏 CMD 窗口 + 中文乱码修复）。新增了 context-manager Skill 用于跨对话上下文接力。

## 📁 近期修改文件

| 文件 | 修改摘要 |
|---|---|
| `README.md` | 全面重写，居中标题+徽章+功能亮点+快速开始 |
| `启动 MD2DOCX.vbs` | 新增，VBScript 隐藏 CMD 窗口启动 |
| `run-win.bat` | 增加 `chcp 65001` 修复中文乱码 |
| `.agents/skills/context-manager/SKILL.md` | 新增，上下文管理 Skill |

## ⚠️ 待办 / 未完成

- [ ] 考虑 Mac 平台的无窗口启动方案
- [ ] dist/MD2DOCX.exe 已从 Git 移除，需要时重新打包

## 🧠 关键决策与上下文

- **样式修改铁律**：先改 `docx_styles.py` 再改 `docx_writer.py`，二者一一匹配
- 核心代码已全部迁移至 `.agents/skills/md2docx-converter/` 目录
- 项目使用 `uv` 管理依赖（`pyproject.toml` + `uv.lock`），无 `requirements.txt`
- GUI 基于 PyQt6，Markdown 解析使用 Mistune v3
- 用户偏好：全过程使用中文

## 🔗 相关文件速查

- 入口：`main.py`（含 sys.path 桥接到 Skill 目录）
- GUI：`ui/main_window.py`
- 转换核心：`.agents/skills/md2docx-converter/converter/docx_writer.py`
- 样式配置：`.agents/skills/md2docx-converter/styles/docx_styles.py`
- 解析器：`.agents/skills/md2docx-converter/converter/md_parser.py`
- CLI 入口：`.agents/skills/md2docx-converter/scripts/convert.py`
- 字体工具：`.agents/skills/md2docx-converter/utils/font_installer.py`
- 打包脚本：`build.py`
