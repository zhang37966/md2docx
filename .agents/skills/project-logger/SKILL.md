---
name: project-logger
description: 自动记录对话结果到项目开发日志 (DEVELOPMENT_LOG.md)
---
# Project Logger Skill

此 Skill 旨在自动维护项目的 `DEVELOPMENT_LOG.md` 文件。

## 核心规则

1. **自动记录**：在每次与用户的重要对话或任务结束时，主动检查是否需要更新开发日志。
2. **位置**：新的记录**必须**插入在本项目文件夹的 `DEVELOPMENT_LOG.md` 文件的**最上方**（保留文件原有的 Markdown 结构）。
3. **频率**：每个独立的任务或功能修复应产生一个条目。

## 记录格式

必须严格遵守以下 Markdown 格式：

```markdown
#### YYYY-MM-DD HH:MM - [对话主题/任务简述]
- **任务表述**：[简洁描述用户的问题或需求]
- **完成情况**：[说明最终解决结果或当前进度]
```

## 执行步骤

1. 确认任务已完成或达到阶段性结果。
2. 获取当前日期 (YYYY-MM-DD HH:MM)。
3. 提取对话的核心主题。
4. 使用 `replace_file_content` 或 `multi_replace_file_content` 将新条目插入到 `DEVELOPMENT_LOG.md` 的第一行。
