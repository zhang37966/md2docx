#!/bin/bash
# MD2DOCX 一键运行脚本文本 (Mac)
# 会自动激活 uv 环境并运行

# 切换到脚本所在目录
cd "$(dirname "$0")"

echo "正在启动 MD2DOCX (Mac版)..."

# 检查是否安装了 uv
if ! command -v uv &> /dev/null; then
    echo "未检测到 uv 环境，请先安装 uv (curl -LsSf https://astral.sh/uv/install.sh | sh)"
    read -n 1 -s -p "按任意键退出..."
    exit 1
fi

# 确保依赖同步（静默）
uv sync --quiet

# 运行主程序
uv run main.py
