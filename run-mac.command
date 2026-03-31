#!/bin/bash
# MD2DOCX 一键运行脚本 (Mac)

# 切换到脚本自身所在的目录（而非用户的 home 目录）
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

echo "正在启动 MD2DOCX (Mac版)..."
echo "工作目录: $SCRIPT_DIR"

# 检查是否安装了 uv
if ! command -v uv &> /dev/null; then
    echo "未检测到 uv 环境，正在自动安装..."
    curl -LsSf https://astral.sh/uv/install.sh | sh
    # 刷新 PATH 以识别刚安装的 uv
    export PATH="$HOME/.cargo/bin:$HOME/.local/bin:$PATH"
    if ! command -v uv &> /dev/null; then
        echo "uv 安装失败，请手动运行: curl -LsSf https://astral.sh/uv/install.sh | sh"
        read -n 1 -s -p "按任意键退出..."
        exit 1
    fi
    echo "✅ uv 安装成功！"
fi

# 确保依赖同步
echo "正在同步依赖..."
uv sync

# 运行主程序
uv run python main.py
