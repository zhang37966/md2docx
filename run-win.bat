@echo off
:: MD2DOCX 一键运行脚本文本 (Windows)
:: 会自动激活 uv 环境并运行

cd /d "%~dp0"

echo 正在启动 MD2DOCX (Windows版)...

:: 检查是否安装了 uv
where uv >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo 未检测到 uv 环境，请先安装 uv ^(powershell -c "irm https://astral.sh/uv/install.ps1 | iex"^)
    pause
    exit /b 1
)

:: 确保依赖同步（静默）
uv sync --quiet

:: 运行主程序
uv run main.py
