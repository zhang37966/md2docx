@echo off
chcp 65001 >nul
:: MD2DOCX 一键运行脚本 (Windows)

cd /d "%~dp0"

:: 检查是否安装了 uv
where uv >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo 未检测到 uv 环境，请先安装 uv
    pause
    exit /b 1
)

:: 确保依赖同步（静默）
uv sync --quiet

:: 运行主程序
uv run main.py
