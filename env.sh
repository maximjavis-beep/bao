#!/bin/bash
# bao 项目快捷激活脚本
# 用法: source env.sh

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
VENV_ACTIVATE="$SCRIPT_DIR/venv/bin/activate"

if [ -f "$VENV_ACTIVATE" ]; then
    source "$VENV_ACTIVATE"
    echo "✅ bao 环境已激活 (Python $(python3 --version 2>&1 | awk '{print $2}'))"
    echo "  可用命令: bao build from-files | bao build preview | bao version"
else
    echo "❌ 虚拟环境未找到，请先运行: python3 -m venv venv && source venv/bin/activate && pip install -r requirements.txt && pip install -e ."
fi
