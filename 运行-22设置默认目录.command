#!/bin/bash
ROOT="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT"
source "$ROOT/mac_launcher_common.sh"
_activate_venv
python3 "22 设置默认目录.py"
status=$?
read -r -p "按 Enter 键关闭窗口..."
exit "$status"
