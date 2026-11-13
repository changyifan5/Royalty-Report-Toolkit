#!/bin/bash
cd "$(dirname "$0")" || exit 1
if [[ -x "$HOME/.venv/bin/python" ]]; then
  PACKER_PYTHON="$HOME/.venv/bin/python"
else
  PACKER_PYTHON="python3"
fi
"$PACKER_PYTHON" "20 Packer比对汇出.py" "$@"
printf '\n按 Enter 关闭窗口…'
read -r
