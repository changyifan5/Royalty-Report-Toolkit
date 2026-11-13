#!/bin/bash
ROOT="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT"
source "$ROOT/mac_launcher_common.sh"
_activate_venv
python "16 词曲比对聚焦回写.py"
