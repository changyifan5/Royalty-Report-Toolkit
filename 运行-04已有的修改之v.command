#!/bin/bash
ROOT="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT"
source "$ROOT/mac_launcher_common.sh"
_activate_venv
python "04 已有的修改之v.py"

