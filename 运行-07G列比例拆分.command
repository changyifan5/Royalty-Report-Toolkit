#!/bin/bash
ROOT="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT"
source "$ROOT/mac_launcher_common.sh"
_activate_venv
python "07 G列比例拆分.py"

