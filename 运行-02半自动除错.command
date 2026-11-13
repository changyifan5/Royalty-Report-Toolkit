#!/bin/bash
ROOT="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT"
source "$ROOT/mac_launcher_common.sh"
_activate_venv
python "02 半自动除错.py"

