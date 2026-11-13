#!/bin/bash

ROOT="$(cd "$(dirname "$0")" && pwd)"

cd "$ROOT"

source "$ROOT/mac_launcher_common.sh"

_activate_venv

python "19 F3汇总.py"

