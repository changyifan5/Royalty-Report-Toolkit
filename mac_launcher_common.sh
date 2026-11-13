#!/bin/bash
# Mac .command 共用；暂停/输入实际由 launcher_common.py（Win/Mac 同一套逻辑）

ROOT="$(cd "$(dirname "$0")" && pwd)"

_activate_venv() {
  if [[ -f "$HOME/.venv/bin/activate" ]]; then
    # shellcheck disable=SC1091
    source "$HOME/.venv/bin/activate"
  elif [[ -f "$ROOT/.venv/bin/activate" ]]; then
    # shellcheck disable=SC1091
    source "$ROOT/.venv/bin/activate"
  fi
}

_py() {
  command python "$@"
}

pause_terminal() {
  _py "$ROOT/pause.py"
}

prompt_terminal() {
  # prompt_terminal "提示: " 变量名
  local __var=$2
  printf -v "$__var" '%s' "$(_py "$ROOT/prompt_cli.py" "$1")"
}
