#!/bin/bash
ROOT="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT"
source "$ROOT/mac_launcher_common.sh"
_activate_venv
python3 "24 报表进度台.py"
status=$?
if [[ $status -ne 0 ]]; then
  read -r -p "启动失败，按 Enter 键关闭窗口..."
  exit "$status"
fi

# 仅关闭由 Finder 启动 .command 的单标签 Terminal 窗口。
# 手动在终端运行时，父 shell 仍在，保留该窗口。
if [[ "${TERM_PROGRAM:-}" == "Apple_Terminal" && -t 0 ]]; then
  terminal_tty="$(tty)"
  shell_pid="$PPID"
  window_id="$(/usr/bin/osascript - "$terminal_tty" <<'APPLESCRIPT'
on run argv
    set terminalTTY to item 1 of argv
    tell application "Terminal"
        repeat with candidate in windows
            if (count of «class ttab» of candidate) is 1 then
                if «class ttty» of «class ttab» 1 of candidate is terminalTTY then return id of candidate
            end if
        end repeat
    end tell
    return ""
end run
APPLESCRIPT
)"
  if [[ "$window_id" =~ ^[0-9]+$ ]]; then
    (
      for attempt in {1..20}; do
        if ! kill -0 "$shell_pid" 2>/dev/null; then
          /usr/bin/osascript - "$window_id" "$terminal_tty" <<'APPLESCRIPT'
on run argv
    set terminalID to (item 1 of argv) as integer
    set terminalTTY to item 2 of argv
    tell application "Terminal"
        if exists window id terminalID then
            set targetWindow to window id terminalID
            if (count of «class ttab» of targetWindow) is 1 then
                set targetTab to «class ttab» 1 of targetWindow
                set currentTTY to «class ttty» of targetTab
                if (currentTTY is terminalTTY or currentTTY is "") and not («class busy» of targetTab) then
                    close targetWindow
                end if
            end if
        end if
    end tell
end run
APPLESCRIPT
          break
        fi
        sleep 0.15
      done
    ) </dev/null >/dev/null 2>&1 &
    disown
  fi
fi
exit 0
