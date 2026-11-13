"""Mac / Windows 启动器共用：结束时等待 Enter，避免窗口一闪就关。"""

from __future__ import annotations

import sys
import traceback


def run_with_pause(main_fn) -> None:
    """执行脚本入口函数，结束后等待 Enter 再退出（便于查看输出）。"""
    exit_code = 0
    try:
        main_fn()
    except SystemExit as exc:
        code = exc.code
        if code is None:
            exit_code = 0
        elif isinstance(code, int):
            exit_code = code
        else:
            exit_code = 1
    except Exception as exc:
        print(f"\n运行出错: {exc}", flush=True)
        traceback.print_exc()
        exit_code = 1
    pause_terminal()
    sys.exit(exit_code)


def pause_terminal(message: str = "按 Enter 键关闭窗口...") -> None:
    print(f"\n{message}", flush=True)
    if sys.platform == "win32":
        _pause_windows()
    else:
        _pause_unix()


def prompt_terminal(message: str) -> str:
    print(message, end="", flush=True)
    if sys.platform == "win32":
        try:
            return input()
        except EOFError:
            return ""
    return _read_line_unix()


def _pause_windows() -> None:
    try:
        input()
    except EOFError:
        pass


def _read_line_unix() -> str:
    try:
        with open("/dev/tty", "r", encoding="utf-8", errors="replace") as tty:
            return tty.readline().rstrip("\r\n")
    except OSError:
        try:
            return input()
        except EOFError:
            return ""


def _pause_unix() -> None:
    _read_line_unix()
