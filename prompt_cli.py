#!/usr/bin/env python3
"""供 shell 启动器读取一行输入：python prompt_cli.py "提示文字" """
from __future__ import annotations

import sys

from launcher_common import prompt_terminal

if __name__ == "__main__":
    msg = sys.argv[1] if len(sys.argv) > 1 else "请输入: "
    sys.stdout.write(prompt_terminal(msg))
