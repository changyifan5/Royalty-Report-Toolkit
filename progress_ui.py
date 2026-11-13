"""进度条：tkinter 横向条 + 终端单行刷新条。"""

from __future__ import annotations

import sys
import tkinter as tk
from tkinter import ttk


class ProgressWindow:
    """置顶进度条：update(current, total, status) 刷新；close() 关闭。"""

    def __init__(self, title: str = "处理中", *, total: int = 100) -> None:
        self.total = max(1, total)
        self._closed = False
        self.root = tk.Tk()
        self.root.title(title)
        self.root.attributes("-topmost", True)
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", lambda: None)

        frame = ttk.Frame(self.root, padding=16)
        frame.grid(row=0, column=0, sticky="nsew")

        self.status_var = tk.StringVar(value="准备中…")
        self.count_var = tk.StringVar(value=f"0 / {self.total}")

        ttk.Label(frame, textvariable=self.status_var, width=48).grid(
            row=0, column=0, sticky="w", pady=(0, 8)
        )
        self.bar = ttk.Progressbar(
            frame,
            orient="horizontal",
            length=360,
            mode="determinate",
            maximum=self.total,
        )
        self.bar.grid(row=1, column=0, sticky="ew")
        ttk.Label(frame, textvariable=self.count_var).grid(
            row=2, column=0, sticky="e", pady=(6, 0)
        )

        self.root.update_idletasks()
        w, h = self.root.winfo_width(), self.root.winfo_height()
        x = (self.root.winfo_screenwidth() - w) // 2
        y = (self.root.winfo_screenheight() - h) // 3
        self.root.geometry(f"+{x}+{y}")
        self.root.update()

    def update(self, current: int, *, status: str | None = None, total: int | None = None) -> None:
        if self._closed:
            return
        if total is not None and total > 0:
            self.total = total
            self.bar.configure(maximum=self.total)
        current = max(0, min(current, self.total))
        self.bar["value"] = current
        self.count_var.set(f"{current} / {self.total}")
        if status is not None:
            self.status_var.set(status)
        try:
            self.root.update()
        except tk.TclError:
            self._closed = True

    def close(self) -> None:
        if self._closed:
            return
        self._closed = True
        try:
            self.root.destroy()
        except tk.TclError:
            pass


class ConsoleProgressBar:
    """终端内从左到右的进度条，用 \\r 原地刷新，不逐行刷屏。"""

    def __init__(self, label: str = "处理中", *, total: int = 100, width: int = 36) -> None:
        self.total = max(1, total)
        self.width = width
        self._label = label
        self._active = True
        self._last_len = 0

    def update(self, current: int, *, status: str | None = None, total: int | None = None) -> None:
        if not self._active:
            return
        if status is not None:
            self._label = status
        if total is not None and total > 0:
            self.total = total
        current = max(0, min(current, self.total))
        pct = 100.0 * current / self.total
        filled = int(self.width * current / self.total)
        bar = "█" * filled + "░" * (self.width - filled)
        text = f"  {self._label} [{bar}] {current}/{self.total} ({pct:.0f}%)"
        pad = max(0, self._last_len - len(text))
        self._last_len = len(text)
        try:
            sys.stdout.write("\r" + text + " " * pad)
            sys.stdout.flush()
        except OSError:
            pass

    def close(self) -> None:
        if not self._active:
            return
        self._active = False
        try:
            sys.stdout.write("\n")
            sys.stdout.flush()
        except OSError:
            pass

    def ensure_line(self) -> None:
        """阶段日志输出前换行，避免冲掉当前进度条行。"""
        if not self._active:
            return
        try:
            sys.stdout.write("\n")
            sys.stdout.flush()
            self._last_len = 0
        except OSError:
            pass


def open_progress_display(
    title: str,
    *,
    total: int = 100,
    prefer_window: bool = False,
) -> ProgressWindow | ConsoleProgressBar:
    """默认终端内横向进度条；prefer_window=True 时尝试 tk 弹窗。"""
    if prefer_window:
        try:
            return ProgressWindow(title, total=total)
        except tk.TclError:
            pass
    return ConsoleProgressBar(label=title, total=total)
