"""Excel COM（PowerShell）共用：实时流式输出进度，避免卡住无反馈。"""

from __future__ import annotations

import subprocess
import sys
import time
from collections.abc import Callable, Mapping
from pathlib import Path

from progress_ui import ConsoleProgressBar, ProgressWindow, open_progress_display

DEFAULT_STAGE_LABELS: dict[str, str] = {
    "open": "正在启动 Excel 并打开文件...",
    "header": "复制表头格式...",
    "save": "正在保存文件...",
    "write_seq": "正在写入 __原序...",
    "sync": "正在同步页签文件...",
    "delete_street": "准备页签「街版」...",
    "formula": "正在批量写入公式...",
    "copy_sheet": "正在复制页签...",
    "columns": "正在批量列操作...",
    "scan": "B 列橙色筛选扫描...",
    "filter": "正在应用 B 列筛选...",
}

DEFAULT_TIMING_LABELS: dict[str, str] = {
    "open_ms": "打开 Excel",
    "formula_ms": "写公式",
    "header_ms": "复制表头",
    "copy_ms": "复制数据行",
    "save_ms": "保存",
    "write_seq_ms": "写序号",
    "sync_ms": "同步文件",
    "scan_ms": "筛选扫描",
}

DEFAULT_PROGRESS_LABELS: dict[str, str] = {
    "formula": "写公式进度",
    "copy": "复制进度",
    "copy_groups": "复制块进度",
    "copy_rows": "逐行复制进度",
    "write_seq": "写序号进度",
    "sync": "同步文件进度",
    "scan": "筛选扫描进度",
}


def com_encoding() -> str:
    return "gbk" if sys.platform == "win32" else "utf-8"


def safe_print(text: str) -> None:
    try:
        print(text, flush=True)
    except UnicodeEncodeError:
        print(
            text.encode(sys.stdout.encoding or "utf-8", errors="replace").decode(
                sys.stdout.encoding or "utf-8", errors="replace"
            ),
            flush=True,
        )


def format_com_line(
    line: str,
    *,
    stage_labels: Mapping[str, str] | None = None,
    timing_labels: Mapping[str, str] | None = None,
    progress_labels: Mapping[str, str] | None = None,
) -> str | None:
    """把 COM 输出行转成可读进度；返回 None 表示静默（如 STATS/SAVED）。"""
    stages = {**DEFAULT_STAGE_LABELS, **(stage_labels or {})}
    timings = {**DEFAULT_TIMING_LABELS, **(timing_labels or {})}
    progresses = {**DEFAULT_PROGRESS_LABELS, **(progress_labels or {})}

    if line.startswith("PROGRESS,stage,"):
        rest = line.split(",", 2)[2] if line.count(",") >= 2 else ""
        stage_name = rest.split(",", 1)[0]
        return f"    {stages.get(stage_name, rest)}"

    if line.startswith("PROGRESS,"):
        parts = line.split(",")
        if len(parts) >= 4:
            kind, done, total = parts[1], parts[2], parts[3]
            label = progresses.get(kind, f"{kind}进度")
            return f"    {label}: {done}/{total}"
        return f"  {line}"

    if line.startswith("TIMING,"):
        parts = line.split(",", 2)
        if len(parts) >= 3:
            label = timings.get(parts[1], parts[1])
            try:
                return f"    {label}完成: {int(parts[2]) / 1000:.1f}s"
            except ValueError:
                return f"  {line}"
        return f"  {line}"

    if line.startswith("STATS,") or line.startswith("SAVED,"):
        return None

    if line.strip():
        return f"  {line}"
    return None


class ComProgressDisplay:
    """把 COM 的 PROGRESS 协议转成横向进度条（窗口或终端单行刷新）。"""

    def __init__(
        self,
        title: str = "Excel 处理中",
        *,
        stage_labels: Mapping[str, str] | None = None,
        timing_labels: Mapping[str, str] | None = None,
        progress_labels: Mapping[str, str] | None = None,
        prefer_window: bool = False,
    ) -> None:
        self._stages = {**DEFAULT_STAGE_LABELS, **(stage_labels or {})}
        self._timings = {**DEFAULT_TIMING_LABELS, **(timing_labels or {})}
        self._progresses = {**DEFAULT_PROGRESS_LABELS, **(progress_labels or {})}
        self._bar: ProgressWindow | ConsoleProgressBar = open_progress_display(
            title,
            total=100,
            prefer_window=prefer_window,
        )

    def feed(self, line: str) -> str | None:
        if line.startswith("PROGRESS,stage,"):
            rest = line.split(",", 2)[2] if line.count(",") >= 2 else ""
            stage_name = rest.split(",", 1)[0]
            status = self._stages.get(stage_name, rest)
            self._bar.update(0, status=status)
            return None

        if line.startswith("PROGRESS,"):
            parts = line.split(",")
            if len(parts) >= 4:
                kind, done_s, total_s = parts[1], parts[2], parts[3]
                try:
                    done, total = int(done_s), int(total_s)
                except ValueError:
                    return format_com_line(
                        line,
                        stage_labels=self._stages,
                        timing_labels=self._timings,
                        progress_labels=self._progresses,
                    )
                label = self._progresses.get(kind, f"{kind}进度")
                self._bar.update(done, status=label, total=total)
            return None

        msg = format_com_line(
            line,
            stage_labels=self._stages,
            timing_labels=self._timings,
            progress_labels=self._progresses,
        )
        if msg is not None and hasattr(self._bar, "ensure_line"):
            self._bar.ensure_line()
        return msg

    def close(self) -> None:
        self._bar.close()


def stream_com(
    cmd: list[str],
    *,
    log_file: Path | None = None,
    encoding: str | None = None,
    on_line: Callable[[str], None] | None = None,
    print_progress: bool = True,
    progress_bar: bool = True,
    progress_title: str = "Excel 处理中",
    prefer_progress_window: bool = False,
    stage_labels: Mapping[str, str] | None = None,
    timing_labels: Mapping[str, str] | None = None,
    progress_labels: Mapping[str, str] | None = None,
    elapsed_label: str = "  COM 总耗时: {elapsed:.1f}s",
) -> tuple[int, list[str]]:
    """实时运行 COM 脚本，边跑边显示进度。返回 (returncode, 全部输出行)。"""
    t0 = time.time()
    enc = encoding or com_encoding()
    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        encoding=enc,
        errors="replace",
        bufsize=1,
    )
    output_lines: list[str] = []
    display: ComProgressDisplay | None = None
    if print_progress and progress_bar:
        display = ComProgressDisplay(
            progress_title,
            stage_labels=stage_labels,
            timing_labels=timing_labels,
            progress_labels=progress_labels,
            prefer_window=prefer_progress_window,
        )
    assert proc.stdout is not None
    try:
        for raw in proc.stdout:
            line = raw.rstrip("\r\n")
            output_lines.append(line)
            if on_line is not None:
                on_line(line)
            if not print_progress:
                continue
            if display is not None:
                msg = display.feed(line)
            else:
                msg = format_com_line(
                    line,
                    stage_labels=stage_labels,
                    timing_labels=timing_labels,
                    progress_labels=progress_labels,
                )
            if msg is not None:
                safe_print(msg)
    finally:
        if display is not None:
            display.close()
    returncode = proc.wait()
    if log_file is not None:
        log_file.write_text(
            "\n".join(output_lines) + ("\n" if output_lines else ""),
            encoding="utf-8",
        )
    if elapsed_label:
        safe_print(elapsed_label.format(elapsed=time.time() - t0))
    return returncode, output_lines
