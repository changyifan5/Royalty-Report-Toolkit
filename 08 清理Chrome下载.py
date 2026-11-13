"""
清理 Chrome 下载目录（代码库 08）

默认目录：Windows 为 D:\\chrome下载文件；Mac 为 ~/Downloads/chrome下载（可用 --dir 指定）
移入回收站的条件：
  · 乱码/hash 文件名（如 32 位 bf6174fa…）→ 不限时间，直接删
  · 超过 1 个月 且（Chrome 重复副本 / 含录音·词曲·授权结算表）

最近 1 个月内、名称正常的文件保留。

用法：
  双击运行 / python "08 清理Chrome下载.py"  → 扫描后直接弹窗询问是否删除
  python "08 清理Chrome下载.py" --preview   # 仅查看，不询问
  python "08 清理Chrome下载.py" -y           # 跳过确认，直接删除

运行：python "08 清理Chrome下载.py"
"""

from __future__ import annotations

import argparse
import re
import sys
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
LOG_FILE = BASE_DIR / "08_cleanup.log"

DEFAULT_DIRS = [
    Path(r"D:\chrome下载文件"),
    Path(r"D:\chrome下载"),
]
if sys.platform == "darwin":
    DEFAULT_DIRS.extend(
        [
            # 这台电脑固定目录：下载里的「chrome下载」
            Path.home() / "Downloads" / "chrome下载",
        ]
    )

DUP_SUFFIX = re.compile(r"^(.+?)\s*\((\d+)\)(\.[^.]+)$")

REASON_OLD = "超过1个月"
REASON_GARBLED = "乱码名"
REASON_DUP = "重复副本"
REASON_REPORT = "录音词曲结算"

# 简体 / 繁体 + 常见英文报表名
REPORT_KEYWORDS = (
    "录音", "錄音",
    "词曲", "詞曲",
    "授权结算表", "授權結算表",
    "授权结算", "授權結算",
)
REPORT_KEYWORDS_LOWER = (
    "recording report",
    "recording.xlsx",
    "composition.xlsx",
    "_recording",
    "_composition",
)


def safe_print(text: str) -> None:
    try:
        print(text)
    except UnicodeEncodeError:
        enc = sys.stdout.encoding or "utf-8"
        print(text.encode(enc, errors="replace").decode(enc, errors="replace"))


def resolve_target_dir(custom: str | None) -> Path:
    if custom:
        path = Path(custom)
        if not path.is_dir():
            raise SystemExit(f"目录不存在：{path}")
        return path

    for candidate in DEFAULT_DIRS:
        if candidate.is_dir():
            return candidate

    raise SystemExit(
        "未找到 Chrome 下载目录。请用 --dir 指定，例如：\n"
        '  python "08 清理Chrome下载.py" --dir "~/Downloads/chrome下载"'
    )


def file_mtime(path: Path) -> datetime:
    return datetime.fromtimestamp(path.stat().st_mtime)


def normalize_name(name: str) -> str:
    """Chrome 重复下载：foo (1).pdf → 与 foo.pdf 视为同一组。"""
    match = DUP_SUFFIX.match(name)
    if match:
        return match.group(1) + match.group(3)
    return name


def is_garbled_name(name: str) -> bool:
    """识别 Chrome 常见的 hash / 乱码文件名（不限修改时间）。"""
    stem = Path(name).stem
    match = DUP_SUFFIX.match(name)
    if match:
        stem = match.group(1)

    # 32 位 UUID hash：bf6174fa635f402d9165ed1d60d3b467
    if re.fullmatch(r"[0-9a-f]{32}", stem, re.IGNORECASE):
        return True
    # 更长或更短的纯十六进制名
    if re.fullmatch(r"[0-9a-f]{8,}", stem, re.IGNORECASE):
        return True
    # 14 位时间戳名：20260703162811
    if re.fullmatch(r"\d{14}", stem):
        return True
    return False


def is_report_name(name: str) -> bool:
    """文件名含录音 / 词曲 / 授权结算表（及常见变体）。"""
    if any(kw in name for kw in REPORT_KEYWORDS):
        return True
    lower = name.lower()
    return any(kw in lower for kw in REPORT_KEYWORDS_LOWER)


def collect_files_to_delete(root: Path, cutoff: datetime) -> dict[Path, set[str]]:
    all_files = [p for p in root.rglob("*") if p.is_file()]

    duplicate_extras: set[Path] = set()
    by_name: dict[str, list[Path]] = defaultdict(list)
    for path in all_files:
        by_name[normalize_name(path.name)].append(path)

    for group in by_name.values():
        if len(group) < 2:
            continue
        group.sort(key=lambda p: (file_mtime(p), p.name), reverse=True)
        duplicate_extras.update(group[1:])

    to_delete: dict[Path, set[str]] = defaultdict(set)
    for path in all_files:
        is_garbled = is_garbled_name(path.name)
        is_old = file_mtime(path) < cutoff
        is_dup = path in duplicate_extras
        is_report = is_report_name(path.name)

        if is_garbled:
            to_delete[path].add(REASON_GARBLED)
            if is_old:
                to_delete[path].add(REASON_OLD)
            continue

        if not is_old:
            continue

        if is_dup:
            to_delete[path].add(REASON_OLD)
            to_delete[path].add(REASON_DUP)
        if is_report:
            to_delete[path].add(REASON_OLD)
            to_delete[path].add(REASON_REPORT)

    return dict(to_delete)


def send_to_recycle_bin(path: Path) -> None:
    if sys.platform == "win32":
        import ctypes
        from ctypes import wintypes

        FO_DELETE = 0x0003
        FOF_ALLOWUNDO = 0x0040
        FOF_NOCONFIRMATION = 0x0010
        FOF_SILENT = 0x0004

        class SHFILEOPSTRUCTW(ctypes.Structure):
            _fields_ = [
                ("hwnd", wintypes.HWND),
                ("wFunc", wintypes.UINT),
                ("pFrom", wintypes.LPCWSTR),
                ("pTo", wintypes.LPCWSTR),
                ("fFlags", wintypes.UINT),
                ("fAnyOperationsAborted", wintypes.BOOL),
                ("hNameMappings", wintypes.LPVOID),
                ("lpszProgressTitle", wintypes.LPCWSTR),
            ]

        shell32 = ctypes.windll.shell32
        op = SHFILEOPSTRUCTW()
        op.wFunc = FO_DELETE
        op.pFrom = str(path.resolve()) + "\0\0"
        op.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT
        result = shell32.SHFileOperationW(ctypes.byref(op))
        if result != 0 or op.fAnyOperationsAborted:
            raise OSError(f"移入回收站失败（code={result}）：{path}")
        return

    path.unlink()


def format_size(num_bytes: int) -> str:
    units = ["B", "KB", "MB", "GB"]
    size = float(num_bytes)
    for unit in units:
        if size < 1024 or unit == units[-1]:
            if unit == "B":
                return f"{int(size)} {unit}"
            return f"{size:.1f} {unit}"
        size /= 1024
    return f"{num_bytes} B"


def append_log(lines: list[str]) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with LOG_FILE.open("a", encoding="utf-8") as f:
        f.write(f"\n=== {timestamp} ===\n")
        f.write("\n".join(lines))
        f.write("\n")


def remove_empty_dirs(root: Path) -> int:
    removed = 0
    for directory in sorted(root.rglob("*"), key=lambda p: len(p.parts), reverse=True):
        if directory == root:
            continue
        if directory.is_dir() and not any(directory.iterdir()):
            directory.rmdir()
            removed += 1
    return removed


def reason_counts(to_delete: dict[Path, set[str]]) -> dict[str, int]:
    counts: dict[str, int] = defaultdict(int)
    for reasons in to_delete.values():
        for reason in reasons:
            counts[reason] += 1
    return dict(counts)


def ask_confirm(count: int, summary: str) -> bool:
    """Windows 默认弹窗；其他系统用命令行询问。"""
    if sys.platform == "win32":
        return ask_confirm_gui("是否删除？", summary)

    safe_print(summary)
    safe_print("")
    answer = input(f"是否将以上 {count} 个文件移入回收站？[y/N] ").strip().lower()
    return answer in ("y", "yes")


def ask_confirm_gui(title: str, message: str) -> bool:
    import tkinter as tk
    from tkinter import messagebox

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    result = messagebox.askyesno(title, message, icon="warning")
    root.destroy()
    return result


def show_info_gui(title: str, message: str) -> None:
    import tkinter as tk
    from tkinter import messagebox

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    messagebox.showinfo(title, message)
    root.destroy()


def show_message(title: str, message: str) -> None:
    if sys.platform == "win32":
        show_info_gui(title, message)
    else:
        safe_print(f"{title}\n{message}")


def pause_before_exit() -> None:
    if sys.stdin.isatty():
        safe_print("")
        try:
            input("按 Enter 键退出...")
        except EOFError:
            pass


def main() -> None:
    parser = argparse.ArgumentParser(description="清理 Chrome 下载目录")
    parser.add_argument("--dir", help="下载目录路径（默认自动检测 D:\\chrome下载文件）")
    parser.add_argument("--months", type=int, default=1, help="保留最近 N 个月，默认 1")
    parser.add_argument("--preview", action="store_true", help="仅查看待清理列表，不询问、不删除")
    parser.add_argument("-y", "--yes", action="store_true", help="跳过确认，直接删除")
    args = parser.parse_args()

    try:
        target = resolve_target_dir(args.dir)
        cutoff = datetime.now() - timedelta(days=30 * args.months)
        to_delete = collect_files_to_delete(target, cutoff)
        delete_list = sorted(to_delete.keys(), key=file_mtime)

        total_size = sum(p.stat().st_size for p in delete_list)
        counts = reason_counts(to_delete)

        summary = (
            f"目录：{target}\n\n"
            f"待清理：{len(delete_list)} 个文件，共 {format_size(total_size)}\n"
            f"· 乱码/hash 名：{counts.get(REASON_GARBLED, 0)} 个\n"
            f"· 超过 {args.months} 个月且重复：{counts.get(REASON_DUP, 0)} 个\n"
            f"· 超过 {args.months} 个月且录音/词曲/结算表：{counts.get(REASON_REPORT, 0)} 个\n\n"
            f"移入回收站后可恢复。\n\n是否删除？"
        )

        if not delete_list:
            show_message("Chrome 下载清理", "没有需要清理的文件。")
            return

        if args.preview:
            safe_print(f"目标目录：{target}")
            safe_print(summary.replace("\n\n是否删除？", ""))
            safe_print("")
            safe_print("示例（最早 15 个）：")
            for path in delete_list[:15]:
                mtime = file_mtime(path).strftime("%Y-%m-%d %H:%M")
                tags = "/".join(sorted(to_delete[path]))
                safe_print(f"  [{mtime}] [{tags}] {path.name}  ({format_size(path.stat().st_size)})")
            if len(delete_list) > 15:
                safe_print(f"  ... 另有 {len(delete_list) - 15} 个")
            return

        confirmed = args.yes or ask_confirm(len(delete_list), summary)
        if not confirmed:
            show_message("已取消", "未删除任何文件。")
            return

        deleted: list[str] = []
        failed: list[str] = []

        for path in delete_list:
            try:
                send_to_recycle_bin(path)
                tags = "/".join(sorted(to_delete[path]))
                deleted.append(f"[{tags}] {path}")
            except OSError as exc:
                failed.append(f"{path}: {exc}")

        empty_removed = remove_empty_dirs(target)

        result = f"已移入回收站：{len(deleted)} 个文件"
        if empty_removed:
            result += f"\n已删除空文件夹：{empty_removed} 个"
        if failed:
            result += f"\n失败：{len(failed)} 个"

        show_message("清理完成", result)

        log_lines = [
            f"目录: {target}",
            f"截止: {cutoff.isoformat()}",
            f"统计: {counts}",
            f"删除: {len(deleted)} 个, 失败: {len(failed)} 个, 空目录: {empty_removed}",
        ] + deleted
        append_log(log_lines)
    finally:
        if args.preview and sys.stdin.isatty():
            pause_before_exit()


if __name__ == "__main__":
    main()
