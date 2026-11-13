"""
词曲比对聚焦回写脚本（代码库 16）

配合 15 导出的 `{主文件名}_词曲比对false聚焦.xlsx` 使用：
  - 聚焦清单 E 列「处理状态」默认为「待处理」
  - 人工看完某分组后，删除该行「待处理」三字
  - 本脚本读取所有**不再含「待处理」**的分组，将主报表中对应 C/D 且「词曲比对」为 false 的行**清空**

弹窗可选聚焦文件或主报表（同目录需有聚焦文件）。运行前请关闭 Excel。
运行：python "16 词曲比对聚焦回写.py"
"""

from __future__ import annotations

import shutil
import sys
import time
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import ttk

from openpyxl import load_workbook

from file_picker_common import confirm_process_file, pick_excel_file

BASE_DIR = Path(__file__).resolve().parent
SHEET_INDEX = 0
HEADER_ROW = 6
DATA_START_ROW = 7
HEADER_SCAN_LIMIT = 500
FOCUS_HEADER_ROW = 1

COL_ARTIST = 3
COL_ALBUM = 4

LYRICS_CMP_HEADER = "词曲比对"
FOCUS_SUFFIX = "_词曲比对false聚焦"
STATUS_HEADER = "处理状态"
PENDING_TEXT = "待处理"


@dataclass
class FocusGroup:
    artist: str
    album: str
    row_count: int
    status: str


def cell_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return str(value).strip()


def to_number(value) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    try:
        return float(text)
    except ValueError:
        return 0.0


def is_cmp_false(value) -> bool:
    if value is None:
        return False
    if isinstance(value, bool):
        return not value
    if isinstance(value, (int, float)):
        return value == 0
    text = str(value).strip().lower()
    return text in ("false", "0", "否", "no")


def read_header_map(sheet, header_row: int) -> dict[str, int]:
    mapping: dict[str, int] = {}
    for col in range(1, HEADER_SCAN_LIMIT + 1):
        val = sheet.cell(row=header_row, column=col).value
        if val is not None and str(val).strip():
            mapping[str(val).strip()] = col
    return mapping


def resolve_paths(picked: Path) -> tuple[Path, Path]:
    """返回 (主报表, 聚焦文件)。"""
    stem = picked.stem
    if stem.endswith(FOCUS_SUFFIX):
        main_stem = stem[: -len(FOCUS_SUFFIX)]
        main_path = picked.with_name(f"{main_stem}{picked.suffix}")
        return main_path, picked

    main_path = picked
    focus_path = picked.with_name(f"{picked.stem}{FOCUS_SUFFIX}{picked.suffix}")
    return main_path, focus_path


def is_pending_status(status: str) -> bool:
    return cell_text(status) == PENDING_TEXT


def load_focus_groups(focus_path: Path) -> tuple[list[FocusGroup], list[FocusGroup]]:
    """返回 (已完成分组, 仍待处理分组)。"""
    wb = load_workbook(focus_path, read_only=True, data_only=True, keep_links=False)
    ws = wb.active
    header_map = read_header_map(ws, FOCUS_HEADER_ROW)

    if STATUS_HEADER not in header_map:
        wb.close()
        print(f"错误：聚焦文件第 {FOCUS_HEADER_ROW} 行找不到列「{STATUS_HEADER}」")
        sys.exit(1)

    status_col = header_map[STATUS_HEADER]
    artist_col = min(header_map.values())
    album_candidates = sorted(c for c in header_map.values() if c != artist_col and c != status_col)
    album_col = album_candidates[0] if album_candidates else artist_col + 1

    done: list[FocusGroup] = []
    pending: list[FocusGroup] = []

    for row_idx, row in enumerate(
        ws.iter_rows(min_row=FOCUS_HEADER_ROW + 1, values_only=True),
        start=FOCUS_HEADER_ROW + 1,
    ):
        if not row:
            continue

        artist = cell_text(row[artist_col - 1] if len(row) >= artist_col else None)
        if artist == "合计":
            break

        album = cell_text(row[album_col - 1] if len(row) >= album_col else None)
        status = cell_text(row[status_col - 1] if len(row) >= status_col else None)
        row_count = int(to_number(row[2] if len(row) >= 3 else 0))

        item = FocusGroup(artist=artist, album=album, row_count=row_count, status=status)
        if is_pending_status(status):
            pending.append(item)
        else:
            done.append(item)

    wb.close()
    return done, pending


def confirm_sync(done_groups: list[FocusGroup], main_path: Path, focus_path: Path) -> bool:
    root = tk.Tk()
    root.title("16 词曲比对聚焦回写 — 确认")
    root.attributes("-topmost", True)
    root.geometry("760x420")

    frame = ttk.Frame(root, padding=10)
    frame.pack(fill=tk.BOTH, expand=True)

    total_rows = sum(g.row_count for g in done_groups)
    ttk.Label(
        frame,
        text=(
            f"聚焦文件: {focus_path.name}\n"
            f"主报表: {main_path.name}\n\n"
            f"以下 {len(done_groups)} 个分组已去掉「{PENDING_TEXT}」，"
            f"将把主报表中对应 C/D 且「{LYRICS_CMP_HEADER}」为 false 的行清空"
            f"（约 {total_rows} 行）："
        ),
        wraplength=720,
        justify="left",
    ).pack(anchor="w", pady=(0, 8))

    listbox = tk.Listbox(frame, height=min(12, len(done_groups)), font=("Consolas", 10))
    for group in done_groups:
        listbox.insert(
            tk.END,
            f"{group.artist}  |  {group.album}  |  {group.row_count} 行  |  状态: {group.status or '(空)'}",
        )
    listbox.pack(fill=tk.BOTH, expand=True)

    choice = {"ok": False}

    def on_ok() -> None:
        choice["ok"] = True
        root.destroy()

    def on_cancel() -> None:
        root.destroy()

    btn = ttk.Frame(root, padding=(10, 0, 10, 10))
    btn.pack(fill=tk.X)
    ttk.Button(btn, text="确定回写", command=on_ok, width=12).pack(side=tk.RIGHT, padx=(6, 0))
    ttk.Button(btn, text="取消", command=on_cancel, width=10).pack(side=tk.RIGHT)

    root.protocol("WM_DELETE_WINDOW", on_cancel)
    root.mainloop()
    return choice["ok"]


def backup_file(path: Path) -> Path:
    backup_path = path.with_suffix(path.suffix + ".bak")
    shutil.copy2(path, backup_path)
    return backup_path


def assert_file_writable(path: Path) -> None:
    """Excel 打开时 Windows 无法覆盖保存，先检测。"""
    try:
        with open(path, "r+b"):
            pass
    except PermissionError:
        print()
        print("=" * 50)
        print("错误：主报表被占用，无法保存")
        print("=" * 50)
        print(f"  文件: {path}")
        print("  请先在 Excel 中关闭该文件，然后重新运行本脚本。")
        sys.exit(1)


def save_workbook_with_fallback(wb, main_path: Path) -> Path:
    try:
        wb.save(main_path)
        return main_path
    except PermissionError:
        alt_path = main_path.with_name(f"{main_path.stem}_聚焦回写{main_path.suffix}")
        if alt_path.exists():
            alt_path.unlink()
        wb.save(alt_path)
        print()
        print("=" * 50)
        print("警告：主报表被占用，已另存为")
        print("=" * 50)
        print(f"  {alt_path}")
        print("  请关闭 Excel 中的原文件后，用此文件替换，或关闭 Excel 后重新运行本脚本。")
        return alt_path


def apply_completed_groups(
    ws,
    row_start: int,
    row_end: int,
    cmp_col: int,
    done_keys: set[tuple[str, str]],
) -> int:
    updated = 0
    for row in range(row_start, row_end + 1):
        cmp_val = ws.cell(row=row, column=cmp_col).value
        if not is_cmp_false(cmp_val):
            continue

        artist = cell_text(ws.cell(row=row, column=COL_ARTIST).value)
        album = cell_text(ws.cell(row=row, column=COL_ALBUM).value)
        if (artist, album) not in done_keys:
            continue

        ws.cell(row=row, column=cmp_col, value=None)
        updated += 1
    return updated


def run_sync() -> None:
    t0 = time.time()
    picked = pick_excel_file("选择词曲比对聚焦文件或主报表")
    if picked is None:
        print("已取消。")
        sys.exit(0)

    main_path, focus_path = resolve_paths(picked)
    if not main_path.exists():
        print(f"错误：找不到主报表: {main_path}")
        sys.exit(1)
    if not focus_path.exists():
        print(f"错误：找不到聚焦文件: {focus_path.name}")
        print("  请先运行 15 词曲比对筛选 生成聚焦清单。")
        sys.exit(1)

    print(f"主报表: {main_path.name}")
    print(f"聚焦文件: {focus_path.name}")

    done_groups, pending_groups = load_focus_groups(focus_path)
    print(f"  已完成（无「{PENDING_TEXT}」）: {len(done_groups)} 个分组")
    print(f"  仍待处理: {len(pending_groups)} 个分组")

    if not done_groups:
        print()
        print("没有可回写的分组。")
        print(f"  请在聚焦文件 E 列删除已看完分组的「{PENDING_TEXT}」后再运行。")
        sys.exit(0)

    if not confirm_sync(done_groups, main_path, focus_path):
        print("已取消。")
        sys.exit(0)

    row_start, row_end = confirm_process_file("16 词曲比对聚焦回写", main_path, data_start_row=DATA_START_ROW)

    assert_file_writable(main_path)

    backup_path = backup_file(main_path)
    print(f"  已备份: {backup_path.name}")

    wb = load_workbook(main_path)
    ws = wb.worksheets[SHEET_INDEX]
    header_map = read_header_map(ws, HEADER_ROW)
    cmp_col = header_map.get(LYRICS_CMP_HEADER)
    if cmp_col is None:
        wb.close()
        print(f"错误：第 {HEADER_ROW} 行找不到列「{LYRICS_CMP_HEADER}」")
        sys.exit(1)

    done_keys = {(g.artist, g.album) for g in done_groups}
    updated = apply_completed_groups(ws, row_start, row_end, cmp_col, done_keys)
    saved_path = save_workbook_with_fallback(wb, main_path)
    wb.close()

    print()
    print("=" * 50)
    print("聚焦回写完成")
    print("=" * 50)
    print(f"  回写分组: {len(done_groups)} 个")
    print(f"  主报表已清空 false: {updated} 行")
    print(f"  仍待处理: {len(pending_groups)} 个分组（聚焦文件未改）")
    print(f"  已保存: {saved_path}")
    print(f"  备份: {backup_path.name}")
    print(f"  总耗时: {time.time() - t0:.1f} 秒")
    print()
    print("提示：聚焦文件中仍含「待处理」的分组未动；全部完成后可再次运行本脚本。")


if __name__ == "__main__":
    from launcher_common import run_with_pause

    run_with_pause(run_sync)
