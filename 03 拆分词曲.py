"""
拆分词曲脚本（代码库 03）

独立工具：弹窗选择 Excel，拆分完成后写回原文件（先备份 .bak）。

对第一个页签：
  1. 复制页签，副本命名为「词曲」，原页签改名为「录音」
  2. 第 7 行起：将「录音」H 列的值写入「词曲」G 列
  3. 第 7 行起：将「录音」「词曲」的 H 列均设为数值 0

若 H 列（第 7 行起）已全部为 0，或已存在「录音」「词曲」页签，则跳过（视为已处理过）。

使用 Excel COM（Windows，需 03_split_com.ps1）；Mac 使用 openpyxl。

运行前请关闭 Excel 中已打开的文件。
运行：python "03 拆分词曲.py"
"""

from __future__ import annotations

import shutil
import subprocess
import sys
import time
from pathlib import Path

from openpyxl import load_workbook

BASE_DIR = Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from excel_sheet_copy import copy_sheet_within_workbook
from file_picker_common import require_excel_file
from row_order_common import find_last_data_row

PS1_SCRIPT = BASE_DIR / "03_split_com.ps1"
SHEET_RECORDING = "录音"
SHEET_LYRICS = "词曲"
DATA_START_ROW = 7
COL_G = 7
COL_H = 8

EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm"}


def safe_print(text: str) -> None:
    try:
        print(text)
    except UnicodeEncodeError:
        print(text.encode(sys.stdout.encoding or "utf-8", errors="replace").decode(sys.stdout.encoding or "utf-8", errors="replace"))


def run_com_split(path: Path) -> str:
    if not PS1_SCRIPT.exists():
        print(f"错误：找不到 COM 脚本: {PS1_SCRIPT}")
        sys.exit(1)

    log_file = BASE_DIR / "03_com_error.log"
    cmd = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(PS1_SCRIPT),
        "-Path",
        str(path.resolve()),
        "-SheetRecording",
        SHEET_RECORDING,
        "-SheetLyrics",
        SHEET_LYRICS,
        "-DataStartRow",
        str(DATA_START_ROW),
    ]

    print("  Excel COM：复制页签 + 批量列操作...")
    proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, encoding="utf-8", errors="replace")
    log_file.write_text(proc.stdout or "", encoding="utf-8")
    for line in (proc.stdout or "").splitlines():
        safe_print(f"  {line}")

    if proc.returncode != 0:
        print()
        print("错误：COM 执行失败，日志:", log_file)
        sys.exit(1)

    return proc.stdout or ""


def cell_is_zero(value: object) -> bool:
    if value is None:
        return True
    if isinstance(value, (int, float)):
        return value == 0
    text = str(value).strip()
    return text in ("", "0", "0.0")


def h_column_all_zero(ws, data_start_row: int) -> bool:
    last_row = find_last_data_row(ws, data_start_row)
    if last_row < data_start_row:
        return False
    for row in range(data_start_row, last_row + 1):
        if not cell_is_zero(ws.cell(row=row, column=COL_H).value):
            return False
    return True


def run_openpyxl_split(path: Path) -> str:
    print("  openpyxl：复制页签 + 列操作（Mac 模式，较慢）...")
    t0 = time.time()

    wb = load_workbook(path, keep_links=False)
    names = wb.sheetnames
    if SHEET_RECORDING in names and SHEET_LYRICS in names:
        orig = wb.sheetnames[0]
        wb.close()
        safe_print(f"  SKIP,already_split,{orig}")
        return f"SKIP,already_split,{orig}"

    ws0 = wb.worksheets[0]
    orig_name = ws0.title

    if h_column_all_zero(ws0, DATA_START_ROW):
        wb.close()
        safe_print(f"  SKIP,h_all_zero,{orig_name}")
        return f"SKIP,h_all_zero,{orig_name}"

    lyrics_ws = copy_sheet_within_workbook(wb, ws0)
    ws0.title = SHEET_RECORDING
    lyrics_ws.title = SHEET_LYRICS

    recording = wb[SHEET_RECORDING]
    lyrics = wb[SHEET_LYRICS]
    last_row = find_last_data_row(recording, DATA_START_ROW)
    data_rows = max(0, last_row - DATA_START_ROW + 1)

    # 行块批量：先一次性读出本段「录音」H 列原值，再批量写「词曲」G 列、并把两表 H 列清零。
    # 本循环只搬运/改写数值，不涉及样式、列宽、行高、合并单元格——页签整体复制已由
    # copy_sheet_within_workbook(=copy_worksheet) 一次性完成；语义与原逐行写法一致。
    if last_row >= DATA_START_ROW:
        h_values = [
            row_cells[0].value
            for row_cells in recording.iter_rows(
                min_row=DATA_START_ROW, max_row=last_row, min_col=COL_H, max_col=COL_H
            )
        ]
        for offset, row in enumerate(range(DATA_START_ROW, last_row + 1)):
            lyrics.cell(row=row, column=COL_G).value = h_values[offset]
            recording.cell(row=row, column=COL_H).value = 0
            lyrics.cell(row=row, column=COL_H).value = 0

    wb.save(path)
    wb.close()
    safe_print(f"  openpyxl 耗时: {time.time() - t0:.1f}s")
    safe_print(f"  STATS,{orig_name},{SHEET_RECORDING},{SHEET_LYRICS},{data_rows}")
    return f"STATS,{orig_name},{SHEET_RECORDING},{SHEET_LYRICS},{data_rows}"


def run_split() -> None:
    t0 = time.time()
    path_in = require_excel_file("选择要拆分词曲的 Excel")

    print(f"处理: {path_in.name}")

    if sys.platform == "win32" and PS1_SCRIPT.exists():
        com_output = run_com_split(path_in)
    else:
        com_output = run_openpyxl_split(path_in)

    orig_name = ""
    rows = 0
    skipped = False
    skip_reason = ""

    for line in com_output.splitlines():
        if line.startswith("STATS,"):
            parts = line.split(",")
            orig_name, _, _, rows_str = parts[1], parts[2], parts[3], parts[4]
            rows = int(rows_str)
        elif line.startswith("SKIP,"):
            parts = line.split(",", 2)
            skip_reason = parts[1] if len(parts) > 1 else "unknown"
            orig_name = parts[2] if len(parts) > 2 else ""
            skipped = True

    print()
    print("=" * 50)
    if skipped:
        reason_text = {
            "h_all_zero": "H 列已全部为 0，视为已拆分过",
            "already_split": "已存在「录音」「词曲」页签，视为已拆分过",
        }.get(skip_reason, "已处理过，无需重复操作")
        print("跳过（无需操作）")
        print("=" * 50)
        if orig_name:
            print(f"  文件: {path_in.name}")
            print(f"  页签: {orig_name!r}")
        print(f"  原因: {reason_text}")
        print(f"  文件未修改: {path_in}")
        print(f"  总耗时: {time.time() - t0:.1f} 秒")
        return

    print("拆分完成")
    print("=" * 50)
    print(f"  原第一页签: {orig_name!r} → {SHEET_RECORDING!r}")
    print(f"  复制页签: {SHEET_LYRICS!r}")
    print(f"  数据行: 第 {DATA_START_ROW} 行起，共 {rows} 行")
    print(f"  「{SHEET_LYRICS}」G 列 ← 「{SHEET_RECORDING}」H 列")
    print(f"  「{SHEET_RECORDING}」「{SHEET_LYRICS}」H 列已设为 0")
    print(f"  已保存: {path_in}")
    print(f"  总耗时: {time.time() - t0:.1f} 秒")


if __name__ == "__main__":
    run_split()
