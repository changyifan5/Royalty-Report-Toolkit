"""
表头检查脚本（代码库 12）

对比待检查文件与表头标准文件（通常是中介文件 XXX平台XXX类型报表-时间.xlsx）的
前 5 行表头区是否一致（含公式 vs 被改成数值等情况），
第 6 行不参与对比与恢复。
不一致时可从标准文件恢复表头（公式、样式、列宽、行高、合并单元格）。

流程：
  1. 弹窗选择表头标准文件（中介）
  2. 弹窗选择待检查的 Excel
  3. 对比第 1～5 行，列出差异
  4. 确认后写回：用标准文件表头覆盖待检查文件表头（先备份 .bak）

仅处理第一个页签。Mac / Windows 均使用 openpyxl（表头区仅 5 行，无需 COM）。
运行前请关闭 Excel。
"""

from __future__ import annotations

import shutil
import sys
import time
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import messagebox

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from excel_sheet_copy import restore_header_region
from file_picker_common import get_initial_dir, pick_excel_file, require_excel_file
from intermediate_common import INTERMEDIATE_NAME_TEMPLATE

BASE_DIR = Path(__file__).resolve().parent
SHEET_INDEX = 0
HEADER_ROWS = 5
HEADER_SCAN_LIMIT = 500
DIFF_DISPLAY_LIMIT = 80


@dataclass(frozen=True)
class HeaderDiff:
    row: int
    col: int
    ref_desc: str
    tgt_desc: str


def safe_print(text: str) -> None:
    try:
        print(text)
    except UnicodeEncodeError:
        print(text.encode(sys.stdout.encoding or "utf-8", errors="replace").decode(
            sys.stdout.encoding or "utf-8", errors="replace"
        ))


def resolve_reference_if_same_as_target(reference: Path, target: Path) -> Path | None:
    if reference.resolve() != target.resolve():
        return reference

    print("  待检查文件与所选标准为同一文件。")
    bak = target.with_suffix(target.suffix + ".bak")
    if bak.exists():
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        use_bak = messagebox.askyesno(
            "使用备份作标准",
            f"是否改用备份文件作为表头标准？\n  {bak.name}",
            parent=root,
        )
        root.destroy()
        if use_bak:
            print(f"  已改用备份作标准: {bak.name}")
            return bak

    print("  请另选一个未损坏的 Excel 作为表头标准（如其他平台中介模板）。")
    picked = pick_excel_file(
        "选择表头标准文件",
        initial_dir=get_initial_dir(),
    )
    return picked


def header_region_last_col(sheet, header_rows: int) -> int:
    last = 0
    for row in range(1, header_rows + 1):
        for col in range(1, HEADER_SCAN_LIMIT + 1):
            cell = sheet.cell(row=row, column=col)
            if cell.data_type == "f":
                last = col
            elif cell.value is not None and str(cell.value).strip():
                last = col
    return last


def describe_cell(cell) -> str:
    if cell.data_type == "f" or (isinstance(cell.value, str) and str(cell.value).startswith("=")):
        return f"公式 {cell.value}"
    if cell.value is None or str(cell.value).strip() == "":
        return "（空）"
    return f"数值/文本 {cell.value!r}"


def cells_match(ref_cell, tgt_cell) -> bool:
    ref_is_formula = ref_cell.data_type == "f" or (
        isinstance(ref_cell.value, str) and str(ref_cell.value).startswith("=")
    )
    tgt_is_formula = tgt_cell.data_type == "f" or (
        isinstance(tgt_cell.value, str) and str(tgt_cell.value).startswith("=")
    )
    if ref_is_formula or tgt_is_formula:
        if ref_is_formula != tgt_is_formula:
            return False
        return str(ref_cell.value or "").strip() == str(tgt_cell.value or "").strip()

    ref_empty = ref_cell.value is None or str(ref_cell.value).strip() == ""
    tgt_empty = tgt_cell.value is None or str(tgt_cell.value).strip() == ""
    if ref_empty and tgt_empty:
        return True
    if ref_empty != tgt_empty:
        return False
    return ref_cell.value == tgt_cell.value or str(ref_cell.value) == str(tgt_cell.value)


def compare_header_region(ref_ws, tgt_ws, header_rows: int, last_col: int) -> list[HeaderDiff]:
    diffs: list[HeaderDiff] = []
    for row in range(1, header_rows + 1):
        for col in range(1, last_col + 1):
            ref_cell = ref_ws.cell(row=row, column=col)
            tgt_cell = tgt_ws.cell(row=row, column=col)
            if cells_match(ref_cell, tgt_cell):
                continue
            diffs.append(
                HeaderDiff(
                    row=row,
                    col=col,
                    ref_desc=describe_cell(ref_cell),
                    tgt_desc=describe_cell(tgt_cell),
                )
            )
    return diffs


def backup_file(path: Path) -> Path:
    backup_path = path.with_suffix(path.suffix + ".bak")
    if backup_path.exists():
        backup_path.unlink()
    shutil.copy2(path, backup_path)
    return backup_path


def ask_restore(count: int, target_name: str) -> bool:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    ok = messagebox.askyesno(
        "恢复表头",
        f"文件「{target_name}」前 {HEADER_ROWS} 行与标准有 {count} 处差异。\n\n"
        f"是否用标准文件的表头覆盖？\n（会先备份为 .bak，仅改第 1～{HEADER_ROWS} 行，第 6 行起不动）",
        parent=root,
    )
    root.destroy()
    return ok


def run_check() -> None:
    t0 = time.time()
    reference = require_excel_file(f"选择表头标准文件（如 {INTERMEDIATE_NAME_TEMPLATE}）")
    safe_print(f"  表头标准: {reference.name}")

    target = require_excel_file("选择待检查的 Excel")
    safe_print(f"  待检查文件: {target.name}")

    reference = resolve_reference_if_same_as_target(reference, target)
    if reference is None:
        print("已取消。")
        sys.exit(0)
    if reference.resolve() != target.resolve():
        safe_print(f"  表头标准: {reference.name}")

    ref_wb = load_workbook(reference, read_only=False, keep_links=False, data_only=False)
    tgt_wb = load_workbook(target, read_only=False, keep_links=False, data_only=False)
    if len(ref_wb.worksheets) <= SHEET_INDEX or len(tgt_wb.worksheets) <= SHEET_INDEX:
        ref_wb.close()
        tgt_wb.close()
        print("错误：工作簿中没有可用的工作表。")
        sys.exit(1)

    ref_ws = ref_wb.worksheets[SHEET_INDEX]
    tgt_ws = tgt_wb.worksheets[SHEET_INDEX]
    ref_sheet_name = ref_ws.title
    tgt_sheet_name = tgt_ws.title
    last_col = max(header_region_last_col(ref_ws, HEADER_ROWS), header_region_last_col(tgt_ws, HEADER_ROWS), 1)

    diffs = compare_header_region(ref_ws, tgt_ws, HEADER_ROWS, last_col)
    ref_wb.close()

    print()
    print("=" * 50)
    print("表头检查结果")
    print("=" * 50)
    print(f"  对比范围: 第 1～{HEADER_ROWS} 行，A～{get_column_letter(last_col)} 列")
    print(f"  标准: {reference.name} / {ref_sheet_name}")
    print(f"  文件: {target.name} / {tgt_sheet_name}")

    if not diffs:
        tgt_wb.close()
        print("  结果: 表头一致，无需修改。")
        print(f"  总耗时: {time.time() - t0:.1f} 秒")
        return

    print(f"  结果: 发现 {len(diffs)} 处差异")
    shown = diffs[:DIFF_DISPLAY_LIMIT]
    for diff in shown:
        coord = f"{get_column_letter(diff.col)}{diff.row}"
        safe_print(f"    {coord}  标准={diff.ref_desc}  |  当前={diff.tgt_desc}")
    if len(diffs) > DIFF_DISPLAY_LIMIT:
        print(f"    … 还有 {len(diffs) - DIFF_DISPLAY_LIMIT} 处未显示")

    if not ask_restore(len(diffs), target.name):
        tgt_wb.close()
        print("已取消：文件未修改。")
        return

    ref_wb = load_workbook(reference, read_only=False, keep_links=False, data_only=False)
    ref_ws = ref_wb.worksheets[SHEET_INDEX]
    tgt_ws = tgt_wb.worksheets[SHEET_INDEX]

    backup_path = backup_file(target)
    print(f"  已备份: {backup_path.name}")
    restore_header_region(ref_ws, tgt_ws, HEADER_ROWS, last_col)
    tgt_wb.save(target)
    ref_wb.close()
    tgt_wb.close()

    print()
    print("=" * 50)
    print("表头已恢复")
    print("=" * 50)
    print(f"  已从「{reference.name}」覆盖第 1～{HEADER_ROWS} 行")
    print(f"  已保存: {target}")
    print(f"  备份: {backup_path}")
    print(f"  总耗时: {time.time() - t0:.1f} 秒")


if __name__ == "__main__":
    run_check()
