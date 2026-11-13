"""Windows 使用 Excel COM；Mac 使用 openpyxl。

按同名页签合并文件夹中的 Excel，也可只合并各文件的首张工作表。
"""

from __future__ import annotations

import json
import subprocess
import sys
import time
import tkinter as tk
from pathlib import Path
from tempfile import TemporaryDirectory
from tkinter import filedialog, messagebox, simpledialog

from openpyxl import load_workbook

from excel_sheet_copy import copy_row_range, copy_sheet_between_workbooks
from file_picker_common import EXCEL_SUFFIXES, get_initial_dir, pick_folder
from row_order_common import DATA_START_ROW

MERGE_SUFFIX = "_合并"
SKIP_STEM_PARTS = (".列映射.work", ".bak")


def list_excel_files(folder: Path, *, exclude: Path | None = None) -> list[Path]:
    results: list[Path] = []
    exclude_resolved = exclude.resolve() if exclude else None
    for path in sorted(folder.iterdir(), key=lambda p: p.name.lower()):
        if not path.is_file():
            continue
        if path.suffix.lower() not in EXCEL_SUFFIXES:
            continue
        if path.name.startswith("~$"):
            continue
        if exclude_resolved and path.resolve() == exclude_resolved:
            continue
        stem_lower = path.stem.lower()
        if any(part in stem_lower for part in SKIP_STEM_PARTS):
            continue
        if MERGE_SUFFIX in stem_lower and not stem_lower.endswith("_tme合并"):
            continue
        results.append(path)
    return results


def pick_save_path(default: Path) -> Path | None:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.asksaveasfilename(
        title="保存合并结果",
        initialdir=str(default.parent),
        initialfile=default.name,
        defaultextension=".xlsx",
        filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
    )
    root.destroy()
    return Path(path) if path else None


def confirm_merge(folder: Path, files: list[Path], out_path: Path, start_row: int, first_only: bool) -> bool:
    preview = "\n".join(f"  {i + 1}. {p.name}" for i, p in enumerate(files[:15]))
    if len(files) > 15:
        preview += f"\n  … 共 {len(files)} 个文件"

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    mode = "只合并第一个页签" if first_only else "按同名页签合并全部"
    ok = messagebox.askyesno(
        "17 文件夹 Excel 合并",
        f"文件夹: {folder.name}\n"
        f"{mode}，共 {len(files)} 个文件（第 {start_row} 行起追加数据）\n\n"
        f"{preview}\n\n"
        f"输出: {out_path.name}\n\n"
        "是否继续？",
        parent=root,
    )
    root.destroy()
    return ok


def pick_merge_options() -> tuple[int, bool] | None:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        first_only = messagebox.askyesnocancel(
            "选择合并方式",
            "只合并每个文件的第一个页签？\n\n"
            "选「是」合并首个页签。\n"
            "选「否」按同名页签合并全部。",
            default=messagebox.NO,
            parent=root,
        )
        if first_only is None:
            return None
        start_row = simpledialog.askinteger(
            "数据起始行", "从第几行开始追加数据？", initialvalue=DATA_START_ROW,
            minvalue=1, maxvalue=1048576, parent=root,
        )
        return (start_row, first_only) if start_row is not None else None
    finally:
        root.destroy()


def data_bounds(sheet) -> tuple[int, int]:
    """取最后一个有内容的行和列，包括 A 列为空的数据。"""
    last_row = last_col = 0
    for (row, col), cell in sheet._cells.items():
        if cell.value is not None and cell.value != "":
            last_row = max(last_row, row)
            last_col = max(last_col, col)
    return last_row, last_col


def merge_openpyxl(files: list[Path], output: Path, start_row: int, first_only: bool) -> tuple[int, int]:
    wb = load_workbook(files[0], keep_links=False)
    try:
        if not wb.worksheets:
            raise ValueError(f"{files[0].name} 中没有可用的工作表")
        wb.template = False
        destinations = {ws.title.casefold(): ws for ws in wb.worksheets}
        initial_sheets = wb.worksheets[:1] if first_only else wb.worksheets
        last_rows = {ws.title: data_bounds(ws)[0] for ws in initial_sheets}
        total_rows = sum(max(0, row - start_row + 1) for row in last_rows.values())
        used_files = 1

        for path in files[1:]:
            src_wb = load_workbook(path, keep_links=False)
            try:
                source_sheets = src_wb.worksheets[:1] if first_only else src_wb.worksheets
                if not source_sheets:
                    continue
                for source in source_sheets:
                    src_last, src_col = data_bounds(source)
                    key = source.title.casefold()
                    if not first_only and key not in destinations:
                        dest = wb.create_sheet(source.title)
                        copy_sheet_between_workbooks(source, dest)
                        destinations[key] = dest
                        last_rows[dest.title] = src_last
                    elif src_last >= start_row:
                        dest = wb.worksheets[0] if first_only else destinations[key]
                        dest_start = max(last_rows[dest.title] + 1, start_row)
                        if dest_start + src_last - start_row > 1048576:
                            raise ValueError(f"页签「{dest.title}」超过 Excel 的行数上限")
                        next_row = copy_row_range(source, dest, start_row, src_last, dest_start, src_col)
                        last_rows[dest.title] = next_row - 1
                    total_rows += max(0, src_last - start_row + 1)
                used_files += 1
            finally:
                src_wb.close()
        wb.save(output)
        return total_rows, used_files
    finally:
        wb.close()


def merge_files(
    files: list[Path], out_path: Path, start_row: int = DATA_START_ROW, first_only: bool = True,
) -> tuple[int, int]:
    """返回数据行数和参与合并的文件数。"""
    if not files:
        return 0, 0
    if not 1 <= start_row <= 1048576:
        raise ValueError("数据起始行超出 Excel 行数范围")
    with TemporaryDirectory(prefix=".merge17-", dir=out_path.parent) as directory:
        temporary = Path(directory)
        output = temporary / "merged.xlsx"
        if sys.platform == "win32":
            manifest = temporary / "manifest.json"
            result = temporary / "result.json"
            manifest.write_text(json.dumps({
                "files": [str(path.resolve()) for path in files],
                "output": str(output.resolve()),
                "result": str(result.resolve()),
                "startRow": start_row,
                "firstOnly": first_only,
            }), encoding="utf-8")
            subprocess.run([
                "powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
                "-File", str(Path(__file__).with_name("17_merge_sheets_com.ps1")), str(manifest),
            ], check=True)
            stats = json.loads(result.read_text(encoding="utf-8-sig"))
            totals = stats["rows"], stats["files"]
        else:
            totals = merge_openpyxl(files, output, start_row, first_only)
        output.replace(out_path)
        return totals


def run_merge() -> None:
    t0 = time.time()
    print("=" * 50)
    print("17 文件夹 Excel 合并")
    print("=" * 50)

    folder = pick_folder("选择要合并的 Excel 文件夹", initial_dir=get_initial_dir())
    if folder is None:
        print("已取消。")
        sys.exit(0)

    default_out = folder / f"{folder.name}{MERGE_SUFFIX}.xlsx"
    files = list_excel_files(folder, exclude=default_out if default_out.exists() else None)
    if len(files) < 2:
        print(f"错误：文件夹内至少需要 2 个 Excel 文件（当前 {len(files)} 个）。")
        sys.exit(1)

    print(f"  文件夹: {folder}")
    print(f"  找到 {len(files)} 个文件")

    out_path = pick_save_path(default_out)
    if out_path is None:
        print("已取消。")
        sys.exit(0)

    files = [f for f in files if f.resolve() != out_path.resolve()]
    if len(files) < 2:
        print("错误：排除输出文件后不足 2 个源文件。")
        sys.exit(1)

    options = pick_merge_options()
    if options is None:
        print("已取消。")
        return
    start_row, first_only = options
    if not confirm_merge(folder, files, out_path, start_row, first_only):
        print("已取消。")
        sys.exit(0)

    print()
    print("  开始合并…")
    total_rows, used_files = merge_files(files, out_path, start_row, first_only)

    print()
    print("=" * 50)
    print("合并完成")
    print("=" * 50)
    print(f"  源文件: {used_files}/{len(files)} 个参与")
    print(f"  数据行: 共 {total_rows:,} 行（均自第 {start_row} 行起计）")
    print(f"  输出: {out_path}")
    print(f"  耗时: {time.time() - t0:.1f} 秒")


if __name__ == "__main__":
    from launcher_common import run_with_pause

    run_with_pause(run_merge)
