"""用待 v 文件 A 的 A 列匹配 B 的 A 列，在 B 的命中行末尾写「已对应」或回填指定列。"""

import hashlib
import tkinter as tk
from pathlib import Path
from tempfile import TemporaryDirectory
from tkinter import messagebox, simpledialog

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import column_index_from_string

from excel_common import backup_file, list_sheet_names_fast, normalize_key
from file_picker_common import require_excel_file
from launcher_common import run_with_pause
from report_chunk import (
    finalize_chunk_output,
    load_file_chunks,
    prepare_chunk_work,
    prompt_next_chunk_after,
    prompt_row_chunk,
)

TITLE = "04 已有的修改之v"
SCRIPT_KEY = "04-match"
START_ROW = 7
MARK = "已对应"


def pick_sheet(path: Path, label: str) -> str | None:
    workbook = load_workbook(path, read_only=True)
    try:
        names = [sheet.title for sheet in workbook.worksheets]
    finally:
        workbook.close()
    if not names:
        raise ValueError(f"{path.name} 没有可用的工作表。")
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        options = "\n".join(f"{i}. {name}" for i, name in enumerate(names, start=1))
        number = simpledialog.askinteger(
            f"选择 {label} 页签", f"{path.name}\n\n{options}\n\n请输入页签编号。",
            initialvalue=1, minvalue=1, maxvalue=len(names), parent=root,
        )
        return names[number - 1] if number is not None else None
    finally:
        root.destroy()


def pick_output_column() -> int | None:
    """0 表示写固定标记，None 表示取消。"""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        while True:
            value = simpledialog.askstring(
                "选择输出内容",
                "留空写「已对应」，或输入 A 文件的列字母（例如 D、AA），回填命中行该列的值。",
                parent=root,
            )
            if value is None:
                return None
            letter = value.strip().upper()
            if not letter:
                return 0
            try:
                column = column_index_from_string(letter)
                if column <= 16384:
                    return column
            except ValueError:
                pass
            messagebox.showwarning("列字母无效", "请输入 A 到 XFD 之间的列字母。", parent=root)
    finally:
        root.destroy()


def load_lookup(path: Path, sheet_name: str, output_column: int) -> dict:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        lookup = {}
        for row in workbook[sheet_name].iter_rows(max_col=max(1, output_column)):
            key = normalize_key(row[0].value)
            if key is not None and key not in lookup:
                if output_column:
                    cell = row[output_column - 1]
                    lookup[key] = (cell.value, cell.data_type)
                else:
                    lookup[key] = (MARK, "s")
        return lookup
    finally:
        workbook.close()


def mark_rows(path: Path, lookup: dict, start: int, end: int, sheet_name: str, *, marker_only: bool) -> int:
    workbook = load_workbook(path, keep_vba=path.suffix.lower() in {".xlsm", ".xltm"})
    try:
        sheet = workbook[sheet_name]
        hidden = {row for row, dimension in sheet.row_dimensions.items() if dimension.hidden}
        last_cells = {}
        for (row, col), cell in sheet._cells.items():
            if start <= row <= end and row not in hidden and cell.value is not None and cell.value != "":
                if row not in last_cells or col > last_cells[row].column:
                    last_cells[row] = cell

        added = 0
        for row, last in last_cells.items():
            match = lookup.get(normalize_key(sheet.cell(row, 1).value))
            if match is None or match[0] in (None, ""):
                continue
            value, data_type = match
            if marker_only and last.column > 1 and last.value == MARK:
                continue
            col = last.column + 1
            while col <= 16384 and isinstance(sheet.cell(row, col), MergedCell):
                col += 1
            if col > 16384:
                raise ValueError(f"第 {row} 行的行末没有可写的空白格。")
            cell = sheet.cell(row, col, value)
            cell.data_type = data_type
            added += 1

        if added:
            with TemporaryDirectory(dir=path.parent) as directory:
                output = Path(directory) / path.name
                workbook.save(output)
                output.replace(path)
        return added
    finally:
        workbook.close()


def main() -> None:
    path_a = require_excel_file("选择 A：待 v 文件（从 A 列查找）")
    sheet_a = pick_sheet(path_a, "A（待 v 文件）")
    if sheet_a is None:
        return
    path_b = require_excel_file("选择 B：写入匹配结果（请先保存筛选，只处理可见行）")
    if path_a.samefile(path_b):
        raise ValueError("请选择两份不同的文件作为 A 和 B。")

    sheet_b = pick_sheet(path_b, "B（写入标记）")
    if sheet_b is None:
        return
    output_column = pick_output_column()
    if output_column is None:
        return
    lookup = load_lookup(path_a, sheet_a, output_column)
    # 续跑必须沿用同一组页签和匹配键，避免把两次操作写进同一工作副本。
    scope_data = (output_column, sorted(lookup.items())) if output_column else sorted(lookup)
    scope = hashlib.sha256(repr((str(path_a.resolve()), sheet_a, sheet_b, scope_data)).encode()).hexdigest()
    script_key = f"{SCRIPT_KEY}-{scope[:16]}"
    chunks = load_file_chunks(path_b, START_ROW, sheet_index=list_sheet_names_fast(path_b).index(sheet_b))
    chunk = prompt_row_chunk(TITLE, path_b, START_ROW, script_key=script_key, chunks=chunks)
    if chunk is None:
        return
    while chunk is not None:
        work, _ = prepare_chunk_work(path_b, script_key, chunk, backup_file)
        added = mark_rows(work, lookup, chunk.data_start_row, chunk.data_end_row, sheet_b, marker_only=not output_column)
        print(f"本段新增 {added:,} 个匹配结果。")
        finished, output = finalize_chunk_output(path_b, script_key, chunk)
        if finished:
            print(f"已处理 B 文件：{output}")
            return
        chunk = prompt_next_chunk_after(TITLE, path_b, script_key, chunks, chunk, finished)


if __name__ == "__main__":
    run_with_pause(main)
