"""Windows 使用 Excel COM；Mac 使用 openpyxl。"""

from __future__ import annotations

import json
import subprocess
import sys
import tkinter as tk
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
from tkinter import messagebox, simpledialog

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter

from file_picker_common import pick_excel_file
from launcher_common import run_with_pause


def summarize_openpyxl(path: Path, column: int, start: int) -> list[tuple[str, str]]:
    values = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        formulas = load_workbook(path, read_only=True, keep_links=False)
        try:
            results = []
            for sheet in values.worksheets:
                options = dict(min_row=start, min_col=column, max_col=column)
                total = Decimal(0)
                errors = 0
                for (cell,), (original,) in zip(
                    sheet.iter_rows(**options), formulas[sheet.title].iter_rows(**options),
                ):
                    if cell.data_type == "e" or (original.data_type == "f" and cell.value is None):
                        errors += 1
                    elif type(cell.value) in (int, float):
                        total += Decimal(str(cell.value))
                detail = format(total, "f")
                if errors:
                    detail = f"无法汇总（{errors} 个错误值或公式缺少计算结果）"
                results.append((sheet.title, detail))
            return results
        finally:
            formulas.close()
    finally:
        values.close()


def summarize_com(path: Path, column: int, start: int) -> list[tuple[str, str]]:
    with TemporaryDirectory() as directory:
        manifest = Path(directory) / "manifest.json"
        result = Path(directory) / "result.json"
        manifest.write_text(json.dumps({
            "source": str(path.resolve()), "column": column, "start": start,
            "result": str(result),
        }), encoding="utf-8")
        subprocess.run([
            "powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
            str(Path(__file__).resolve().with_name("23_column_sum_com.ps1")), str(manifest),
        ], check=True)
        items = json.loads(result.read_text(encoding="utf-8-sig"))
        return [
            (item["sheet"], format(Decimal(item["total"]), "f")
             if item["total"] is not None else "无法汇总（所选列有错误值）")
            for item in items
        ]


def main() -> None:
    path = pick_excel_file("选择要逐页签加总的 Excel 文件")
    if path is None:
        return

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        while True:
            entered = simpledialog.askstring("选择列", "输入要加总的列字母，例如 G 或 AA。", parent=root)
            if entered is None:
                return
            try:
                column = column_index_from_string(entered.strip().upper())
                if 1 <= column <= 16384:
                    break
            except ValueError:
                pass
            messagebox.showerror("列字母无效", "请输入 A 到 XFD 之间的列字母。", parent=root)
        start = simpledialog.askinteger(
            "起始行", "从第几行加总到页签末尾？",
            initialvalue=2, minvalue=1, maxvalue=1048576, parent=root,
        )
        if start is None:
            return
    finally:
        root.destroy()

    summarize = summarize_com if sys.platform == "win32" else summarize_openpyxl
    results = summarize(path, column, start)
    print(f"\n{path.name}  |  {get_column_letter(column)} 列，从第 {start} 行起")
    print("页签\t合计")
    for title, total in results:
        print(f"{title}\t{total}")


if __name__ == "__main__":
    run_with_pause(main)
