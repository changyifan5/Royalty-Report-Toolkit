"""弹窗选择 Excel / 文件夹路径，供各脚本共用。"""

from __future__ import annotations

import os
import sys
import tkinter as tk
from pathlib import Path
from tempfile import NamedTemporaryFile
from tkinter import filedialog

EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm"}
EXCEL_FILETYPES = [("Excel 文件", "*.xlsx;*.xlsm;*.xltx;*.xltm"), ("所有文件", "*.*")]
DEFAULT_DIRECTORY_FILE = Path(__file__).resolve().with_name(".toolbox-default-directory")


def get_initial_dir() -> Path:
    try:
        saved = Path(DEFAULT_DIRECTORY_FILE.read_text(encoding="utf-8"))
        if saved.is_absolute() and saved.is_dir():
            return saved
    except (OSError, ValueError):
        pass

    profile = Path.home()
    for name in ("Desktop", "桌面"):
        candidate = profile / name
        if candidate.is_dir():
            return candidate
    userprofile = os.environ.get("USERPROFILE", "")
    if userprofile:
        for name in ("Desktop", "桌面"):
            candidate = Path(userprofile) / name
            if candidate.is_dir():
                return candidate
    return profile


def set_default_directory(folder: Path) -> None:
    folder = folder.expanduser().resolve()
    if not folder.is_dir():
        raise NotADirectoryError(str(folder))

    with NamedTemporaryFile(
        mode="w", encoding="utf-8", dir=DEFAULT_DIRECTORY_FILE.parent,
        prefix=f"{DEFAULT_DIRECTORY_FILE.name}.", delete=False,
    ) as stream:
        temporary = Path(stream.name)
        try:
            stream.write(str(folder))
            stream.close()
            temporary.replace(DEFAULT_DIRECTORY_FILE)
        finally:
            temporary.unlink(missing_ok=True)


def pick_excel_file(title: str, *, initial_dir: Path | None = None) -> Path | None:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askopenfilename(
        title=title,
        initialdir=str(initial_dir or get_initial_dir()),
        filetypes=EXCEL_FILETYPES,
    )
    root.destroy()
    if not path:
        return None
    picked = Path(path)
    if picked.suffix.lower() not in EXCEL_SUFFIXES:
        return None
    return picked


def pick_excel_files(title: str, *, initial_dir: Path | None = None) -> list[Path]:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    paths = filedialog.askopenfilenames(
        title=title,
        initialdir=str(initial_dir or get_initial_dir()),
        filetypes=EXCEL_FILETYPES,
    )
    root.destroy()
    result: list[Path] = []
    for raw in paths:
        path = Path(raw)
        if path.suffix.lower() in EXCEL_SUFFIXES:
            result.append(path)
    return result


def pick_folder(title: str, *, initial_dir: Path | None = None) -> Path | None:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    folder = filedialog.askdirectory(title=title, initialdir=str(initial_dir or get_initial_dir()))
    root.destroy()
    return Path(folder) if folder else None


def require_excel_file(title: str, *, initial_dir: Path | None = None) -> Path:
    path = pick_excel_file(title, initial_dir=initial_dir)
    if path is None:
        print("已取消。")
        sys.exit(0)
    return path


def get_data_row_range(
    path: Path,
    *,
    data_start_row: int = 7,
    sheet_index: int = 0,
    key_col: int = 1,
) -> tuple[int, int, int]:
    """返回 (data_start_row, data_end_row, total_data_rows)。"""
    from openpyxl import load_workbook

    from row_order_common import find_last_data_row

    wb = load_workbook(path, read_only=True, keep_links=False)
    if sheet_index >= len(wb.worksheets):
        wb.close()
        print("错误：工作簿中没有可用的工作表")
        sys.exit(1)
    ws = wb.worksheets[sheet_index]
    last_row = find_last_data_row(ws, data_start_row=data_start_row, key_col=key_col)
    wb.close()
    if last_row < data_start_row:
        print(f"错误：第 {data_start_row} 行起没有数据（A 列为空）。")
        sys.exit(1)
    total = last_row - data_start_row + 1
    return data_start_row, last_row, total


def confirm_process_file(
    script_name: str,
    path: Path,
    *,
    data_start_row: int = 7,
    sheet_index: int = 0,
    key_col: int = 1,
) -> tuple[int, int]:
    """统计数据行并返回 (data_start_row, data_end_row)。"""
    print(f"  [{script_name}] 统计数据行数: {path.name} ...")
    row_start, row_end, total = get_data_row_range(
        path,
        data_start_row=data_start_row,
        sheet_index=sheet_index,
        key_col=key_col,
    )
    print(f"  数据行: {total:,}（第 {row_start}～{row_end} 行）")
    return row_start, row_end
