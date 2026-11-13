"""05 建档行序 / 05 恢复行序 共用工具。"""

from __future__ import annotations

import sys
from copy import copy
from pathlib import Path

from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent
BASELINE_DIR = BASE_DIR / "00-原始基准"

SEQ_HEADER = "__原序"
HEADER_ROW = 6
DATA_START_ROW = 7
HEADER_SCAN_LIMIT = 500
KEY_COL = 1

EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm"}
FILL_YELLOW = PatternFill("solid", fgColor="FFFF00")


def resolve_excel_path(path: Path, label: str) -> Path:
    if not path.exists():
        print(f"错误：{label} 路径不存在: {path}")
        sys.exit(1)

    if path.is_file():
        if path.suffix.lower() not in EXCEL_SUFFIXES:
            print(f"错误：{label} 不是支持的 Excel 格式: {path}")
            sys.exit(1)
        return path

    candidates = sorted(
        f
        for f in path.iterdir()
        if f.is_file() and f.suffix.lower() in EXCEL_SUFFIXES and not f.name.startswith("~$")
    )
    if not candidates:
        print(f"错误：{label} 文件夹内没有找到 Excel 文件: {path}")
        sys.exit(1)
    if len(candidates) > 1:
        print(f"错误：{label} 文件夹内有多个 Excel 文件，请只保留一个:")
        for f in candidates:
            print(f"  - {f}")
        sys.exit(1)

    print(f"  已自动选用 {label}: {candidates[0].name}")
    return candidates[0]


def normalize_key(value: object) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def find_header_last_col(sheet, header_row: int = HEADER_ROW, limit: int = HEADER_SCAN_LIMIT) -> int:
    last = 0
    for col in range(1, limit + 1):
        val = sheet.cell(row=header_row, column=col).value
        if val is not None and str(val).strip():
            last = col
    return last


def read_header_map(sheet, header_row: int = HEADER_ROW) -> dict[str, int]:
    mapping: dict[str, int] = {}
    for col in range(1, HEADER_SCAN_LIMIT + 1):
        val = sheet.cell(row=header_row, column=col).value
        if val is not None and str(val).strip():
            mapping[str(val).strip()] = col
    return mapping


def find_last_data_row(sheet, data_start_row: int = DATA_START_ROW, key_col: int = KEY_COL) -> int:
    if getattr(sheet.parent, "read_only", False):
        last = data_start_row - 1
        for idx, row in enumerate(
            sheet.iter_rows(
                min_row=data_start_row,
                min_col=key_col,
                max_col=key_col,
                values_only=True,
            )
        ):
            if normalize_key(row[0] if row else None) is not None:
                last = data_start_row + idx
        return max(last, data_start_row - 1)

    max_row = sheet.max_row
    if max_row < data_start_row:
        return data_start_row - 1
    for row in range(max_row, data_start_row - 1, -1):
        if normalize_key(sheet.cell(row=row, column=key_col).value) is not None:
            return row
    return data_start_row - 1


def parse_seq(value: object) -> int | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if float(value) == int(value):
            return int(value)
        return None
    text = str(value).strip()
    if text.isdigit():
        return int(text)
    return None


def sheet_used_max_col(sheet) -> int:
    return max(sheet.max_column, find_header_last_col(sheet))


def row_has_content(sheet, row: int, max_col: int) -> bool:
    for col in range(1, max_col + 1):
        val = sheet.cell(row=row, column=col).value
        if val is not None and str(val).strip():
            return True
    return False


def copy_cell_style(src, dst) -> None:
    if src.has_style:
        dst.font = copy(src.font)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.number_format = copy(src.number_format)
        dst.protection = copy(src.protection)
        dst.alignment = copy(src.alignment)


def read_row(sheet, row: int, max_col: int) -> list:
    cells = []
    for col in range(1, max_col + 1):
        src = sheet.cell(row=row, column=col)
        cells.append((src.value, src))
    return cells


def write_row(sheet, row: int, cells: list) -> None:
    for col, (value, src_cell) in enumerate(cells, start=1):
        dst = sheet.cell(row=row, column=col)
        dst.value = value
        copy_cell_style(src_cell, dst)


def collect_seq_values(sheet, seq_col: int, data_start_row: int, last_row: int) -> set[int]:
    values: set[int] = set()
    for row in range(data_start_row, last_row + 1):
        seq = parse_seq(sheet.cell(row=row, column=seq_col).value)
        if seq is not None:
            values.add(seq)
    return values


def find_seq_column(sheet) -> int | None:
    header_map = read_header_map(sheet)
    return header_map.get(SEQ_HEADER)


def format_seq_col(seq_col: int) -> str:
    return f"{get_column_letter(seq_col)}（{SEQ_HEADER}）"
