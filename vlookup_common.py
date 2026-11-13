"""01 和 02 共用的比对表读取、结果列定位与匹配写入。"""

from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

from excel_common import normalize_key


BASE_DIR = Path(__file__).resolve().parent
FILE_B = BASE_DIR / "01-vlookup-比对表"

DATA_START_ROW_A = 7
DATA_START_ROW_B = 2
LOOKUP_KEY_COL_A = "A"
LOOKUP_KEY_COL_B = "B"
RETURN_COLS_B = ["B", "D", "E", "F"]
HEADER_ROW_A = 6
HEADER_ROW_B = 1
NOT_FOUND_VALUE = "#N/A"

# 长任务进度打印步长（每处理多少行打印一次）。
PROGRESS_STEP = 10_000


def lookup_headers(sheet) -> list[str]:
    headers = []
    for col in RETURN_COLS_B:
        value = sheet[f"{col}{HEADER_ROW_B}"].value
        headers.append(str(value).strip() if value is not None and str(value).strip() else f"match_{col}")
    return headers


def read_lookup_headers(path: Path) -> list[str]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        return lookup_headers(workbook.worksheets[0])
    finally:
        workbook.close()


def load_lookup_table(path_b: Path) -> tuple[list[str], dict[str, tuple[object, ...]]]:
    key_idx = column_index_from_string(LOOKUP_KEY_COL_B)
    ret_idxs = [column_index_from_string(col) for col in RETURN_COLS_B]
    min_col = min([key_idx, *ret_idxs])
    max_col = max([key_idx, *ret_idxs])

    wb = load_workbook(path_b, read_only=True, data_only=True)
    ws = wb.worksheets[0]

    headers = lookup_headers(ws)

    lookup: dict[str, tuple[object, ...]] = {}
    for row in ws.iter_rows(
        min_row=DATA_START_ROW_B,
        min_col=min_col,
        max_col=max_col,
        values_only=True,
    ):
        key = normalize_key(row[key_idx - min_col])
        if key is None:
            continue
        lookup[key] = tuple(row[idx - min_col] for idx in ret_idxs)

    wb.close()
    return headers, lookup


def resolve_result_start_col(sheet, headers: list[str], last_header_col: int) -> int:
    start_col = last_header_col + 1
    num_cols = len(headers)
    # 02 可能已在结果列后追加处理标记。再次筛选运行时复用同一组结果列。
    result_end = last_header_col
    while result_end > 0 and str(sheet.cell(HEADER_ROW_A, result_end).value or "").strip() in {
        "错误类型", "已完成", "__原序",
    }:
        result_end -= 1
    if result_end >= num_cols:
        candidate = result_end - num_cols + 1
        if all(
            str(sheet.cell(row=HEADER_ROW_A, column=candidate + i).value or "").strip() == headers[i]
            for i in range(num_cols)
        ):
            return candidate
    return start_col


def write_lookup_rows(sheet, rows, start_col: int, headers: list[str], lookup: dict) -> int:
    for offset, header in enumerate(headers):
        sheet.cell(HEADER_ROW_A, start_col + offset).value = header
    matched = 0
    key_col = column_index_from_string(LOOKUP_KEY_COL_A)
    not_found = (NOT_FOUND_VALUE,) * len(headers)
    for processed, row in enumerate(rows, start=1):
        key = normalize_key(sheet.cell(row, key_col).value)
        values = lookup.get(key)
        if values is None:
            values = not_found
        else:
            matched += 1
        for offset, value in enumerate(values):
            sheet.cell(row, start_col + offset).value = value
        if processed % PROGRESS_STEP == 0:
            print(f"  进度: {processed:,}/{len(rows):,} 行，已匹配 {matched:,}")
    return matched
