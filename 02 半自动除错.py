"""
半自动除错脚本（代码库 02）

缺少 VLOOKUP 结果时，先用 01 的比对表和匹配规则补齐，再除错。

从第 7 行起读取「比對結果」列，按错误类型处理：
  - 找不到演唱人 → 用新增四列中第 3 列覆盖 C 列
  - 找不到歌曲   → 用新增四列中第 2 列（歌曲名称）覆盖 B 列
  - 找不到专辑   → 用新增四列中第 4 列覆盖 D 列
  - 收益时间错误 → 仅识别，不自动改值
  - 标记投母     → 整个数据行标灰底并加粗

新增四列按 01 追加顺序从左往右编号（第 1～4 列），不依赖表头文字匹配。

修改过的 B/C/D 单元格标黄色；末尾追加「已完成」列，已自动改值的行标记为「已完成」。
处理前会在原文件旁备份一份 .bak；结果写回原文件。

运行前请关闭 Excel 中已打开的文件。运行：python "02 半自动除错.py"
"""

from __future__ import annotations

import re
import sys
import time
from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

from excel_common import (
    EXCEL_SUFFIXES,
    append_log,
    backup_file,
    find_header_last_col,
    read_header_map,
    resolve_excel_path,
    show_error_popup,
)
from file_picker_common import require_excel_file
from report_chunk import (
    finalize_chunk_output,
    load_file_chunks,
    prepare_chunk_work,
    print_chunk_banner,
    prompt_next_chunk_after,
    prompt_row_chunk,
    prompt_visible_rows_only,
    select_chunk_rows,
)
from vlookup_common import (
    FILE_B, load_lookup_table, read_lookup_headers, resolve_result_start_col, write_lookup_rows,
)

# ---------------------------------------------------------------------------
# 可配置参数
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
SHEET_INDEX = 0  # 第一个页签（与 01 一致，保持原页签名）
HEADER_ROW = 6
DATA_START_ROW = 7
HEADER_SCAN_LIMIT = 500

# 报表待修正列（B=歌曲, C=演唱者, D=专辑）
COL_SONG = 2
COL_ARTIST = 3
COL_ALBUM = 4

# 01 追加的四列 VLOOKUP 结果：从左往右第 1～4 列（0-based 索引）
VLOOKUP_COL_COUNT = 4
VLOOKUP_IDX_SONG = 1    # 第 2 列 → 歌曲名称
VLOOKUP_IDX_ARTIST = 2  # 第 3 列 → 演唱者
VLOOKUP_IDX_ALBUM = 3   # 第 4 列 → 专辑

CMP_RESULT_HEADER = "比對結果"
ERROR_TYPE_HEADER = "错误类型"
DONE_HEADER = "已完成"
DONE_VALUE = "已完成"

FILL_YELLOW = PatternFill("solid", fgColor="FFFF00")
FILL_PUBLISHER = PatternFill("solid", fgColor="FFC0C0C0")
PUBLISHER_ERROR = "授權人 [示例版权管理有限公司] 沒有參與版稅業務 [詞曲：數位發行]"
PUBLISHER_LABEL = "标记投母"
NOT_FOUND_VALUES = {"#N/A", "#N/A!", "#REF!", ""}

LOG_FILE = BASE_DIR / "02_fix.log"

# 错误类型 → 比對結果 列全文匹配（一条带命名分组的正则，按 alternation 顺序返回首个命中组）
_ERROR_RE = re.compile(
    r"(?P<找不到演唱人>找不到名稱為.+的演唱人或藝人)"
    r"|(?P<找不到歌曲>找不到歌曲名稱為.+的歌曲)"
    r"|(?P<找不到专辑>找不到專輯名稱為.+的專輯)"
    r"|(?P<收益时间错误>收益時間.+不符合專輯上、下架（月底）區間)"
)


def classify_error(text: str) -> str | None:
    """返回首个命中的错误类型名；空文本 / 未命中返回 None。"""
    text = text.strip()
    if text == PUBLISHER_ERROR:
        return PUBLISHER_LABEL
    m = _ERROR_RE.search(text)
    return m.lastgroup if m else None


def available_lookup_path() -> Path | None:
    candidates = [
        path for path in FILE_B.iterdir()
        if path.is_file() and path.suffix.lower() in EXCEL_SUFFIXES and not path.name.startswith("~$")
    ] if FILE_B.is_dir() else []
    return candidates[0] if len(candidates) == 1 else None


def ensure_lookup_results(sheet, rows, headers: list[str], get_lookup) -> list[int]:
    last_col = find_header_last_col(sheet, HEADER_ROW)
    start_col = resolve_result_start_col(sheet, headers, last_col)
    columns = list(range(start_col, start_col + VLOOKUP_COL_COUNT))
    missing_rows = rows if start_col > last_col else [
        row for row in rows if all(sheet.cell(row, col).value is None for col in columns)
    ]
    if start_col > last_col or missing_rows:
        lookup_headers, lookup = get_lookup()
        if lookup_headers != headers:
            start_col = resolve_result_start_col(sheet, lookup_headers, last_col)
            columns = list(range(start_col, start_col + VLOOKUP_COL_COUNT))
            missing_rows = rows
        print(f"  自动补跑 01：匹配 {len(missing_rows):,} 行")
        write_lookup_rows(sheet, missing_rows, start_col, lookup_headers, lookup)
    return columns


def vlookup_header_labels(sheet, cols: list[int]) -> list[str]:
    labels: list[str] = []
    for col in cols:
        val = sheet.cell(row=HEADER_ROW, column=col).value
        labels.append(str(val).strip() if val is not None else "")
    return labels


def is_valid_lookup_value(value) -> bool:
    if value is None:
        return False
    text = str(value).strip()
    return text not in NOT_FOUND_VALUES


def cell_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return str(value).strip()


def ensure_error_type_column(sheet, header_map: dict[str, int]) -> int:
    if ERROR_TYPE_HEADER in header_map:
        return header_map[ERROR_TYPE_HEADER]

    last_col = find_header_last_col(sheet, HEADER_ROW)
    col = last_col + 1
    sheet.cell(row=HEADER_ROW, column=col, value=ERROR_TYPE_HEADER)
    return col


def ensure_done_column(sheet, header_map: dict[str, int], after_col: int) -> int:
    if DONE_HEADER in header_map:
        return header_map[DONE_HEADER]

    col = after_col + 1
    sheet.cell(row=HEADER_ROW, column=col, value=DONE_HEADER)
    return col


def run_fix():
    t0 = time.time()
    path_a = require_excel_file("选择要除错的主报表")
    visible_only = prompt_visible_rows_only("02 半自动除错", path_a, "02")
    if visible_only is None:
        return

    chunks = load_file_chunks(path_a, DATA_START_ROW)
    chunk = prompt_row_chunk("02 半自动除错", path_a, DATA_START_ROW, script_key="02", chunks=chunks)
    if chunk is None:
        print("已取消。")
        sys.exit(0)

    path_b = available_lookup_path()
    headers = read_lookup_headers(path_b) if path_b else ["歌曲 ISRC", "歌曲", "演唱人", "專輯"]
    lookup_cache = None

    def get_lookup():
        nonlocal lookup_cache
        if lookup_cache is None:
            lookup_cache = load_lookup_table(path_b or resolve_excel_path(FILE_B, "比对文件 B"))
        return lookup_cache

    while chunk is not None:
        print_chunk_banner("02 半自动除错", chunk)

        work_path, backup_path = prepare_chunk_work(
            path_a, "02", chunk, backup_file, visible_only=visible_only,
        )
        print(f"读取: {path_a.name}")
        if chunk.is_first:
            print(f"  已备份: {backup_path.name}")
        wb = load_workbook(work_path)
        if len(wb.worksheets) <= SHEET_INDEX:
            print("错误：工作簿中没有可用的工作表")
            show_error_popup("02 半自动除错", "工作簿中没有可用的工作表")
            append_log(["[ERROR] 工作簿中没有可用的工作表"], LOG_FILE)
            sys.exit(1)

        ws = wb.worksheets[SHEET_INDEX]
        sheet_name = ws.title
        header_map = read_header_map(ws, HEADER_ROW)

        if CMP_RESULT_HEADER not in header_map:
            print(f"错误：第 {HEADER_ROW} 行找不到列「{CMP_RESULT_HEADER}」")
            show_error_popup("02 半自动除错", f"第 {HEADER_ROW} 行找不到列「{CMP_RESULT_HEADER}」")
            append_log([f"[ERROR] 第 {HEADER_ROW} 行找不到列「{CMP_RESULT_HEADER}」"], LOG_FILE)
            sys.exit(1)

        cmp_col = header_map[CMP_RESULT_HEADER]
        rows = select_chunk_rows(ws, chunk, visible_only)
        vlookup_cols = ensure_lookup_results(ws, rows, headers, get_lookup)
        vlookup_labels = vlookup_header_labels(ws, vlookup_cols)
        error_type_col = ensure_error_type_column(ws, header_map)
        done_col = ensure_done_column(ws, header_map, error_type_col)
        last_col = ws.max_column

        stats = {
            "找不到演唱人": {"found": 0, "fixed": 0, "skipped": 0},
            "找不到歌曲": {"found": 0, "fixed": 0, "skipped": 0},
            "找不到专辑": {"found": 0, "fixed": 0, "skipped": 0},
            "收益时间错误": {"found": 0, "fixed": 0, "skipped": 0},
            PUBLISHER_LABEL: {"found": 0, "fixed": 0, "skipped": 0},
            "未识别": {"found": 0, "fixed": 0, "skipped": 0},
        }

        # 错误类型 → (报表目标列, 新增四列中的索引 0-based)
        fix_plan = {
            "找不到演唱人": (COL_ARTIST, VLOOKUP_IDX_ARTIST),
            "找不到歌曲": (COL_SONG, VLOOKUP_IDX_SONG),
            "找不到专辑": (COL_ALBUM, VLOOKUP_IDX_ALBUM),
        }

        total_rows = len(rows)
        for processed, row in enumerate(rows, start=1):
            if processed % 10000 == 0:
                print(f"  进度: 已处理 {processed}/{total_rows} 行 ...")

            cmp_val = ws.cell(row=row, column=cmp_col).value
            if cmp_val is None or not str(cmp_val).strip():
                continue

            error_type = classify_error(str(cmp_val))
            if error_type is None:
                error_type = "未识别"

            ws.cell(row=row, column=error_type_col, value=error_type)
            bucket = stats[error_type]
            bucket["found"] += 1

            if error_type == PUBLISHER_LABEL:
                for col in range(1, last_col + 1):
                    cell = ws.cell(row, col)
                    cell.fill = FILL_PUBLISHER
                    if not cell.font.bold:
                        font = copy(cell.font)
                        font.bold = True
                        cell.font = font
                continue

            if error_type not in fix_plan:
                continue

            target_col, vlookup_idx = fix_plan[error_type]
            source_col = vlookup_cols[vlookup_idx]
            new_val = ws.cell(row=row, column=source_col).value

            if not is_valid_lookup_value(new_val):
                bucket["skipped"] += 1
                continue

            old_val = ws.cell(row=row, column=target_col).value
            if cell_text(old_val) == cell_text(new_val):
                bucket["skipped"] += 1
                continue

            ws.cell(row=row, column=target_col, value=new_val)
            ws.cell(row=row, column=target_col).fill = FILL_YELLOW
            ws.cell(row=row, column=done_col, value=DONE_VALUE)
            bucket["fixed"] += 1

        wb.save(work_path)
        wb.close()

        finalized, out_path = finalize_chunk_output(path_a, "02", chunk)

        print()
        print("=" * 50)
        print("全部段完成，半自动除错完成" if finalized else "本段处理完成（尚未输出最终文件）")
        print("=" * 50)
        print(f"  工作表: {sheet_name}")
        print(f"  比對結果列: {get_column_letter(cmp_col)}")
        print(f"  修正目标列: B={COL_SONG}, C={COL_ARTIST}, D={COL_ALBUM}")
        print("  新增四列 VLOOKUP（从左往右）:")
        for i, (col, label) in enumerate(zip(vlookup_cols, vlookup_labels)):
            mark = ""
            if i == VLOOKUP_IDX_SONG:
                mark = " → 改 B 列"
            elif i == VLOOKUP_IDX_ARTIST:
                mark = " → 改 C 列"
            elif i == VLOOKUP_IDX_ALBUM:
                mark = " → 改 D 列"
            print(f"    第{i + 1}列 {get_column_letter(col)}: {label or '(无表头)'}{mark}")
        print(f"  错误类型列: {get_column_letter(error_type_col)}")
        print(f"  已完成列: {get_column_letter(done_col)}")
        print()
        for name, s in stats.items():
            if s["found"] == 0:
                continue
            print(f"  {name}: 识别 {s['found']} 行", end="")
            if name in fix_plan:
                print(f"，已修正 {s['fixed']} 行，跳过 {s['skipped']} 行")
            else:
                print()
        print()
        if finalized and out_path:
            print(f"  已保存: {out_path}")
        print(f"  备份文件: {backup_path}")
        print(f"  总耗时: {time.time() - t0:.1f} 秒")
        if finalized:
            print()
            print("提示：黄色单元格为已自动替换的 B/C/D 列，请打开原文件复核。")
            break

        chunk = prompt_next_chunk_after("02 半自动除错", path_a, "02", chunks, chunk, finalized)


if __name__ == "__main__":
    run_fix()
