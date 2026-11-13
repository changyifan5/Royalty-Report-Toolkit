"""
恢复行序脚本（代码库 05）

按 __原序 列将第 7 行起的数据行重排为初始顺序，保留当前单元格内容与样式。
对每个含 __原序 的页签分别处理；无 __原序 的页签跳过。

可选对照「00-原始基准/」中同名文件，报告基准有但当前缺失的序号。

弹窗选择待恢复报表，写回原文件（先备份 .bak）。

运行前请关闭 Excel。
运行：python "05 恢复行序.py"
"""

from __future__ import annotations

import sys
import time
from collections import Counter
from pathlib import Path

from openpyxl import load_workbook

from excel_common import backup_file
from file_picker_common import require_excel_file
from row_order_common import (
    BASELINE_DIR,
    DATA_START_ROW,
    FILL_YELLOW,
    SEQ_HEADER,
    collect_seq_values,
    find_last_data_row,
    find_seq_column,
    format_seq_col,
    parse_seq,
    read_row,
    row_has_content,
    sheet_used_max_col,
    write_row,
)

BASE_DIR = Path(__file__).resolve().parent


def load_baseline_seqs(path: Path) -> dict[str, set[int]] | None:
    baseline_path = BASELINE_DIR / path.name
    if not baseline_path.exists():
        return None

    wb = load_workbook(baseline_path, read_only=True, data_only=True)
    result: dict[str, set[int]] = {}
    for ws in wb.worksheets:
        header_map: dict[str, int] = {}
        for col in range(1, 501):
            val = ws.cell(row=6, column=col).value
            if val is not None and str(val).strip():
                header_map[str(val).strip()] = col
        seq_col = header_map.get(SEQ_HEADER)
        if seq_col is None:
            continue
        last_row = find_last_data_row(ws, DATA_START_ROW)
        if last_row >= DATA_START_ROW:
            result[ws.title] = collect_seq_values(ws, seq_col, DATA_START_ROW, last_row)
    wb.close()
    return result or None


def reorder_sheet(ws, stats: dict) -> None:
    seq_col = find_seq_column(ws)
    if seq_col is None:
        stats["sheets_skipped"] += 1
        return

    last_row = find_last_data_row(ws)
    max_col = sheet_used_max_col(ws)

    with_seq: list[tuple[int, int, list]] = []
    without_seq: list[tuple[int, list]] = []
    seq_counter: Counter[int] = Counter()

    for row in range(DATA_START_ROW, last_row + 1):
        if not row_has_content(ws, row, max_col):
            continue
        seq = parse_seq(ws.cell(row=row, column=seq_col).value)
        row_data = read_row(ws, row, max_col)
        if seq is None:
            without_seq.append((row, row_data))
            stats["rows_no_seq"] += 1
        else:
            with_seq.append((seq, row, row_data))
            seq_counter[seq] += 1

    dup_seqs = sorted(seq for seq, count in seq_counter.items() if count > 1)
    if dup_seqs:
        stats["duplicate_seqs"].extend((ws.title, seq) for seq in dup_seqs)

    with_seq.sort(key=lambda item: (item[0], item[1]))

    target_row = DATA_START_ROW
    for _seq, _orig_row, row_data in with_seq:
        write_row(ws, target_row, row_data)
        target_row += 1

    for orig_row, row_data in without_seq:
        write_row(ws, target_row, row_data)
        ws.cell(row=target_row, column=seq_col).fill = FILL_YELLOW
        stats["rows_no_seq_highlighted"] += 1
        target_row += 1

    for row in range(target_row, last_row + 1):
        for col in range(1, max_col + 1):
            ws.cell(row=row, column=col).value = None

    stats["sheets_done"] += 1
    stats["rows_sorted"] += len(with_seq)
    stats["sheet_details"].append(
        (ws.title, format_seq_col(seq_col), len(with_seq), len(without_seq))
    )


def compare_baseline(path: Path, baseline_by_sheet: dict[str, set[int]] | None, stats: dict) -> None:
    if baseline_by_sheet is None:
        stats["baseline_missing_file"] = True
        return

    wb = load_workbook(path, read_only=True, data_only=True)
    for ws in wb.worksheets:
        if ws.title not in baseline_by_sheet:
            continue
        header_map: dict[str, int] = {}
        for col in range(1, 501):
            val = ws.cell(row=6, column=col).value
            if val is not None and str(val).strip():
                header_map[str(val).strip()] = col
        seq_col = header_map.get(SEQ_HEADER)
        if seq_col is None:
            continue

        last_row = find_last_data_row(ws, DATA_START_ROW)
        if last_row < DATA_START_ROW:
            continue

        current_seqs = collect_seq_values(ws, seq_col, DATA_START_ROW, last_row)
        missing = sorted(baseline_by_sheet[ws.title] - current_seqs)
        if missing:
            stats["missing_from_baseline"].append((ws.title, missing))
    wb.close()


def run_restore() -> None:
    t0 = time.time()
    path = require_excel_file("选择待恢复行序的报表")

    baseline_by_sheet = load_baseline_seqs(path)
    backup_path = backup_file(path)
    print(f"读取: {path}")
    print(f"  已备份: {backup_path.name}")

    wb = load_workbook(path)
    stats = {
        "sheets_done": 0,
        "sheets_skipped": 0,
        "rows_sorted": 0,
        "rows_no_seq": 0,
        "rows_no_seq_highlighted": 0,
        "duplicate_seqs": [],
        "sheet_details": [],
        "missing_from_baseline": [],
        "baseline_missing_file": False,
    }

    for ws in wb.worksheets:
        reorder_sheet(ws, stats)

    if stats["sheets_done"] == 0:
        wb.close()
        print(f"错误：所有页签均未找到「{SEQ_HEADER}」列，请先运行 05 建档行序.py。")
        sys.exit(1)

    wb.save(path)
    wb.close()

    compare_baseline(path, baseline_by_sheet, stats)

    print()
    print("=" * 50)
    print("恢复行序完成")
    print("=" * 50)
    for sheet_name, seq_label, sorted_count, no_seq_count in stats["sheet_details"]:
        print(f"  页签「{sheet_name}」: {seq_label}，重排 {sorted_count} 行", end="")
        if no_seq_count:
            print(f"，无序号 {no_seq_count} 行（已排至末尾并标黄）")
        else:
            print()
    if stats["sheets_skipped"]:
        print(f"  跳过（无 {SEQ_HEADER}）: {stats['sheets_skipped']} 个页签")
    if stats["duplicate_seqs"]:
        print("  警告：发现重复 __原序：")
        for sheet_name, seq in stats["duplicate_seqs"]:
            print(f"    {sheet_name}: {seq}")
    if stats["baseline_missing_file"]:
        print(f"  提示：未找到基准文件 {BASELINE_DIR / path.name}，未做缺失行对照。")
    elif stats["missing_from_baseline"]:
        print("  基准有、当前缺失的 __原序：")
        for sheet_name, missing in stats["missing_from_baseline"]:
            preview = missing[:20]
            suffix = f" …共 {len(missing)} 个" if len(missing) > 20 else ""
            print(f"    {sheet_name}: {preview}{suffix}")
    print()
    print(f"  已保存: {path}")
    print(f"  备份: {backup_path}")
    print(f"  总耗时: {time.time() - t0:.1f} 秒")


if __name__ == "__main__":
    run_restore()
