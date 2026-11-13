"""
Excel VLOOKUP 四列匹配脚本（代码库 01）

主文件 A：弹窗选择；第 7 行起，用 A 列作为查找键
比对文件 B：从「01-vlookup-比对表/」读取（文件夹内只保留 1 个 Excel）
返回 B 文件的 B、D、E、F 四列，追加到主文件第一个页签末尾（保持原页签名称不变）。

运行前请关闭 Excel 中已打开的文件。
运行：python "01 excel_vlookup.py"
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from excel_common import (
    append_log,
    backup_file,
    find_header_last_col,
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
    DATA_START_ROW_A, FILE_B, HEADER_ROW_A, load_lookup_table,
    resolve_result_start_col, write_lookup_rows,
)

# ---------------------------------------------------------------------------
# 可配置参数
# ---------------------------------------------------------------------------
LOG_FILE = Path(__file__).with_name("01_vlookup.log")


def run_vlookup() -> None:
    t0 = time.time()
    path_a = require_excel_file("选择主文件 A（待 VLOOKUP 的报表）")
    visible_only = prompt_visible_rows_only("01 VLOOKUP", path_a, "01")
    if visible_only is None:
        return
    path_b = resolve_excel_path(FILE_B, "比对文件 B")

    print(f"主文件 A: {path_a.name} ({path_a.stat().st_size // 1024 // 1024} MB)")
    print(f"比对文件 B: {path_b.name}")

    chunks = load_file_chunks(path_a, DATA_START_ROW_A)
    chunk = prompt_row_chunk("01 VLOOKUP", path_a, DATA_START_ROW_A, script_key="01", chunks=chunks)
    if chunk is None:
        print("已取消。")
        sys.exit(0)

    print("  读取比对表...")
    headers, lookup = load_lookup_table(path_b)
    print(f"  比对表键数量: {len(lookup)}")

    while chunk is not None:
        print_chunk_banner("01 VLOOKUP", chunk)

        work_path, backup_path = prepare_chunk_work(
            path_a, "01", chunk, backup_file, visible_only=visible_only,
        )
        if chunk.is_first:
            print(f"  已备份: {backup_path.name}")

        print("  打开主文件...")
        wb_a = load_workbook(work_path)
        ws = wb_a.worksheets[0]
        sheet_name = ws.title

        rows = select_chunk_rows(ws, chunk, visible_only)
        total_rows = len(rows)

        last_header_col = find_header_last_col(ws, HEADER_ROW_A)
        if last_header_col == 0:
            wb_a.close()
            err_msg = f"第 {HEADER_ROW_A} 行未找到表头"
            print(f"错误：{err_msg}")
            show_error_popup("01 VLOOKUP", err_msg)
            append_log([f"[ERROR] {err_msg}", f"主文件: {path_a}"], LOG_FILE)
            sys.exit(1)

        start_col = resolve_result_start_col(ws, headers, last_header_col)
        end_col = start_col + len(headers) - 1

        print(f"  写入 VLOOKUP 结果 {get_column_letter(start_col)}-{get_column_letter(end_col)}，本段 {total_rows} 行...")
        matched = write_lookup_rows(ws, rows, start_col, headers, lookup)

        print("  保存...")
        wb_a.save(work_path)
        wb_a.close()

        finalized, out_path = finalize_chunk_output(path_a, "01", chunk)

        print()
        print("=" * 50)
        print("全部段完成，处理完成" if finalized else "本段处理完成（尚未输出最终文件）")
        print("=" * 50)
        print(f"  工作表: {sheet_name}（页签名称未改）")
        print(f"  结果列: {get_column_letter(start_col)}-{get_column_letter(end_col)}")
        print(f"  数据行: 本段 {total_rows}，匹配成功: {matched}")
        if finalized and out_path:
            print(f"  已输出: {out_path}")
        print(f"  总耗时: {time.time() - t0:.1f} 秒")
        if finalized:
            print()
            print("提示：四列在现有表头右侧，请在第一个页签向右滚动查看。")
            break

        chunk = prompt_next_chunk_after("01 VLOOKUP", path_a, "01", chunks, chunk, finalized)


if __name__ == "__main__":
    run_vlookup()
