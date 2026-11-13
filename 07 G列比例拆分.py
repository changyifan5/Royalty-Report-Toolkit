"""
G 列比例拆分脚本（代码库 07）

对 A～D 列「字体颜色」为指定颜色的行，将 G 列按比例拆成两个数，写入末尾两列。
仅看字体颜色，忽略单元格背景色/填充色（例如黄色底纹不算）。

运行时会依次询问：
  1. 弹窗选择 Excel 文件
  2. 字体颜色（从表格中扫描到的颜色列表里选）
  3. 拆分比例（如 3:7）
  4. 两列输出表头名

第 6 行表头，第 7 行起为数据；处理第一个页签。
Mac：只读快速扫描 + 批量读 G 列（openpyxl，无需 Excel）。
运行前请关闭 Excel。
运行：python "07 G列比例拆分.py"
"""

from __future__ import annotations

import sys
import time
import tkinter as tk
from collections import Counter
from pathlib import Path
from tkinter import messagebox, simpledialog

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from excel_common import (
    EXCEL_SUFFIXES,
    backup_file,
    find_header_last_col,
    read_header_map,
    safe_print,
    tk_mono_font,
)
from excel_font_color import color_hint, font_rgb_triple, rgb_to_hex
from file_picker_common import require_excel_file
from report_chunk import (
    finalize_chunk_output,
    load_file_chunks,
    prepare_chunk_work,
    print_chunk_banner,
    prompt_next_chunk_after,
    prompt_row_chunk,
)

# ---------------------------------------------------------------------------
# 可配置参数
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
SOURCE_COL = 7  # G 列
HEADER_ROW = 6
DATA_START_ROW = 7
SHEET_INDEX = 0
HEADER_SCAN_LIMIT = 500
COLOR_CHECK_COLS = 4  # A～D，与 04 一致


def ask_string(title: str, prompt: str) -> str | None:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    value = simpledialog.askstring(title, prompt, parent=root)
    root.destroy()
    return value.strip() if value else None


def scan_chunk_colors(
    sheet, row_start: int, row_end: int
) -> tuple[Counter[tuple[int, int, int]], dict[int, set[tuple[int, int, int]]]]:
    """只读扫描：统计颜色 + 缓存每行 A～D 字体色（避免二次扫描）。"""
    color_counts: Counter[tuple[int, int, int]] = Counter()
    row_color_sets: dict[int, set[tuple[int, int, int]]] = {}

    scanned = 0
    for row_cells in sheet.iter_rows(
        min_row=row_start, max_row=row_end, min_col=1, max_col=COLOR_CHECK_COLS
    ):
        row_num = row_cells[0].row
        seen: set[tuple[int, int, int]] = set()
        for cell in row_cells:
            rgb = font_rgb_triple(cell.font)
            if rgb is not None:
                seen.add(rgb)
        row_color_sets[row_num] = seen
        for rgb in seen:
            color_counts[rgb] += 1
        scanned += 1
        if scanned % 2000 == 0:
            safe_print(f"  已扫描 {scanned} 行字体颜色...")

    return color_counts, row_color_sets


def read_g_column_batch(sheet, row_start: int, row_end: int) -> dict[int, tuple[object, str]]:
    batch: dict[int, tuple[object, str]] = {}
    for row_cells in sheet.iter_rows(
        min_row=row_start, max_row=row_end, min_col=SOURCE_COL, max_col=SOURCE_COL
    ):
        cell = row_cells[0]
        batch[cell.row] = (cell.value, cell.number_format or "General")
    return batch


def pick_color_from_list(color_counts: Counter[tuple[int, int, int]]) -> tuple[int, int, int] | None:
    if not color_counts:
        return None

    items = sorted(color_counts.items(), key=lambda x: (-x[1], x[0]))
    labels = [
        f"{rgb_to_hex(rgb)}  {color_hint(rgb)}  （约 {count} 行）"
        for rgb, count in items
    ]

    choice: dict[str, tuple[int, int, int] | None] = {"rgb": None}

    root = tk.Tk()
    root.title("选择字体颜色")
    root.attributes("-topmost", True)
    root.geometry("420x320")

    tk.Label(root, text="选择要处理的行（A～D 列字体颜色）：", anchor="w").pack(fill="x", padx=10, pady=(10, 4))

    listbox = tk.Listbox(root, height=min(12, len(labels)), font=tk_mono_font(10))
    for label in labels:
        listbox.insert(tk.END, label)
    listbox.selection_set(0)
    listbox.pack(fill="both", expand=True, padx=10, pady=4)

    def on_ok() -> None:
        sel = listbox.curselection()
        if not sel:
            messagebox.showwarning("提示", "请先选择一种颜色。", parent=root)
            return
        choice["rgb"] = items[sel[0]][0]
        root.destroy()

    def on_cancel() -> None:
        root.destroy()

    btn_frame = tk.Frame(root)
    btn_frame.pack(fill="x", padx=10, pady=10)
    tk.Button(btn_frame, text="确定", width=10, command=on_ok).pack(side="left")
    tk.Button(btn_frame, text="取消", width=10, command=on_cancel).pack(side="left", padx=8)

    root.mainloop()
    return choice["rgb"]


def parse_ratio(spec: str) -> tuple[float, float]:
    text = spec.strip().replace("：", ":").replace("，", ",")
    if ":" in text:
        a, b = text.split(":", 1)
    elif "," in text:
        a, b = text.split(",", 1)
    else:
        parts = text.split()
        if len(parts) != 2:
            raise ValueError("比例需为两个数，如 3:7 或 0.3,0.7")
        a, b = parts

    r1, r2 = float(a.strip()), float(b.strip())
    if r1 < 0 or r2 < 0:
        raise ValueError("比例不能为负数")
    if r1 == 0 and r2 == 0:
        raise ValueError("比例不能同时为 0")
    return r1, r2


def ensure_output_columns(sheet, header_map: dict[str, int], name1: str, name2: str) -> tuple[int, int]:
    if name1 in header_map and name2 in header_map:
        return header_map[name1], header_map[name2]
    if name1 in header_map or name2 in header_map:
        print("错误：两个表头只能同时已存在或同时新建，当前只存在一个。")
        sys.exit(1)

    last = find_header_last_col(sheet)
    col1, col2 = last + 1, last + 2
    sheet.cell(row=HEADER_ROW, column=col1, value=name1)
    sheet.cell(row=HEADER_ROW, column=col2, value=name2)
    return col1, col2


def to_number(value) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


INTEGER_LIKE_FORMATS = frozenset({"0", "#,##0", "0_ ", "#,##0_ ", "0;-0;;@", "#,##0;-#,##0;;@"})


def output_number_format_from_fmt(fmt: str | None) -> str:
    fmt = (fmt or "General").strip()
    if fmt in INTEGER_LIKE_FORMATS:
        return "0.##########"
    return fmt


def write_split_cells(ws, row: int, col1: int, col2: int, v1: float, v2: float, g_fmt: str) -> None:
    fmt = output_number_format_from_fmt(g_fmt)
    c1 = ws.cell(row=row, column=col1, value=v1)
    c2 = ws.cell(row=row, column=col2, value=v2)
    c1.number_format = fmt
    c2.number_format = fmt


def apply_openpyxl_split(
    work_path: Path,
    chunk,
    matched_rows: list[int],
    r1: float,
    r2: float,
    header1: str,
    header2: str,
) -> tuple[int, int, dict[str, int], tuple[int, float, float, float] | None]:
    print("  openpyxl：批量写入拆分列...")
    t0 = time.time()
    wb = load_workbook(work_path, keep_links=False)
    ws = wb.worksheets[SHEET_INDEX]
    header_map = read_header_map(ws)
    col1, col2 = ensure_output_columns(ws, header_map, header1, header2)
    g_column = read_g_column_batch(ws, chunk.data_start_row, chunk.data_end_row)

    stats = {"matched": len(matched_rows), "filled": 0, "skipped_empty_g": 0}
    sample: tuple[int, float, float, float] | None = None

    for row in matched_rows:
        g_val_raw, g_fmt = g_column[row]
        g_val = to_number(g_val_raw)
        if g_val is None:
            stats["skipped_empty_g"] += 1
            continue
        v1, v2 = split_value(g_val, r1, r2)
        write_split_cells(ws, row, col1, col2, v1, v2, g_fmt)
        if sample is None:
            sample = (row, g_val, v1, v2)
        stats["filled"] += 1

    wb.save(work_path)
    wb.close()
    safe_print(f"  openpyxl 耗时: {time.time() - t0:.1f}s")
    return col1, col2, stats, sample


def split_value(total: float, r1: float, r2: float) -> tuple[float, float]:
    part1 = round(total * r1 / (r1 + r2), 10)
    part2 = total - part1
    return part1, part2


def run_split() -> None:
    t0 = time.time()
    path = require_excel_file("选择要拆分 G 列比例的 Excel")
    print(f"  文件: {path.name}")

    chunks = load_file_chunks(path, DATA_START_ROW)
    chunk = prompt_row_chunk("07 G列比例拆分", path, DATA_START_ROW, script_key="07", chunks=chunks)
    if chunk is None:
        print("已取消。")
        sys.exit(0)
    print_chunk_banner("07 G列比例拆分", chunk)

    print("  扫描字体颜色（只读模式）...")
    t_scan = time.time()
    wb_scan = load_workbook(path, read_only=True, keep_links=False)
    if len(wb_scan.worksheets) <= SHEET_INDEX:
        wb_scan.close()
        print("错误：工作簿中没有可用的工作表")
        sys.exit(1)
    ws_scan = wb_scan.worksheets[SHEET_INDEX]
    sheet_name = ws_scan.title
    color_counts, _row_color_sets = scan_chunk_colors(
        ws_scan, chunk.data_start_row, chunk.data_end_row
    )
    wb_scan.close()
    safe_print(f"  扫描完成，耗时 {time.time() - t_scan:.1f}s")

    if not color_counts:
        print(f"错误：本段（Excel 第 {chunk.data_start_row}～{chunk.data_end_row} 行）A～D 列未扫描到带颜色的字体。")
        sys.exit(1)

    print("  扫描到的字体颜色:")
    for rgb, count in sorted(color_counts.items(), key=lambda x: -x[1]):
        safe_print(f"    {rgb_to_hex(rgb)}  {color_hint(rgb)}  约 {count} 行")

    target_rgb = pick_color_from_list(color_counts)
    if target_rgb is None:
        print("已取消。")
        sys.exit(0)

    print(f"  已选颜色: {rgb_to_hex(target_rgb)} ({color_hint(target_rgb)})")

    ratio_spec = ask_string("比例", "拆分比例（例 3:7 或 0.3,0.7）:")
    if not ratio_spec:
        print("已取消。")
        sys.exit(0)

    header1 = ask_string("表头 1", "第一列输出表头名:")
    if not header1:
        print("已取消。")
        sys.exit(0)

    header2 = ask_string("表头 2", "第二列输出表头名:")
    if not header2:
        print("已取消。")
        sys.exit(0)

    if header1 == header2:
        print("错误：两列表头名不能相同。")
        sys.exit(1)

    try:
        r1, r2 = parse_ratio(ratio_spec)
    except ValueError as exc:
        print(f"错误：{exc}")
        sys.exit(1)

    while chunk is not None:
        if chunk.chunk_index > 1:
            print_chunk_banner("07 G列比例拆分", chunk)

        work_path, backup_path = prepare_chunk_work(path, "07", chunk, backup_file)
        if chunk.is_first:
            print(f"  已备份: {backup_path.name}")

        print("  扫描本段字体颜色...")
        wb_scan = load_workbook(work_path, read_only=True, keep_links=False)
        ws_scan = wb_scan.worksheets[SHEET_INDEX]
        _color_counts, row_color_sets = scan_chunk_colors(
            ws_scan, chunk.data_start_row, chunk.data_end_row
        )
        wb_scan.close()

        matched_rows = [
            row
            for row in range(chunk.data_start_row, chunk.data_end_row + 1)
            if target_rgb in row_color_sets.get(row, ())
        ]
        print(f"  本段匹配行: {len(matched_rows)} 行（仅处理这些行的 G 列）")

        col1, col2, stats, sample = apply_openpyxl_split(
            work_path, chunk, matched_rows, r1, r2, header1, header2
        )

        finalized, out_path = finalize_chunk_output(path, "07", chunk)

        print()
        print("=" * 50)
        print("全部段完成，G 列比例拆分完成" if finalized else "本段处理完成（尚未输出最终文件）")
        print("=" * 50)
        print(f"  工作表: {sheet_name}")
        print(f"  匹配颜色: {rgb_to_hex(target_rgb)} ({color_hint(target_rgb)})")
        print(f"  G 列: {get_column_letter(SOURCE_COL)}")
        print(f"  比例: {r1}:{r2}")
        print(f"  输出列: {get_column_letter(col1)}={header1!r}, {get_column_letter(col2)}={header2!r}")
        print(f"  匹配行: {stats['matched']} 行")
        print(f"  已填写: {stats['filled']} 行")
        print(f"  跳过（G 为空）: {stats['skipped_empty_g']} 行")
        if sample:
            s_row, s_g, s_v1, s_v2 = sample
            print(f"  示例（第 {s_row} 行）: G={s_g} → {s_v1} + {s_v2} = {s_v1 + s_v2}")
        if finalized and out_path:
            print(f"  已输出: {out_path}")
        print(f"  备份: {backup_path}")
        print(f"  总耗时: {time.time() - t0:.1f} 秒")
        if finalized:
            break

        chunk = prompt_next_chunk_after("07 G列比例拆分", path, "07", chunks, chunk, finalized)


if __name__ == "__main__":
    run_split()
