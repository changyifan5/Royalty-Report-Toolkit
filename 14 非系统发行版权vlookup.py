"""
非系统发行版权 VLOOKUP（代码库 14）

仿照 01：主文件 A 弹窗选择；第 7 行起用 A 列作为查找键。
比对文件 B：弹窗选择文件 + 选择页签。
  - 表头在第 2 行，数据从第 3 行起
  - 查找键：B 的 A 列
  - 返回 B 的 A～D 四列，追加到主文件第一个页签表头右侧
  - 若 A 列未匹配：询问是否用主报表 B 列二次匹配（比对表按 B 列匹配，只回填 B～D）

Windows / Mac 均用 openpyxl（只读写单元格，无需 Excel COM）。

运行前请关闭 Excel 中已打开的文件。
"""

from __future__ import annotations

import sys
import time
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter

from excel_common import (
    backup_file,
    find_header_last_col,
    list_sheet_names_fast,
    normalize_key,
    tk_mono_font,
    tk_ui_font,
)
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
DATA_START_ROW_A = 7
DATA_START_ROW_B = 3
LOOKUP_KEY_COL_A = "A"
LOOKUP_KEY_COL_B = "A"
RETURN_COLS_B = ["A", "B", "C", "D"]
FALLBACK_KEY_COL_A = "B"
FALLBACK_KEY_COL_B = "B"
FALLBACK_RETURN_COLS_B = ["B", "C", "D"]
HEADER_ROW_A = 6
HEADER_ROW_B = 2
NOT_FOUND_VALUE = "#N/A"
SCRIPT_KEY = "14"
SCRIPT_TITLE = "14 非系统发行版权 VLOOKUP"

# 长任务进度打印步长（每处理多少行打印一次）。
PROGRESS_STEP = 10_000


def list_sheet_names(path: Path) -> list[str]:
    return list_sheet_names_fast(path)


def pick_sheet(path: Path) -> str | None:
    names = list_sheet_names(path)
    if not names:
        print("错误：比对文件中没有页签。")
        sys.exit(1)
    if len(names) == 1:
        print(f"  比对页签（仅一个）: {names[0]}")
        return names[0]

    root = tk.Tk()
    root.title(f"选择比对页签 — {path.name}")
    root.attributes("-topmost", True)
    root.resizable(False, False)

    selected: list[str] = []

    frame = ttk.Frame(root, padding=12)
    frame.grid(row=0, column=0, sticky="nsew")

    ttk.Label(
        frame,
        text="请选择比对用报表的页签：",
        font=tk_ui_font(10),
    ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))

    listbox = tk.Listbox(
        frame,
        width=56,
        height=min(14, max(4, len(names))),
        font=tk_mono_font(10),
    )
    listbox.grid(row=1, column=0, columnspan=2, sticky="ew")
    for name in names:
        listbox.insert(tk.END, name)
    listbox.selection_set(0)

    def confirm() -> None:
        sel = listbox.curselection()
        if not sel:
            messagebox.showwarning("提示", "请先选择一个页签。", parent=root)
            return
        selected.append(names[sel[0]])
        root.destroy()

    def cancel() -> None:
        root.destroy()

    btn = ttk.Frame(frame)
    btn.grid(row=2, column=0, columnspan=2, pady=(10, 0), sticky="e")
    ttk.Button(btn, text="确定", command=confirm, width=10).grid(row=0, column=0, padx=(0, 6))
    ttk.Button(btn, text="取消", command=cancel, width=10).grid(row=0, column=1)

    root.bind("<Return>", lambda _e: confirm())
    root.bind("<Escape>", lambda _e: cancel())
    listbox.bind("<Double-Button-1>", lambda _e: confirm())
    root.mainloop()

    return selected[0] if selected else None


def _header_label(ws, col_idx: int) -> str:
    val = ws.cell(row=HEADER_ROW_B, column=col_idx).value
    if val is not None and str(val).strip():
        return str(val).strip()
    return f"match_{get_column_letter(col_idx)}"


def load_lookup_table(
    path_b: Path, sheet_name: str
) -> tuple[list[str], dict[str, tuple[object, ...]], dict[str, tuple[object, ...]]]:
    """返回 (A～D 表头, A 键→A～D, B 键→B～D)。"""
    key_a_idx = column_index_from_string(LOOKUP_KEY_COL_B)
    key_b_idx = column_index_from_string(FALLBACK_KEY_COL_B)
    ret_a_idxs = [column_index_from_string(col) for col in RETURN_COLS_B]
    ret_b_idxs = [column_index_from_string(col) for col in FALLBACK_RETURN_COLS_B]
    min_col = min(key_a_idx, key_b_idx, *ret_a_idxs, *ret_b_idxs)
    max_col = max(key_a_idx, key_b_idx, *ret_a_idxs, *ret_b_idxs)

    wb = load_workbook(path_b, read_only=True, data_only=True)
    if sheet_name not in wb.sheetnames:
        wb.close()
        print(f"错误：比对文件中找不到页签「{sheet_name}」。")
        sys.exit(1)
    ws = wb[sheet_name]

    headers = [_header_label(ws, idx) for idx in ret_a_idxs]

    lookup_by_a: dict[str, tuple[object, ...]] = {}
    lookup_by_b: dict[str, tuple[object, ...]] = {}
    for row in ws.iter_rows(
        min_row=DATA_START_ROW_B,
        min_col=min_col,
        max_col=max_col,
        values_only=True,
    ):
        key_a = normalize_key(row[key_a_idx - min_col])
        if key_a is not None:
            lookup_by_a[key_a] = tuple(row[idx - min_col] for idx in ret_a_idxs)

        key_b = normalize_key(row[key_b_idx - min_col])
        if key_b is not None:
            lookup_by_b[key_b] = tuple(row[idx - min_col] for idx in ret_b_idxs)

    wb.close()
    return headers, lookup_by_a, lookup_by_b


def ask_fallback_b_lookup(unmatched: int) -> bool:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    ok = messagebox.askyesno(
        SCRIPT_TITLE,
        f"本段有 {unmatched:,} 行用 A 列未匹配到。\n\n"
        f"是否用主报表 B 列二次 VLOOKUP？\n"
        f"（比对表按 B 列匹配，只回填 B～D；A 列结果仍为 {NOT_FOUND_VALUE}）",
        parent=root,
    )
    root.destroy()
    return ok


def resolve_result_start_col(sheet, headers: list[str], last_header_col: int) -> int:
    start_col = last_header_col + 1
    num_cols = len(headers)
    if last_header_col >= num_cols:
        candidate = last_header_col - num_cols + 1
        if all(
            str(sheet.cell(row=HEADER_ROW_A, column=candidate + i).value or "").strip()
            == headers[i]
            for i in range(num_cols)
        ):
            return candidate
    return start_col


def run_vlookup() -> None:
    t0 = time.time()
    path_a = require_excel_file("选择主文件 A（待 VLOOKUP 的报表）")
    path_b = require_excel_file("选择比对文件 B（非系统发行版权）")

    print(f"主文件 A: {path_a.name} ({path_a.stat().st_size // 1024 // 1024} MB)")
    print(f"比对文件 B: {path_b.name}")

    sheet_b = pick_sheet(path_b)
    if sheet_b is None:
        print("已取消。")
        sys.exit(0)
    print(f"  比对页签: {sheet_b}")

    chunks = load_file_chunks(path_a, DATA_START_ROW_A)
    chunk = prompt_row_chunk(
        SCRIPT_TITLE, path_a, DATA_START_ROW_A, script_key=SCRIPT_KEY, chunks=chunks
    )
    if chunk is None:
        print("已取消。")
        sys.exit(0)

    print("  读取比对表...")
    headers, lookup_by_a, lookup_by_b = load_lookup_table(path_b, sheet_b)
    print(f"  比对表 A 键数量: {len(lookup_by_a)}")
    print(f"  比对表 B 键数量: {len(lookup_by_b)}")
    print(f"  返回列标题: {headers}")

    # None=尚未询问；True/False=本轮运行内沿用选择
    do_fallback_b: bool | None = None

    while chunk is not None:
        print_chunk_banner(SCRIPT_TITLE, chunk)

        work_path, backup_path = prepare_chunk_work(path_a, SCRIPT_KEY, chunk, backup_file)
        if chunk.is_first:
            print(f"  已备份: {backup_path.name}")

        print("  打开主文件...")
        wb_a = load_workbook(work_path)
        ws = wb_a.worksheets[0]
        sheet_name = ws.title

        key_col_a = column_index_from_string(LOOKUP_KEY_COL_A)
        key_col_b = column_index_from_string(FALLBACK_KEY_COL_A)
        row_start = chunk.data_start_row
        row_end = chunk.data_end_row
        total_rows = row_end - row_start + 1

        last_header_col = find_header_last_col(ws, HEADER_ROW_A)
        if last_header_col == 0:
            wb_a.close()
            print(f"错误：第 {HEADER_ROW_A} 行未找到表头")
            sys.exit(1)

        start_col = resolve_result_start_col(ws, headers, last_header_col)
        end_col = start_col + len(headers) - 1

        for i, header in enumerate(headers):
            ws.cell(row=HEADER_ROW_A, column=start_col + i, value=header)

        print(
            f"  写入 VLOOKUP 结果 {get_column_letter(start_col)}-{get_column_letter(end_col)}，"
            f"本段 {total_rows} 行..."
        )
        matched_a = 0
        unmatched_rows: list[int] = []
        # 提速：先 values_only 批量读主表 A 列键，再缓存结果区（A～D）cell 对象统一赋值。
        # A 精确命中 / 未命中填 #N/A 语义不变。
        key_a_values = [
            row[0]
            for row in ws.iter_rows(
                min_row=row_start,
                max_row=row_end,
                min_col=key_col_a,
                max_col=key_col_a,
                values_only=True,
            )
        ]
        result_cells = list(
            ws.iter_rows(
                min_row=row_start,
                max_row=row_end,
                min_col=start_col,
                max_col=end_col,
            )
        )
        for offset, raw_key in enumerate(key_a_values):
            row = row_start + offset
            key = normalize_key(raw_key)
            cells = result_cells[offset]
            if key and key in lookup_by_a:
                values = lookup_by_a[key]
                matched_a += 1
            else:
                values = (NOT_FOUND_VALUE,) * len(headers)
                unmatched_rows.append(row)
            for i, val in enumerate(values):
                cells[i].value = val

            if (offset + 1) % PROGRESS_STEP == 0:
                print(f"  进度: {offset + 1:,}/{total_rows} 行，A 列已匹配 {matched_a:,}")

        matched_b = 0
        if unmatched_rows:
            if do_fallback_b is None:
                do_fallback_b = ask_fallback_b_lookup(len(unmatched_rows))
                if do_fallback_b:
                    print("  已选择：对未匹配行用主报表 B 列二次 VLOOKUP")
                else:
                    print("  已跳过：B 列二次 VLOOKUP")

            if do_fallback_b:
                print(
                    f"  B 列二次匹配：未匹配 {len(unmatched_rows):,} 行，"
                    f"回填 {get_column_letter(start_col + 1)}-"
                    f"{get_column_letter(start_col + len(FALLBACK_RETURN_COLS_B))}..."
                )
                # 批量读主表 B 列键 + 缓存 B～D 三列 cell 对象；只回填 B～D，A 列保留 #N/A。
                key_b_values = [
                    row[0]
                    for row in ws.iter_rows(
                        min_row=row_start,
                        max_row=row_end,
                        min_col=key_col_b,
                        max_col=key_col_b,
                        values_only=True,
                    )
                ]
                fallback_cells = list(
                    ws.iter_rows(
                        min_row=row_start,
                        max_row=row_end,
                        min_col=start_col + 1,
                        max_col=start_col + len(FALLBACK_RETURN_COLS_B),
                    )
                )
                for row in unmatched_rows:
                    offset = row - row_start
                    key = normalize_key(key_b_values[offset])
                    if not key or key not in lookup_by_b:
                        continue
                    values = lookup_by_b[key]
                    cells = fallback_cells[offset]
                    for i, val in enumerate(values):
                        cells[i].value = val
                    matched_b += 1

        print("  保存...")
        wb_a.save(work_path)
        wb_a.close()

        finalized, out_path = finalize_chunk_output(path_a, SCRIPT_KEY, chunk)

        print()
        print("=" * 50)
        print("全部段完成，处理完成" if finalized else "本段处理完成（尚未输出最终文件）")
        print("=" * 50)
        print(f"  工作表: {sheet_name}（页签名称未改）")
        print(f"  比对页签: {sheet_b}")
        print(f"  结果列: {get_column_letter(start_col)}-{get_column_letter(end_col)}（来自 B 的 A～D）")
        print(f"  数据行: 本段 {total_rows}，A 列匹配: {matched_a}")
        if do_fallback_b:
            print(f"  B 列二次匹配成功: {matched_b}（仅回填 B～D）")
        elif unmatched_rows and do_fallback_b is False:
            print(f"  A 列未匹配: {len(unmatched_rows)}（已跳过 B 列二次匹配）")
        if finalized and out_path:
            print(f"  已输出: {out_path}")
        print(f"  总耗时: {time.time() - t0:.1f} 秒")
        if finalized:
            print()
            print("提示：四列在现有表头右侧，请在第一个页签向右滚动查看。")
            break

        chunk = prompt_next_chunk_after(
            SCRIPT_TITLE, path_a, SCRIPT_KEY, chunks, chunk, finalized
        )


if __name__ == "__main__":
    run_vlookup()
