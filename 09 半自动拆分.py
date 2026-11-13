"""
半自动拆分脚本（代码库 09）

运行时：弹窗选 Excel → 再选拆分步骤。

步骤（A～D 列，第 7 行起；投母看灰底+加粗）→ 新页签，前 6 行格式/列宽/行高与源页签一致：
  1. 退回     纯红字体      → 需退回
  2. 需修改   蓝色字体      → 需修改
  3. 线下     紫色字体      → 线下
  4. 街版     橙色字体      → 街版
  5. 投母     灰色背景+加粗 → 投母

结果写回原文件；目标页签已存在时会先询问是否删除重建。
Windows 使用 Excel COM 批量复制；Mac 使用 openpyxl。
运行：python "09 半自动拆分.py"
"""

from __future__ import annotations

import base64
import json
import sys
import time
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import messagebox, ttk

from openpyxl import load_workbook

BASE_DIR = Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from excel_common import (
    EXCEL_SUFFIXES,
    backup_file,
    list_sheet_names_fast,
    run_powershell,
    safe_print,
    tk_mono_font,
    tk_ui_font,
)
from excel_font_color import (
    is_blue_font,
    is_gray_bold_cell,
    is_orange_font,
    is_purple_font,
    is_red_font,
)
from excel_sheet_copy import copy_header_layout, copy_row_range
from row_order_common import find_last_data_row
from file_picker_common import require_excel_file
from report_chunk import (
    RowChunk,
    finalize_chunk_output,
    load_file_chunks,
    prepare_chunk_work,
    print_chunk_banner,
    prompt_next_chunk_after,
    prompt_row_chunk,
)

PS1_SPLIT = BASE_DIR / "09_split_return_com.ps1"
LOG_FILE = BASE_DIR / "09_com.log"
MANIFEST_FILE = BASE_DIR / "09_return_manifest.json"

HEADER_ROW = 6
DATA_START_ROW = 7
SHEET_INDEX = 0
COLOR_CHECK_COLS = 4

HEADER_SCAN_LIMIT = 500


def compute_row_groups(rows: list[int]) -> list[list[int]]:
    if not rows:
        return []
    groups: list[list[int]] = []
    start = rows[0]
    prev = rows[0]
    for row in rows[1:]:
        if row != prev + 1:
            groups.append([start, prev])
            start = row
        prev = row
    groups.append([start, prev])
    return groups


@dataclass(frozen=True)
class SplitOperation:
    key: str
    title: str
    description: str
    sheet_name: str
    scan_mode: str


OPERATIONS: list[SplitOperation] = [
    SplitOperation(
        key="return",
        title="1. 退回",
        description="A～D 纯红字体 → 页签「需退回」",
        sheet_name="需退回",
        scan_mode="red_font",
    ),
    SplitOperation(
        key="modify",
        title="2. 需修改",
        description="A～D 蓝色字体 → 页签「需修改」",
        sheet_name="需修改",
        scan_mode="blue_font",
    ),
    SplitOperation(
        key="offline",
        title="3. 线下",
        description="A～D 紫色字体 → 页签「线下」",
        sheet_name="线下",
        scan_mode="purple_font",
    ),
    SplitOperation(
        key="street",
        title="4. 街版",
        description="A～D 橙色字体 → 页签「街版」",
        sheet_name="街版",
        scan_mode="orange_font",
    ),
    SplitOperation(
        key="parent",
        title="5. 投母",
        description="A～D 灰色背景且加粗 → 页签「投母」",
        sheet_name="投母",
        scan_mode="gray_bold",
    ),
]


def pick_operation(file_name: str) -> SplitOperation | None:
    root = tk.Tk()
    root.title(f"半自动拆分 — {file_name}")
    root.attributes("-topmost", True)
    root.resizable(False, False)

    selected: list[SplitOperation] = []

    frame = ttk.Frame(root, padding=12)
    frame.grid(row=0, column=0, sticky="nsew")

    ttk.Label(frame, text="请选择要执行的拆分步骤：", font=tk_ui_font(10)).grid(
        row=0, column=0, columnspan=2, sticky="w", pady=(0, 8)
    )

    listbox = tk.Listbox(frame, width=78, height=min(12, max(6, len(OPERATIONS))), font=tk_mono_font(10))
    listbox.grid(row=1, column=0, columnspan=2, sticky="ew")
    for idx, op in enumerate(OPERATIONS):
        listbox.insert(tk.END, f"{op.title}  —  {op.description}")
    listbox.selection_set(0)

    def confirm() -> None:
        sel = listbox.curselection()
        if not sel:
            messagebox.showwarning("提示", "请先选择一个步骤。", parent=root)
            return
        selected.append(OPERATIONS[sel[0]])
        root.destroy()

    def cancel() -> None:
        root.destroy()

    btn_frame = ttk.Frame(frame)
    btn_frame.grid(row=2, column=0, columnspan=2, pady=(10, 0), sticky="e")
    ttk.Button(btn_frame, text="确定", command=confirm, width=10).grid(row=0, column=0, padx=(0, 6))
    ttk.Button(btn_frame, text="取消", command=cancel, width=10).grid(row=0, column=1)

    root.bind("<Return>", lambda _e: confirm())
    root.bind("<Escape>", lambda _e: cancel())
    root.mainloop()

    return selected[0] if selected else None


def row_has_content(row_cells) -> bool:
    return any(cell.value is not None and str(cell.value).strip() for cell in row_cells)


def cell_matches_mode(cell, scan_mode: str) -> bool:
    if scan_mode == "red_font":
        return is_red_font(cell.font)
    if scan_mode == "blue_font":
        return is_blue_font(cell.font)
    if scan_mode == "purple_font":
        return is_purple_font(cell.font)
    if scan_mode == "orange_font":
        return is_orange_font(cell.font)
    if scan_mode == "gray_bold":
        return is_gray_bold_cell(cell)
    return False


def row_matches_mode(row_cells, scan_mode: str) -> bool:
    return any(cell_matches_mode(cell, scan_mode) for cell in row_cells)


def scan_label(scan_mode: str) -> str:
    return {
        "red_font": "纯红字体",
        "blue_font": "蓝色字体",
        "purple_font": "紫色字体",
        "orange_font": "橙色字体",
        "gray_bold": "灰底+加粗",
    }.get(scan_mode, scan_mode)


def scan_matching_rows(path: Path, scan_mode: str, chunk: RowChunk) -> tuple[list[int], int]:
    label = scan_label(scan_mode)
    print(f"  扫描 A～D {label}（只读，本段 Excel 第 {chunk.data_start_row}～{chunk.data_end_row} 行）...")
    t0 = time.time()
    wb = load_workbook(path, read_only=True, keep_links=False)
    ws = wb.worksheets[SHEET_INDEX]
    matched: list[int] = []
    scanned = 0
    last_col = 0

    for row_cells in ws.iter_rows(
        min_row=HEADER_ROW,
        max_row=HEADER_ROW,
        min_col=1,
        max_col=HEADER_SCAN_LIMIT,
    ):
        for idx, cell in enumerate(row_cells, start=1):
            if cell.value is not None and str(cell.value).strip():
                last_col = idx

    for row_cells in ws.iter_rows(
        min_row=chunk.data_start_row,
        max_row=chunk.data_end_row,
        min_col=1,
        max_col=COLOR_CHECK_COLS,
    ):
        row_num = row_cells[0].row
        scanned += 1
        if scanned % 50000 == 0:
            safe_print(f"    已扫描 {scanned} 行，命中 {len(matched)} 行...")

        if not row_has_content(row_cells):
            continue
        if row_matches_mode(row_cells, scan_mode):
            matched.append(row_num)

    wb.close()
    safe_print(f"  扫描完成: {scanned} 行，命中 {len(matched)} 行，耗时 {time.time() - t0:.1f} 秒")
    return matched, max(last_col, 1)


def write_manifest(
    path: Path,
    sheet_name: str,
    data_rows: list[int],
    row_groups: list[list[int]],
    last_col: int,
    append_to_existing: bool,
) -> None:
    payload = {
        "path": str(path.resolve()),
        "dest_sheet_name_b64": base64.b64encode(sheet_name.encode("utf-8")).decode("ascii"),
        "source_sheet_index": SHEET_INDEX,
        "header_rows": HEADER_ROW,
        "data_start_row": DATA_START_ROW,
        "data_rows": data_rows,
        "row_groups": row_groups,
        "last_col": last_col,
        "append_to_existing": append_to_existing,
    }
    MANIFEST_FILE.write_text(json.dumps(payload, ensure_ascii=True), encoding="utf-8")


def run_com_split(
    path: Path,
    sheet_name: str,
    data_rows: list[int],
    row_groups: list[list[int]],
    last_col: int,
    append_to_existing: bool,
) -> tuple[str, int]:
    if not PS1_SPLIT.exists():
        print(f"错误：找不到 COM 脚本: {PS1_SPLIT}")
        sys.exit(1)

    group_count = len(row_groups)
    print(
        f"  Excel COM：{len(data_rows)} 行 → {group_count} 个连续块"
        f"（块越多越慢；颜色行很分散时会明显变慢）"
    )
    write_manifest(path, sheet_name, data_rows, row_groups, last_col, append_to_existing)

    print(f"  Excel COM：复制到「{sheet_name}」...")
    t_com = time.time()
    # 09 的 COM 输出编码固定为 utf-8（与 05/06/11 的 gbk 不同），必须显式指定。
    com_output = run_powershell(
        PS1_SPLIT,
        ["-ManifestPath", str(MANIFEST_FILE.resolve())],
        LOG_FILE,
        encoding="utf-8",
    )
    com_elapsed = time.time() - t_com
    timing: dict[str, int] = {}
    for line in com_output.splitlines():
        if line.startswith("TIMING,"):
            parts = line.split(",", 3)
            if len(parts) >= 3:
                timing[parts[1]] = int(parts[2])
        elif line.startswith("PROGRESS,"):
            parts = line.split(",", 4)
            if len(parts) >= 4:
                safe_print(f"    复制进度: {parts[2]}/{parts[3]} 块...")
        else:
            safe_print(f"  {line}")

    if timing:
        parts_desc = []
        if "open_ms" in timing:
            parts_desc.append(f"打开 {timing['open_ms']}ms")
        if "copy_ms" in timing:
            parts_desc.append(f"复制 {timing['copy_ms']}ms")
        if "save_ms" in timing:
            parts_desc.append(f"保存 {timing['save_ms']}ms")
        if parts_desc:
            safe_print(f"  COM 耗时: {com_elapsed:.1f}s（{', '.join(parts_desc)}）")
    else:
        safe_print(f"  COM 耗时: {com_elapsed:.1f}s")

    for line in com_output.splitlines():
        if line.startswith("STATS,"):
            parts = line.split(",", 3)
            if len(parts) >= 4:
                return parts[2], int(parts[3])

    print("错误：COM 未返回 STATS，日志:", LOG_FILE)
    sys.exit(1)


def run_openpyxl_split(
    path: Path,
    sheet_name: str,
    data_rows: list[int],
    row_groups: list[list[int]],
    last_col: int,
    append_to_existing: bool,
) -> tuple[str, int]:
    group_count = len(row_groups)
    print(
        f"  openpyxl：{len(data_rows)} 行 → {group_count} 个连续块"
        f"（Mac 模式；连续行块批量复制）"
    )

    t0 = time.time()
    wb = load_workbook(path, keep_links=False)
    if len(wb.worksheets) <= SHEET_INDEX:
        wb.close()
        print("错误：工作簿中没有可用的工作表")
        sys.exit(1)

    src = wb.worksheets[SHEET_INDEX]
    source_name = src.title

    if append_to_existing:
        if sheet_name not in wb.sheetnames:
            wb.close()
            print(f"错误：页签「{sheet_name}」不存在，无法追加")
            sys.exit(1)
        dest = wb[sheet_name]
        last = find_last_data_row(dest, DATA_START_ROW)
        dest_start = max(last + 1, DATA_START_ROW)
    else:
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]
        dest = wb.create_sheet(sheet_name)
        copy_header_layout(src, dest, HEADER_ROW, last_col)
        dest_start = DATA_START_ROW

    for gi, group in enumerate(row_groups, 1):
        start_row, end_row = group
        dest_start = copy_row_range(src, dest, start_row, end_row, dest_start, last_col)
        if group_count >= 10 and (gi % 10 == 0 or gi == group_count):
            safe_print(f"    复制进度: {gi}/{group_count} 块...")

    wb.save(path)
    wb.close()
    safe_print(f"  openpyxl 耗时: {time.time() - t0:.1f}s")
    return source_name, len(data_rows)


def list_sheet_names(path: Path) -> list[str]:
    """页签名（按 workbook.xml 中的页签顺序）；zip 直读，避免为看页签名而 load_workbook。"""
    return list_sheet_names_fast(path)


def confirm_replace_sheet(path: Path, sheet_name: str, chunk: RowChunk) -> bool:
    """首段且目标页签已存在时询问是否删除重建；后续段要求页签已存在。"""
    exists = sheet_name in list_sheet_names(path)
    if chunk.is_first:
        if not exists:
            return True
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        ok = messagebox.askyesno(
            "页签已存在",
            f"文件「{path.name}」中已有页签「{sheet_name}」。\n\n"
            f"是否删除该页签并重新复制？\n（选「否」将不做任何修改）",
            parent=root,
        )
        root.destroy()
        return ok

    if not exists:
        print()
        print(f"错误：页签「{sheet_name}」不存在。请先运行第 1 段以创建该页签。")
        return False
    return True


def run_split(path: Path, op: SplitOperation, chunk: RowChunk) -> bool | None:
    script_key = f"09-{op.key}"
    work_path, backup_path = prepare_chunk_work(path, script_key, chunk, backup_file)
    if chunk.is_first:
        print(f"  已备份: {backup_path.name}")

    data_rows, last_col = scan_matching_rows(work_path, op.scan_mode, chunk)
    row_groups = compute_row_groups(data_rows)

    if not confirm_replace_sheet(work_path, op.sheet_name, chunk):
        print()
        print("已取消：文件未修改。")
        return None

    append_mode = not chunk.is_first
    if append_mode:
        print(f"  追加模式：复制到已有页签「{op.sheet_name}」末尾")

    source_name, copied = (
        run_com_split(work_path, op.sheet_name, data_rows, row_groups, last_col, append_mode)
        if sys.platform == "win32"
        else run_openpyxl_split(
            work_path, op.sheet_name, data_rows, row_groups, last_col, append_mode
        )
    )
    verify_result(work_path, op.sheet_name, copied)

    if MANIFEST_FILE.exists():
        MANIFEST_FILE.unlink()

    finalized, out_path = finalize_chunk_output(path, script_key, chunk)

    print()
    print("=" * 50)
    title = f"{op.title} 完成 — 段 {chunk.chunk_index}/{chunk.chunk_total}"
    if finalized:
        title = f"{op.title} 全部段完成"
    print(title if finalized else f"{op.title} 本段完成（尚未输出最终文件）")
    print("=" * 50)
    if finalized and out_path:
        print(f"  已输出: {out_path}")
    else:
        print(f"  工作文件: {work_path.name}")
    print(f"  源页签: {source_name}")
    print(f"  目标页签: {op.sheet_name}")
    print(f"  条件: A～D {scan_label(op.scan_mode)}")
    print(f"  本段复制: {copied} 行")
    print(f"  前 {HEADER_ROW} 行: 格式、列宽、行高已与源页签对齐")
    if finalized:
        print(f"  请关闭 Excel 后重新打开上述文件")
    print(f"  备份: {backup_path}")
    return finalized


def verify_result(path: Path, sheet_name: str, copied: int) -> None:
    names = list_sheet_names(path)
    if sheet_name not in names:
        print()
        print("=" * 50)
        print(f"警告：保存后未找到页签「{sheet_name}」")
        print("=" * 50)
        print(f"  文件: {path.resolve()}")
        print(f"  当前全部页签: {names}")
        sys.exit(1)
    print(f"  已确认页签「{sheet_name}」存在")
    print(f"  全部页签: {names}")
    if copied == 0:
        print(f"  提示：本段未扫描到匹配行。")


def main() -> None:
    t0 = time.time()

    path = require_excel_file("选择要拆分的 Excel")
    print(f"已选文件: {path.name}")

    op = pick_operation(path.name)
    if op is None:
        print("已取消。")
        sys.exit(0)

    print(f"已选步骤: {op.title}")

    script_key = f"09-{op.key}"
    chunks = load_file_chunks(path, DATA_START_ROW)
    chunk = prompt_row_chunk("09 半自动拆分", path, DATA_START_ROW, script_key=script_key, chunks=chunks)
    if chunk is None:
        print("已取消。")
        sys.exit(0)

    while chunk is not None:
        print_chunk_banner("09 半自动拆分", chunk)
        finalized = run_split(path, op, chunk)
        if finalized is None:
            break
        if finalized:
            break
        chunk = prompt_next_chunk_after("09 半自动拆分", path, script_key, chunks, chunk, finalized)

    print(f"  总耗时: {time.time() - t0:.1f} 秒")


if __name__ == "__main__":
    main()
