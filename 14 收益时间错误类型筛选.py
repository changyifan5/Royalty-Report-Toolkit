"""
收益时间筛选脚本（代码库 14）

配合 02 导出的「收益时间错误汇总」，按 C/D 分组选择：
  - 勾选「忽略」→ 主报表中该组「收益时间错误」行的「比對結果」「错误类型」清空
  - 未勾选 → 保留为「聚焦」，需人工修改

弹窗可选汇总文件 `{主文件名}_收益时间错误汇总.xlsx`，或直接选主报表（同目录需有汇总）。
处理完成后写回主报表，并导出 `{主文件名}_收益时间聚焦.xlsx`（待处理分组清单）。

运行前请关闭 Excel。运行：python "14 收益时间筛选.py"
"""

from __future__ import annotations

import importlib.util
import shutil
import sys
import time
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import ttk

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

from file_picker_common import confirm_process_file, pick_excel_file

BASE_DIR = Path(__file__).resolve().parent
SHEET_INDEX = 0
HEADER_ROW = 6
DATA_START_ROW = 7
HEADER_SCAN_LIMIT = 500

COL_ARTIST = 3
COL_ALBUM = 4

CMP_RESULT_HEADER = "比對結果"
ERROR_TYPE_HEADER = "错误类型"

SUMMARY_SUFFIX = "_收益时间错误汇总"
FOCUS_SUFFIX = "_收益时间聚焦"
REVENUE_TIME_ERROR = "收益时间错误"


@dataclass
class SummaryGroup:
    artist: str
    album: str
    row_count: int
    sum_g: float
    artist_header: str = "C列"
    album_header: str = "D列"
    g_header: str = "G列"


def load_script_module(filename: str):
    path = BASE_DIR / filename
    module_name = path.stem.replace(" ", "_")
    spec = importlib.util.spec_from_file_location(module_name, path)
    if spec is None or spec.loader is None:
        print(f"错误：无法加载 {filename}")
        sys.exit(1)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = mod
    spec.loader.exec_module(mod)
    return mod


def cell_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return str(value).strip()


def to_number(value) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    try:
        return float(text)
    except ValueError:
        return 0.0


def read_header_map(sheet, header_row: int = HEADER_ROW) -> dict[str, int]:
    mapping: dict[str, int] = {}
    for col in range(1, HEADER_SCAN_LIMIT + 1):
        val = sheet.cell(row=header_row, column=col).value
        if val is not None and str(val).strip():
            mapping[str(val).strip()] = col
    return mapping


def resolve_paths(picked: Path) -> tuple[Path, Path | None]:
    stem = picked.stem
    if stem.endswith(SUMMARY_SUFFIX):
        main_stem = stem[: -len(SUMMARY_SUFFIX)]
        main_path = picked.with_name(f"{main_stem}{picked.suffix}")
        return main_path, picked
    main_path = picked
    summary_path = picked.with_name(f"{picked.stem}{SUMMARY_SUFFIX}{picked.suffix}")
    return main_path, summary_path if summary_path.exists() else None


def load_groups_from_summary(summary_path: Path) -> list[SummaryGroup]:
    wb = load_workbook(summary_path, read_only=True, data_only=True, keep_links=False)
    ws = wb.active
    headers = [ws.cell(row=1, column=c).value for c in range(1, 5)]
    artist_h = str(headers[0]).strip() if headers[0] else "C列"
    album_h = str(headers[1]).strip() if headers[1] else "D列"
    g_h = str(headers[3]).strip() if len(headers) >= 4 and headers[3] else "G列"

    groups: list[SummaryGroup] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        artist = cell_text(row[0])
        if artist == "合计":
            break
        album = cell_text(row[1] if len(row) > 1 else None)
        groups.append(
            SummaryGroup(
                artist=artist,
                album=album,
                row_count=int(to_number(row[2] if len(row) > 2 else 0)),
                sum_g=to_number(row[3] if len(row) > 3 else 0),
                artist_header=artist_h,
                album_header=album_h,
                g_header=g_h,
            )
        )
    wb.close()
    return groups


def load_groups_from_main(main_path: Path) -> list[SummaryGroup]:
    mod02 = load_script_module("02 半自动除错.py")
    raw_groups, labels = mod02.collect_revenue_time_groups(main_path)
    return [
        SummaryGroup(
            artist=g.artist,
            album=g.album,
            row_count=g.row_count,
            sum_g=g.sum_g,
            artist_header=labels.get("artist", "C列"),
            album_header=labels.get("album", "D列"),
            g_header=labels.get("g", "G列"),
        )
        for g in raw_groups
    ]


def prompt_ignore_selection(groups: list[SummaryGroup]) -> set[tuple[str, str]] | None:
    root = tk.Tk()
    root.title("收益时间筛选 — 选择要忽略的分组")
    root.attributes("-topmost", True)
    root.geometry("820x520")

    chosen: dict[str, set[tuple[str, str]] | None] = {"ignore": None}

    frame = ttk.Frame(root, padding=10)
    frame.pack(fill=tk.BOTH, expand=True)

    ttk.Label(
        frame,
        text="勾选「忽略」的分组将不再需要处理（主报表中对应行的比對結果会被清空）。\n"
        "未勾选的分组为「聚焦」，会写入聚焦清单供你后续修改。",
        wraplength=780,
    ).pack(anchor="w", pady=(0, 8))

    canvas = tk.Canvas(frame, highlightthickness=0)
    scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=canvas.yview)
    inner = ttk.Frame(canvas)
    inner.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=inner, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    vars_map: dict[tuple[str, str], tk.BooleanVar] = {}
    for idx, group in enumerate(groups):
        key = (group.artist, group.album)
        var = tk.BooleanVar(value=False)
        vars_map[key] = var
        text = (
            f"{group.artist}  |  {group.album}  |  "
            f"{group.row_count} 行  |  {group.g_header}合计 {group.sum_g:g}"
        )
        ttk.Checkbutton(inner, text=text, variable=var).grid(row=idx, column=0, sticky="w", pady=2)

    btn_frame = ttk.Frame(root, padding=(10, 0, 10, 10))
    btn_frame.pack(fill=tk.X)

    def set_all(ignore: bool) -> None:
        for var in vars_map.values():
            var.set(ignore)

    def on_ok() -> None:
        chosen["ignore"] = {key for key, var in vars_map.items() if var.get()}
        root.destroy()

    def on_cancel() -> None:
        chosen["ignore"] = None
        root.destroy()

    ttk.Button(btn_frame, text="全部忽略", command=lambda: set_all(True)).pack(side=tk.LEFT, padx=(0, 6))
    ttk.Button(btn_frame, text="全部聚焦", command=lambda: set_all(False)).pack(side=tk.LEFT, padx=(0, 6))
    ttk.Button(btn_frame, text="确定", command=on_ok, width=10).pack(side=tk.RIGHT, padx=(6, 0))
    ttk.Button(btn_frame, text="取消", command=on_cancel, width=10).pack(side=tk.RIGHT)

    root.protocol("WM_DELETE_WINDOW", on_cancel)
    root.mainloop()
    return chosen["ignore"]


def export_focus_list(main_path: Path, focus_groups: list[SummaryGroup]) -> Path | None:
    if not focus_groups:
        return None
    out_path = main_path.with_name(f"{main_path.stem}{FOCUS_SUFFIX}{main_path.suffix}")
    if out_path.exists():
        out_path.unlink()

    sample = focus_groups[0]
    wb = Workbook()
    ws = wb.active
    ws.title = "收益时间聚焦"
    headers = [sample.artist_header, sample.album_header, "错误行数", f"{sample.g_header}合计", "处理状态"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    total_rows = 0
    total_g = 0.0
    for group in focus_groups:
        ws.append([group.artist, group.album, group.row_count, group.sum_g, "待处理"])
        total_rows += group.row_count
        total_g += group.sum_g

    total_row = ws.max_row + 1
    ws.cell(row=total_row, column=1, value="合计")
    ws.cell(row=total_row, column=3, value=total_rows)
    ws.cell(row=total_row, column=4, value=total_g)
    for col in range(1, 5):
        ws.cell(row=total_row, column=col).font = Font(bold=True)

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 12

    wb.save(out_path)
    wb.close()
    return out_path


def backup_file(path: Path) -> Path:
    backup_path = path.with_suffix(path.suffix + ".bak")
    shutil.copy2(path, backup_path)
    return backup_path


def apply_ignore_rows(
    ws,
    row_start: int,
    row_end: int,
    cmp_col: int,
    error_type_col: int | None,
    ignore_keys: set[tuple[str, str]],
    mod02,
) -> int:
    cleared = 0
    for row in range(row_start, row_end + 1):
        cmp_val = ws.cell(row=row, column=cmp_col).value
        if mod02.is_cmp_ignored(cmp_val):
            continue
        if mod02.classify_error(str(cmp_val)) != REVENUE_TIME_ERROR:
            continue

        artist = cell_text(ws.cell(row=row, column=COL_ARTIST).value)
        album = cell_text(ws.cell(row=row, column=COL_ALBUM).value)
        if (artist, album) not in ignore_keys:
            continue

        ws.cell(row=row, column=cmp_col, value=None)
        if error_type_col is not None:
            ws.cell(row=row, column=error_type_col, value=None)
        cleared += 1
    return cleared


def run_filter():
    t0 = time.time()
    picked = pick_excel_file("选择收益时间汇总文件或主报表")
    if picked is None:
        print("已取消。")
        sys.exit(0)

    main_path, summary_path = resolve_paths(picked)
    if not main_path.exists():
        print(f"错误：找不到主报表: {main_path}")
        sys.exit(1)

    if summary_path and summary_path.exists():
        print(f"汇总文件: {summary_path.name}")
        groups = load_groups_from_summary(summary_path)
    else:
        print(f"未找到汇总文件，从主报表扫描: {main_path.name}")
        groups = load_groups_from_main(main_path)

    if not groups:
        print("没有可筛选的「收益时间错误」分组。")
        sys.exit(0)

    print(f"主报表: {main_path.name}")
    print(f"共 {len(groups)} 个 C/D 分组")

    ignore_keys = prompt_ignore_selection(groups)
    if ignore_keys is None:
        print("已取消。")
        sys.exit(0)

    focus_groups = [g for g in groups if (g.artist, g.album) not in ignore_keys]
    if not ignore_keys and not focus_groups:
        print("未选择任何操作。")
        sys.exit(0)

    print()
    print(f"  忽略分组: {len(ignore_keys)} 个")
    print(f"  聚焦分组: {len(focus_groups)} 个")

    if not ignore_keys:
        focus_path = export_focus_list(main_path, focus_groups)
        print()
        print("=" * 50)
        print("未选择忽略分组，主报表未修改")
        print("=" * 50)
        if focus_path:
            print(f"  已导出聚焦清单: {focus_path}")
        print(f"  总耗时: {time.time() - t0:.1f} 秒")
        return

    mod02 = load_script_module("02 半自动除错.py")
    row_start, row_end = confirm_process_file("14 收益时间筛选", main_path, data_start_row=DATA_START_ROW)

    backup_path = backup_file(main_path)
    print(f"  已备份: {backup_path.name}")

    wb = load_workbook(main_path)
    ws = wb.worksheets[SHEET_INDEX]
    header_map = read_header_map(ws, HEADER_ROW)
    if CMP_RESULT_HEADER not in header_map:
        wb.close()
        print(f"错误：第 {HEADER_ROW} 行找不到列「{CMP_RESULT_HEADER}」")
        sys.exit(1)

    cmp_col = header_map[CMP_RESULT_HEADER]
    error_type_col = header_map.get(ERROR_TYPE_HEADER)

    cleared = apply_ignore_rows(ws, row_start, row_end, cmp_col, error_type_col, ignore_keys, mod02)
    wb.save(main_path)
    wb.close()

    print()
    print("=" * 50)
    print("筛选完成")
    print("=" * 50)
    print(f"  清空比對結果: {cleared} 行")

    focus_path = export_focus_list(main_path, focus_groups)
    print(f"  已保存主报表: {main_path}")
    if focus_path:
        print(f"  聚焦清单: {focus_path.name}")
    print(f"  备份: {backup_path.name}")

    print(f"  总耗时: {time.time() - t0:.1f} 秒")
    print()
    print("提示：忽略的分组已清空比對結果；聚焦清单中的分组请人工修改。")


if __name__ == "__main__":
    from launcher_common import run_with_pause

    run_with_pause(run_filter)
