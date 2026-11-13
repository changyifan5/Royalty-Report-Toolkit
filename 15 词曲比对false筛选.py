"""
词曲比对筛选脚本（代码库 15）

扫描主报表「词曲比对」列为 false 的行，按 C/D 分组汇总并导出：
  `{主文件名}_词曲比对false汇总.xlsx`（页签 1 汇总 + 页签 2 明细，含源文件行号、A/B/C/D/G）
  `{主文件名}_词曲比对false聚焦.xlsx`（待处理分组清单，E 列默认「待处理」）

主报表不修改；看完聚焦清单后跑 16 回写。

弹窗可选汇总文件或主报表。运行前请关闭 Excel。
运行：python "15 词曲比对筛选.py"
"""

from __future__ import annotations

import sys
import time
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

from file_picker_common import pick_excel_file
from row_order_common import find_last_data_row

BASE_DIR = Path(__file__).resolve().parent
SHEET_INDEX = 0
HEADER_ROW = 6
DATA_START_ROW = 7
HEADER_SCAN_LIMIT = 500

COL_KEY = 1
COL_SONG = 2
COL_ARTIST = 3
COL_ALBUM = 4
COL_G = 7

LYRICS_CMP_HEADER = "词曲比对"
SUMMARY_SUFFIX = "_词曲比对false汇总"
FOCUS_SUFFIX = "_词曲比对false聚焦"


@dataclass
class SummaryGroup:
    artist: str
    album: str
    row_count: int
    sum_g: float
    artist_header: str = "C列"
    album_header: str = "D列"
    g_header: str = "G列"


@dataclass
class FalseDetail:
    source_row: int
    col_a: str
    song: str
    artist: str
    album: str
    g_value: float
    cmp_value: str


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


def is_cmp_false(value) -> bool:
    if value is None:
        return False
    if isinstance(value, bool):
        return not value
    if isinstance(value, (int, float)):
        return value == 0
    text = str(value).strip().lower()
    return text in ("false", "0", "否", "no")


def read_header_map(sheet, header_row: int = HEADER_ROW) -> dict[str, int]:
    mapping: dict[str, int] = {}
    for col in range(1, HEADER_SCAN_LIMIT + 1):
        val = sheet.cell(row=header_row, column=col).value
        if val is not None and str(val).strip():
            mapping[str(val).strip()] = col
    return mapping


def header_label(sheet, col: int, default: str) -> str:
    val = sheet.cell(row=HEADER_ROW, column=col).value
    if val is None or not str(val).strip():
        return default
    return str(val).strip()


def collect_false_data(path: Path) -> tuple[list[SummaryGroup], dict[str, str], list[FalseDetail]]:
    wb = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    if len(wb.worksheets) <= SHEET_INDEX:
        wb.close()
        return [], {}, []

    ws = wb.worksheets[SHEET_INDEX]
    header_map = read_header_map(ws, HEADER_ROW)
    cmp_col = header_map.get(LYRICS_CMP_HEADER)
    if cmp_col is None:
        wb.close()
        print(f"错误：第 {HEADER_ROW} 行找不到列「{LYRICS_CMP_HEADER}」")
        sys.exit(1)

    labels = {
        "key": header_label(ws, COL_KEY, "A列"),
        "song": header_label(ws, COL_SONG, "B列"),
        "artist": header_label(ws, COL_ARTIST, "C列"),
        "album": header_label(ws, COL_ALBUM, "D列"),
        "g": header_label(ws, COL_G, "G列"),
    }

    groups: dict[tuple[str, str], SummaryGroup] = {}
    details: list[FalseDetail] = []
    last_row = find_last_data_row(ws, DATA_START_ROW)

    for source_row, row in enumerate(
        ws.iter_rows(min_row=DATA_START_ROW, max_row=last_row, values_only=True),
        start=DATA_START_ROW,
    ):
        if not row:
            continue

        cmp_idx = cmp_col - 1
        cmp_val = row[cmp_idx] if cmp_idx < len(row) else None
        if not is_cmp_false(cmp_val):
            continue

        col_a = cell_text(row[COL_KEY - 1] if len(row) >= COL_KEY else None)
        song = cell_text(row[COL_SONG - 1] if len(row) >= COL_SONG else None)
        artist = cell_text(row[COL_ARTIST - 1] if len(row) >= COL_ARTIST else None)
        album = cell_text(row[COL_ALBUM - 1] if len(row) >= COL_ALBUM else None)
        g_val = to_number(row[COL_G - 1] if len(row) >= COL_G else None)

        details.append(
            FalseDetail(
                source_row=source_row,
                col_a=col_a,
                song=song,
                artist=artist,
                album=album,
                g_value=g_val,
                cmp_value=cell_text(cmp_val),
            )
        )

        key = (artist, album)
        if key not in groups:
            groups[key] = SummaryGroup(
                artist=artist,
                album=album,
                row_count=0,
                sum_g=0.0,
                artist_header=labels["artist"],
                album_header=labels["album"],
                g_header=labels["g"],
            )
        bucket = groups[key]
        bucket.row_count += 1
        bucket.sum_g += g_val

    wb.close()
    ordered_groups = sorted(
        groups.values(),
        key=lambda g: (-g.row_count, g.artist, g.album),
    )
    ordered_details = sorted(details, key=lambda d: d.source_row)
    return ordered_groups, labels, ordered_details


def export_summary(main_path: Path, groups: list[SummaryGroup], labels: dict[str, str], details: list[FalseDetail]) -> Path | None:
    if not groups:
        return None

    out_path = main_path.with_name(f"{main_path.stem}{SUMMARY_SUFFIX}{main_path.suffix}")
    if out_path.exists():
        out_path.unlink()

    wb = Workbook()
    ws = wb.active
    ws.title = "词曲比对false汇总"

    headers = [labels["artist"], labels["album"], "false行数", f'{labels["g"]}合计']
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    total_rows = 0
    total_g = 0.0
    for group in groups:
        ws.append([group.artist, group.album, group.row_count, group.sum_g])
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

    ws_detail = wb.create_sheet("明细")
    detail_headers = [
        "源文件行号",
        labels["key"],
        labels["song"],
        labels["artist"],
        labels["album"],
        labels["g"],
        LYRICS_CMP_HEADER,
    ]
    ws_detail.append(detail_headers)
    for cell in ws_detail[1]:
        cell.font = Font(bold=True)

    for item in details:
        ws_detail.append(
            [
                item.source_row,
                item.col_a,
                item.song,
                item.artist,
                item.album,
                item.g_value,
                item.cmp_value,
            ]
        )

    ws_detail.column_dimensions["A"].width = 12
    ws_detail.column_dimensions["B"].width = 20
    ws_detail.column_dimensions["C"].width = 28
    ws_detail.column_dimensions["D"].width = 24
    ws_detail.column_dimensions["E"].width = 28
    ws_detail.column_dimensions["F"].width = 14
    ws_detail.column_dimensions["G"].width = 12

    wb.save(out_path)
    wb.close()
    return out_path


def resolve_paths(picked: Path) -> tuple[Path, Path | None, bool]:
    """返回 (主报表, 汇总文件或 None, 是否直接选了汇总)。"""
    stem = picked.stem
    if stem.endswith(SUMMARY_SUFFIX):
        main_stem = stem[: -len(SUMMARY_SUFFIX)]
        main_path = picked.with_name(f"{main_stem}{picked.suffix}")
        return main_path, picked, True
    main_path = picked
    summary_path = picked.with_name(f"{picked.stem}{SUMMARY_SUFFIX}{picked.suffix}")
    return main_path, summary_path if summary_path.exists() else None, False


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
        groups.append(
            SummaryGroup(
                artist=artist,
                album=cell_text(row[1] if len(row) > 1 else None),
                row_count=int(to_number(row[2] if len(row) > 2 else 0)),
                sum_g=to_number(row[3] if len(row) > 3 else 0),
                artist_header=artist_h,
                album_header=album_h,
                g_header=g_h,
            )
        )
    wb.close()
    return groups


def export_focus_list(main_path: Path, focus_groups: list[SummaryGroup]) -> Path | None:
    if not focus_groups:
        return None
    out_path = main_path.with_name(f"{main_path.stem}{FOCUS_SUFFIX}{main_path.suffix}")
    if out_path.exists():
        out_path.unlink()

    sample = focus_groups[0]
    wb = Workbook()
    ws = wb.active
    ws.title = "词曲比对false聚焦"
    headers = [sample.artist_header, sample.album_header, "false行数", f"{sample.g_header}合计", "处理状态"]
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


def run_filter():
    t0 = time.time()
    picked = pick_excel_file("选择词曲比对汇总文件或主报表")
    if picked is None:
        print("已取消。")
        sys.exit(0)

    main_path, summary_path, picked_is_summary = resolve_paths(picked)
    if not main_path.exists():
        print(f"错误：找不到主报表: {main_path}")
        sys.exit(1)

    print(f"主报表: {main_path.name}")

    if not picked_is_summary:
        print("  扫描 false 行并导出汇总…")
        groups, labels, details = collect_false_data(main_path)
        summary_exported = export_summary(main_path, groups, labels, details)
        if summary_exported:
            print(f"  已导出汇总: {summary_exported.name}")
    elif summary_path and summary_path.exists():
        print(f"汇总文件: {summary_path.name}")
        groups = load_groups_from_summary(summary_path)
    else:
        groups = []

    if not groups:
        print("没有「词曲比对」为 false 的分组。")
        sys.exit(0)

    print(f"共 {len(groups)} 个 C/D 分组")

    focus_path = export_focus_list(main_path, groups)
    print()
    print("=" * 50)
    print("导出完成（主报表未修改）")
    print("=" * 50)
    if focus_path:
        print(f"  聚焦清单: {focus_path.name}")
    print(f"  总耗时: {time.time() - t0:.1f} 秒")
    print()
    print("提示：在聚焦清单 E 列删掉已看完分组的「待处理」后，跑 16 回写主报表。")


if __name__ == "__main__":
    from launcher_common import run_with_pause

    run_with_pause(run_filter)
