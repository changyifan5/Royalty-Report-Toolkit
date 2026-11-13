"""
B 列筛选 G 列汇总脚本（代码库 13）

扫描文件夹内 Excel（仅第一个页签）：
  1. 弹窗选文件夹
  2. 输入 B 列筛选内容（与单元格去首尾空格后完全一致）
  3. 每个文件一行：I 列为匹配行的日期范围，K 列用于区分，并加总 G 列
  4. 生成汇总表 xlsx（保存在所选文件夹）

I 列为日期，输出该文件匹配行涉及的最小～最大日期；K 列取首条匹配行的值。
第 6 行为表头，第 7 行起为数据。Mac / Windows 均使用 openpyxl。
"""

from __future__ import annotations

import calendar
import re
import sys
import time
import tkinter as tk
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from tkinter import messagebox, simpledialog

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from file_picker_common import EXCEL_SUFFIXES, pick_folder
from row_order_common import DATA_START_ROW, HEADER_ROW, find_last_data_row

BASE_DIR = Path(__file__).resolve().parent
SHEET_INDEX = 0
COL_B = 2
COL_G = 7
COL_I = 9
COL_K = 11
SCAN_MAX_COL = COL_K

SKIP_STEM_PARTS = (".chunkwork", ".列映射.work", ".bak")

EXCEL_DATE_ORIGIN = datetime(1899, 12, 30)
EXCEL_SERIAL_MIN = 30000
EXCEL_SERIAL_MAX = 60000
DATE_DISPLAY = "%Y/%m/%d"


@dataclass
class FileSummary:
    source_name: str
    col_k: str = ""
    sum_g: float = 0.0
    row_count: int = 0
    i_dates: list[date] = field(default_factory=list)
    i_raw_fallback: list[str] = field(default_factory=list)

    @property
    def col_i(self) -> str:
        if self.i_dates:
            return format_date_range(self.i_dates)
        if self.i_raw_fallback:
            seen: list[str] = []
            for text in self.i_raw_fallback:
                if text and text not in seen:
                    seen.append(text)
            return "、".join(seen)
        return ""


def safe_print(text: str) -> None:
    try:
        print(text)
    except UnicodeEncodeError:
        print(
            text.encode(sys.stdout.encoding or "utf-8", errors="replace").decode(
                sys.stdout.encoding or "utf-8", errors="replace"
            )
        )


def ask_filter_text() -> str | None:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    value = simpledialog.askstring(
        "B 列筛选",
        "请输入 B 列要匹配的内容：\n（与单元格去首尾空格后完全一致）",
        parent=root,
    )
    root.destroy()
    if value is None:
        return None
    text = value.strip()
    return text or None


def normalize_match_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return str(value).strip()


def to_number(value: object) -> float:
    if value is None:
        return 0.0
    if isinstance(value, bool):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def int_to_year_month(value: int) -> tuple[int, int] | None:
    if 190001 <= value <= 209912:
        year, month = divmod(value, 100)
        if 1900 <= year <= 2099 and 1 <= month <= 12:
            return year, month
    return None


def excel_serial_to_date(value: float | int) -> date | None:
    serial = int(value)
    if EXCEL_SERIAL_MIN <= serial <= EXCEL_SERIAL_MAX:
        return (EXCEL_DATE_ORIGIN + timedelta(days=serial)).date()
    return None


def month_end(year: int, month: int) -> date:
    return date(year, month, calendar.monthrange(year, month)[1])


def parse_i_date(value: object) -> date | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        number = float(value)
        if number == int(number):
            as_int = int(number)
            ym = int_to_year_month(as_int)
            if ym:
                return month_end(*ym)
            return excel_serial_to_date(as_int)
    if isinstance(value, str):
        text = value.strip()
        if len(text) == 6 and text.isdigit():
            ym = int_to_year_month(int(text))
            if ym:
                return month_end(*ym)
        digits = re.sub(r"\D", "", text)
        if len(digits) >= 6:
            ym = int_to_year_month(int(digits[:6]))
            if ym:
                return month_end(*ym)
        for fmt, size in (("%Y/%m/%d", 10), ("%Y-%m-%d", 10), ("%Y%m%d", 8)):
            try:
                return datetime.strptime(text[:size], fmt).date()
            except ValueError:
                continue
    return None


def format_date_range(dates: list[date]) -> str:
    if not dates:
        return ""
    lo, hi = min(dates), max(dates)
    if lo == hi:
        return lo.strftime(DATE_DISPLAY)
    return f"{lo.strftime(DATE_DISPLAY)}～{hi.strftime(DATE_DISPLAY)}"


def cell_display(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    if isinstance(value, datetime):
        return value.strftime(DATE_DISPLAY)
    if isinstance(value, date):
        return value.strftime(DATE_DISPLAY)
    return str(value).strip()


def safe_filename_part(text: str, limit: int = 24) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|\s]+', "_", text.strip())
    cleaned = cleaned.strip("_") or "筛选"
    return cleaned[:limit]


def list_excel_files(folder: Path) -> list[Path]:
    results: list[Path] = []
    for path in sorted(folder.iterdir(), key=lambda p: p.name.lower()):
        if not path.is_file():
            continue
        if path.suffix.lower() not in EXCEL_SUFFIXES:
            continue
        if path.name.startswith("~$"):
            continue
        stem_lower = path.stem.lower()
        if any(part in stem_lower for part in SKIP_STEM_PARTS):
            continue
        if stem_lower.startswith("汇总_b列筛选"):
            continue
        results.append(path)
    return results


def read_header_labels(ws) -> tuple[str, str]:
    i_header = cell_display(ws.cell(row=HEADER_ROW, column=COL_I).value) or "I列"
    k_header = cell_display(ws.cell(row=HEADER_ROW, column=COL_K).value) or "K列"
    return i_header, k_header


def process_workbook(path: Path, filter_b: str) -> tuple[FileSummary | None, str, str, int]:
    wb = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    ws = wb.worksheets[SHEET_INDEX]
    i_header, k_header = read_header_labels(ws)
    last_row = find_last_data_row(ws, DATA_START_ROW)
    scanned = 0
    summary: FileSummary | None = None

    if last_row >= DATA_START_ROW:
        for row in ws.iter_rows(
            min_row=DATA_START_ROW,
            max_row=last_row,
            min_col=1,
            max_col=SCAN_MAX_COL,
            values_only=True,
        ):
            scanned += 1
            b_val = row[COL_B - 1] if len(row) >= COL_B else None
            if normalize_match_text(b_val) != filter_b:
                continue

            i_raw = row[COL_I - 1] if len(row) >= COL_I else None
            k_val = cell_display(row[COL_K - 1] if len(row) >= COL_K else None)
            g_val = to_number(row[COL_G - 1] if len(row) >= COL_G else None)

            if summary is None:
                summary = FileSummary(source_name=path.name, col_k=k_val)
            elif not summary.col_k and k_val:
                summary.col_k = k_val

            parsed_i = parse_i_date(i_raw)
            if parsed_i is not None:
                summary.i_dates.append(parsed_i)
            else:
                text_i = cell_display(i_raw)
                if text_i:
                    summary.i_raw_fallback.append(text_i)

            summary.sum_g += g_val
            summary.row_count += 1

    wb.close()
    return summary, i_header, k_header, scanned


def write_summary(
    output_path: Path,
    filter_b: str,
    folder: Path,
    rows: list[FileSummary],
    i_header: str,
    k_header: str,
) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "汇总"

    meta = [
        ("B 列筛选", filter_b),
        ("来源文件夹", str(folder)),
        ("生成时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("说明", "每个文件一行；I 列为匹配行的日期范围，K 列用于区分文件"),
    ]
    for idx, (label, value) in enumerate(meta, start=1):
        ws.cell(row=idx, column=1, value=label)
        ws.cell(row=idx, column=2, value=value)

    header_row = len(meta) + 2
    headers = ["来源文件", f"{i_header}（日期范围）", k_header, "G列合计", "匹配行数"]
    for col, title in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col, value=title)
        cell.font = Font(bold=True)

    data_row = header_row + 1
    grand_sum = 0.0
    grand_rows = 0
    for item in rows:
        ws.cell(row=data_row, column=1, value=item.source_name)
        ws.cell(row=data_row, column=2, value=item.col_i)
        ws.cell(row=data_row, column=3, value=item.col_k)
        ws.cell(row=data_row, column=4, value=item.sum_g)
        ws.cell(row=data_row, column=5, value=item.row_count)
        grand_sum += item.sum_g
        grand_rows += item.row_count
        data_row += 1

    total_row = data_row
    ws.cell(row=total_row, column=1, value="合计").font = Font(bold=True)
    ws.cell(row=total_row, column=4, value=grand_sum).font = Font(bold=True)
    ws.cell(row=total_row, column=5, value=grand_rows).font = Font(bold=True)

    for col in range(1, 6):
        letter = get_column_letter(col)
        max_len = max(len(str(ws.cell(row=r, column=col).value or "")) for r in range(1, total_row + 1))
        ws.column_dimensions[letter].width = min(max(max_len + 2, 10), 48)

    wb.save(output_path)
    wb.close()


def run_aggregate() -> None:
    t0 = time.time()
    print("请选择包含 Excel 的文件夹...")
    folder = pick_folder("选择要汇总的 Excel 文件夹")
    if folder is None:
        print("已取消。")
        sys.exit(0)
    safe_print(f"  文件夹: {folder}")

    filter_b = ask_filter_text()
    if filter_b is None:
        print("已取消。")
        sys.exit(0)
    safe_print(f"  B 列筛选: {filter_b!r}")

    files = list_excel_files(folder)
    if not files:
        print("错误：文件夹内没有可处理的 Excel 文件。")
        sys.exit(1)
    safe_print(f"  待处理文件: {len(files)} 个")

    summary_rows: list[FileSummary] = []
    i_header = "I列"
    k_header = "K列"
    total_matched_rows = 0

    for path in files:
        safe_print(f"  读取: {path.name}")
        try:
            summary, file_i_header, file_k_header, scanned = process_workbook(path, filter_b)
        except Exception as exc:
            safe_print(f"    跳过（无法读取）: {exc}")
            continue

        if file_i_header:
            i_header = file_i_header
        if file_k_header:
            k_header = file_k_header

        if summary is None:
            safe_print(f"    无匹配（扫描 {scanned} 行）")
            continue

        summary_rows.append(summary)
        total_matched_rows += summary.row_count
        safe_print(
            f"    I={summary.col_i!r} K={summary.col_k!r} → G合计={summary.sum_g:g}（{summary.row_count} 行）"
        )

    if not summary_rows:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        messagebox.showinfo(
            "无匹配数据",
            f"文件夹内 {len(files)} 个文件中，B 列均没有匹配 {filter_b!r} 的行。",
            parent=root,
        )
        root.destroy()
        print("未生成汇总表。")
        sys.exit(0)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_name = f"汇总_B列筛选_{safe_filename_part(filter_b)}_{stamp}.xlsx"
    output_path = folder / output_name
    write_summary(output_path, filter_b, folder, summary_rows, i_header, k_header)

    print()
    print("=" * 50)
    print("汇总完成")
    print("=" * 50)
    print(f"  B 列筛选: {filter_b!r}")
    print(f"  扫描文件: {len(files)} 个，有匹配: {len(summary_rows)} 个")
    print(f"  匹配数据行: {total_matched_rows}")
    print(f"  已保存: {output_path}")
    print(f"  总耗时: {time.time() - t0:.1f} 秒")


if __name__ == "__main__":
    run_aggregate()
