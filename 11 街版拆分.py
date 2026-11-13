"""
街版拆分脚本（代码库 11）

与 07 类似，但比例来自表内列而非手填比例：
  1. A～D 列橙色字体的行（第 7 行起）
  2. 表尾新增两列，按「词曲派歌 / (词曲派歌 + 词曲示例版权方)」拆分 G 列（写入 Excel 公式）
  3. 新建页签「街版」：前 6 行格式/列宽/行高与源页签一致；橙色行从第 7 行起粘贴
  4. 「街版」页签 G 列 = 源页签新增第二列的数值（非公式）

表头列名支持繁体：詞曲派歌、詞曲示例版權方（亦认简体）。

Windows 使用 Excel COM 批量写入；Mac 使用 openpyxl。
运行：弹窗选 Excel。运行前请关闭 Excel。
"""

from __future__ import annotations

import json
import sys
import time
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, simpledialog

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from excel_common import EXCEL_SUFFIXES, backup_file, run_powershell, safe_print
from excel_font_color import is_orange_font
from excel_sheet_copy import copy_header_layout, copy_row_range
from file_picker_common import require_excel_file
from row_order_common import find_header_last_col, find_last_data_row, read_header_map

BASE_DIR = Path(__file__).resolve().parent
PS1_SCRIPT = BASE_DIR / "11_split_street_com.ps1"
MANIFEST_FILE = BASE_DIR / "11_street_manifest.json"
LOG_FILE = BASE_DIR / "11_com.log"

SOURCE_COL = 7  # G
HEADER_ROW = 6
DATA_START_ROW = 7
SHEET_INDEX = 0
COLOR_CHECK_COLS = 4
STREET_SHEET = "街版"

PAIGE_HEADER_NAMES = ("词曲派歌", "詞曲派歌")
JIE_HEADER_NAMES = ("词曲示例版权方", "詞曲示例版權方")
DEFAULT_HEADER1 = "詞曲派歌拆分"
DEFAULT_HEADER2 = "詞曲示例版權方拆分"

INTEGER_LIKE_FORMATS = frozenset({"0", "#,##0", "0_ ", "#,##0_ ", "0;-0;;@", "#,##0;-#,##0;;@"})


def ask_string(title: str, prompt: str, default: str = "") -> str | None:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    value = simpledialog.askstring(title, prompt, initialvalue=default, parent=root)
    root.destroy()
    return value.strip() if value else None


def row_has_orange_font(row_cells) -> bool:
    return any(is_orange_font(cell.font) for cell in row_cells)


def scan_orange_rows(path: Path, data_start: int, data_end: int) -> list[int]:
    print(f"  扫描 A～D 橙色字体（只读，第 {data_start}～{data_end} 行）...")
    t0 = time.time()
    matched: list[int] = []
    wb = load_workbook(path, read_only=True, keep_links=False)
    ws = wb.worksheets[SHEET_INDEX]
    scanned = 0
    for row_cells in ws.iter_rows(
        min_row=data_start,
        max_row=data_end,
        min_col=1,
        max_col=COLOR_CHECK_COLS,
    ):
        scanned += 1
        if scanned % 50000 == 0:
            safe_print(f"    已扫描 {scanned} 行，命中 {len(matched)} 行...")
        if not any(c.value is not None and str(c.value).strip() for c in row_cells):
            continue
        if row_has_orange_font(row_cells):
            matched.append(row_cells[0].row)
    wb.close()
    safe_print(f"  扫描完成: {scanned} 行，命中 {len(matched)} 行，耗时 {time.time() - t0:.1f}s")
    return matched


def find_named_column(header_map: dict[str, int], names: tuple[str, ...], label: str) -> int:
    for name in names:
        if name in header_map:
            return header_map[name]
    print(f"错误：第 {HEADER_ROW} 行未找到「{label}」列。")
    print(f"  已识别表头: {list(header_map.keys())}")
    print(f"  期望列名（任一）: {list(names)}")
    sys.exit(1)


def ensure_output_columns(
    sheet, header_map: dict[str, int], name1: str, name2: str
) -> tuple[int, int]:
    if name1 in header_map and name2 in header_map:
        return header_map[name1], header_map[name2]
    if name1 in header_map or name2 in header_map:
        print("错误：两个输出表头只能同时已存在或同时新建，当前只存在一个。")
        sys.exit(1)
    last = find_header_last_col(sheet)
    col1, col2 = last + 1, last + 2
    sheet.cell(row=HEADER_ROW, column=col1, value=name1)
    sheet.cell(row=HEADER_ROW, column=col2, value=name2)
    return col1, col2


def to_number(value: object) -> float | None:
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


def output_number_format_from_g(fmt: str | None) -> str:
    fmt = (fmt or "General").strip()
    if fmt in INTEGER_LIKE_FORMATS:
        return "0.##########"
    return fmt


def split_formula(
    row: int,
    g_col: int,
    pai_col: int,
    jie_col: int,
    out_col1: int,
) -> tuple[str, str]:
    g = get_column_letter(g_col)
    pai = get_column_letter(pai_col)
    jie = get_column_letter(jie_col)
    c1 = get_column_letter(out_col1)
    denom = f"({pai}{row}+{jie}{row})"
    f1 = f'=IF({denom}=0,"",{g}{row}*{pai}{row}/{denom})'
    f2 = f'=IF({denom}=0,"",{g}{row}-{c1}{row})'
    return f1, f2


def calc_street_g_value(g_raw: object, pai_raw: object, jie_raw: object) -> float | None:
    """第二列数值：G × 词曲示例版权方 / (词曲派歌 + 词曲示例版权方)。"""
    g = to_number(g_raw)
    if g is None:
        return None
    pai = to_number(pai_raw) or 0.0
    jie = to_number(jie_raw) or 0.0
    denom = pai + jie
    if denom == 0:
        return None
    return round(g * jie / denom, 10)


def confirm_replace_street_sheet(path: Path) -> bool:
    wb = load_workbook(path, read_only=True, keep_links=False)
    exists = STREET_SHEET in wb.sheetnames
    wb.close()
    if not exists:
        return True
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    ok = messagebox.askyesno(
        "页签已存在",
        f"文件「{path.name}」中已有页签「{STREET_SHEET}」。\n\n"
        f"是否删除该页签并重新生成？\n（选「否」将不做任何修改）",
        parent=root,
    )
    root.destroy()
    return ok


def read_header_map_from_file(path: Path) -> dict[str, int]:
    wb = load_workbook(path, read_only=True, keep_links=False)
    header_map = read_header_map(wb.worksheets[SHEET_INDEX])
    wb.close()
    return header_map


def write_manifest(
    path: Path,
    orange_rows: list[int],
    pai_col: int,
    jie_col: int,
    header1: str,
    header2: str,
) -> None:
    payload = {
        "path": str(path.resolve()),
        "sheet_index": SHEET_INDEX,
        "header_row": HEADER_ROW,
        "data_start_row": DATA_START_ROW,
        "source_col": SOURCE_COL,
        "street_sheet": STREET_SHEET,
        "pai_col": pai_col,
        "jie_col": jie_col,
        "header1": header1,
        "header2": header2,
        "orange_rows": orange_rows,
    }
    MANIFEST_FILE.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")


def run_com_apply(
    path: Path,
    orange_rows: list[int],
    pai_col: int,
    jie_col: int,
    header1: str,
    header2: str,
) -> tuple[str, int, int, int, int]:
    if not PS1_SCRIPT.exists():
        print(f"错误：找不到 COM 脚本: {PS1_SCRIPT}")
        sys.exit(1)

    write_manifest(path, orange_rows, pai_col, jie_col, header1, header2)

    print(f"  Excel COM：写入公式 + 生成「{STREET_SHEET}」页签（{len(orange_rows)} 行）...")
    t0 = time.time()
    # 11 的 COM 输出编码为 gbk（win32），run_powershell 默认编码即 gbk，不要显式传 encoding。
    com_output = run_powershell(
        PS1_SCRIPT,
        ["-ManifestPath", str(MANIFEST_FILE.resolve())],
        LOG_FILE,
    )
    for line in com_output.splitlines():
        safe_print(f"  {line}")
    safe_print(f"  COM 耗时: {time.time() - t0:.1f}s")

    for line in com_output.splitlines():
        if line.startswith("STATS,"):
            parts = line.split(",", 5)
            if len(parts) >= 6:
                sheet_name = parts[1]
                out_col1 = int(parts[2])
                out_col2 = int(parts[3])
                formula_rows = int(parts[4])
                pasted = int(parts[5])
                return sheet_name, out_col1, out_col2, formula_rows, pasted

    print("错误：COM 未返回 STATS，日志:", LOG_FILE)
    sys.exit(1)


def apply_openpyxl(
    path: Path,
    orange_rows: list[int],
    pai_col: int,
    jie_col: int,
    header1: str,
    header2: str,
) -> tuple[str, int, int, int, int]:
    print("  openpyxl：写入公式 + 生成「街版」页签...")
    t0 = time.time()
    wb = load_workbook(path, keep_links=False)
    if len(wb.worksheets) <= SHEET_INDEX:
        wb.close()
        print("错误：工作簿中没有可用的工作表")
        sys.exit(1)

    ws = wb.worksheets[SHEET_INDEX]
    sheet_name = ws.title
    header_map = read_header_map(ws)
    out_col1, out_col2 = ensure_output_columns(ws, header_map, header1, header2)

    formula_rows = 0
    for row in orange_rows:
        f1, f2 = split_formula(row, SOURCE_COL, pai_col, jie_col, out_col1)
        out_fmt = output_number_format_from_g(ws.cell(row=row, column=SOURCE_COL).number_format)
        c1 = ws.cell(row=row, column=out_col1, value=f1)
        c2 = ws.cell(row=row, column=out_col2, value=f2)
        c1.number_format = out_fmt
        c2.number_format = out_fmt
        formula_rows += 1

    last_col = find_header_last_col(ws)

    if STREET_SHEET in wb.sheetnames:
        del wb[STREET_SHEET]
    street_ws = wb.create_sheet(STREET_SHEET)
    copy_header_layout(ws, street_ws, HEADER_ROW, last_col)

    dest_row = DATA_START_ROW
    pasted = 0
    for row in orange_rows:
        copy_row_range(ws, street_ws, row, row, dest_row, last_col)
        street_val = calc_street_g_value(
            ws.cell(row=row, column=SOURCE_COL).value,
            ws.cell(row=row, column=pai_col).value,
            ws.cell(row=row, column=jie_col).value,
        )
        g_cell = street_ws.cell(row=dest_row, column=SOURCE_COL)
        g_cell.value = street_val
        g_cell.number_format = output_number_format_from_g(
            ws.cell(row=row, column=SOURCE_COL).number_format
        )
        dest_row += 1
        pasted += 1

    wb.save(path)
    wb.close()
    safe_print(f"  openpyxl 耗时: {time.time() - t0:.1f}s")
    return sheet_name, out_col1, out_col2, formula_rows, pasted


def print_summary(
    *,
    sheet_name: str,
    orange_count: int,
    pai_col: int,
    jie_col: int,
    header1: str,
    header2: str,
    out_col1: int,
    out_col2: int,
    formula_rows: int,
    pasted: int,
    path: Path,
    backup_path: Path,
    elapsed: float,
) -> None:
    print()
    print("=" * 50)
    print("街版拆分完成")
    print("=" * 50)
    print(f"  源页签: {sheet_name}")
    print(f"  橙色行: {orange_count} 行")
    print(f"  比例列: {get_column_letter(pai_col)}=词曲派歌, {get_column_letter(jie_col)}=词曲示例版权方")
    print(f"  G 列: {get_column_letter(SOURCE_COL)}")
    print(
        f"  新增列（公式）: {get_column_letter(out_col1)}={header1!r}, "
        f"{get_column_letter(out_col2)}={header2!r}"
    )
    print(f"  已写公式行: {formula_rows} 行")
    print(f"  新页签: {STREET_SHEET!r}（{pasted} 行数据，G 列为示例版权方拆分数值）")
    print(f"  已保存: {path}")
    print(f"  备份: {backup_path}")
    print(f"  总耗时: {elapsed:.1f} 秒")


def run_split() -> None:
    t0 = time.time()
    path = require_excel_file("选择要街版拆分的 Excel")
    print(f"  文件: {path.name}")

    if not confirm_replace_street_sheet(path):
        print("已取消：文件未修改。")
        sys.exit(0)

    backup_path = backup_file(path)
    print(f"  已备份: {backup_path.name}")

    wb_probe = load_workbook(path, read_only=True, keep_links=False)
    ws_probe = wb_probe.worksheets[SHEET_INDEX]
    last_row = find_last_data_row(ws_probe, DATA_START_ROW)
    wb_probe.close()
    if last_row < DATA_START_ROW:
        print(f"错误：第 {DATA_START_ROW} 行起没有数据。")
        sys.exit(1)

    orange_rows = scan_orange_rows(path, DATA_START_ROW, last_row)
    if not orange_rows:
        print("错误：未找到 A～D 橙色字体的数据行。")
        sys.exit(1)

    header1 = ask_string("表头 1", "第一列输出表头名：", DEFAULT_HEADER1)
    if not header1:
        print("已取消。")
        sys.exit(0)
    header2 = ask_string("表头 2", "第二列输出表头名（街版 G 列取此列数值）：", DEFAULT_HEADER2)
    if not header2:
        print("已取消。")
        sys.exit(0)
    if header1 == header2:
        print("错误：两列表头名不能相同。")
        sys.exit(1)

    header_map = read_header_map_from_file(path)
    pai_col = find_named_column(header_map, PAIGE_HEADER_NAMES, "词曲派歌")
    jie_col = find_named_column(header_map, JIE_HEADER_NAMES, "词曲示例版权方")

    if sys.platform == "win32":
        sheet_name, out_col1, out_col2, formula_rows, pasted = run_com_apply(
            path, orange_rows, pai_col, jie_col, header1, header2
        )
    else:
        sheet_name, out_col1, out_col2, formula_rows, pasted = apply_openpyxl(
            path, orange_rows, pai_col, jie_col, header1, header2
        )

    if MANIFEST_FILE.exists():
        MANIFEST_FILE.unlink()

    print_summary(
        sheet_name=sheet_name,
        orange_count=len(orange_rows),
        pai_col=pai_col,
        jie_col=jie_col,
        header1=header1,
        header2=header2,
        out_col1=out_col1,
        out_col2=out_col2,
        formula_rows=formula_rows,
        pasted=pasted,
        path=path,
        backup_path=backup_path,
        elapsed=time.time() - t0,
    )


if __name__ == "__main__":
    run_split()
