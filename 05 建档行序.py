"""
建档行序脚本（代码库 05）

在进 01 之前运行一次：为第一个页签第 7 行起的数据行写入 __原序（1, 2, 3…），
并复制一份到「00-原始基准/」作为只读基准。

Windows 下使用 Excel COM 批量写入；Mac 使用 openpyxl。

弹窗选择待建档报表；写回原文件，并复制一份到「00-原始基准/」作为只读基准。

若已存在 __原序 列则拒绝重复建档，避免序号错乱。

运行前请关闭 Excel。
运行：python "05 建档行序.py"
"""

from __future__ import annotations

import shutil
import sys
import time
from pathlib import Path

from openpyxl import load_workbook

from excel_common import run_powershell
from file_picker_common import require_excel_file
from row_order_common import (
    BASELINE_DIR,
    DATA_START_ROW,
    HEADER_ROW,
    SEQ_HEADER,
    find_header_last_col,
    find_last_data_row,
    format_seq_col,
    read_header_map,
)

BASE_DIR = Path(__file__).resolve().parent
PS1_SCRIPT = BASE_DIR / "05_init_seq_com.ps1"
LOG_FILE = BASE_DIR / "05_com.log"
SHEET_INDEX = 0


def safe_print(text: str) -> None:
    try:
        print(text)
    except UnicodeEncodeError:
        print(text.encode(sys.stdout.encoding or "utf-8", errors="replace").decode(sys.stdout.encoding or "utf-8", errors="replace"))


def run_com_init(path: Path) -> tuple[str, str, int]:
    if not PS1_SCRIPT.exists():
        print(f"错误：找不到 COM 脚本: {PS1_SCRIPT}")
        sys.exit(1)

    # 命令行与旧版逐字等价：powershell -NoProfile -ExecutionPolicy Bypass -File <ps1> <extra_args>
    extra_args = [
        "-Path",
        str(path.resolve()),
        "-HeaderRow",
        str(HEADER_ROW),
        "-DataStartRow",
        str(DATA_START_ROW),
        "-SeqHeader",
        SEQ_HEADER,
        "-SheetIndex",
        str(SHEET_INDEX),
    ]

    print("  Excel COM：批量写入 __原序 列...")
    # 不传 encoding → win32 默认 gbk（与旧版一致），其余平台 utf-8。
    output = run_powershell(PS1_SCRIPT, extra_args, LOG_FILE)
    for line in output.splitlines():
        safe_print(f"  {line}")

    for line in output.splitlines():
        if line.startswith("ERROR,already_has_seq,"):
            col = line.split(",", 2)[2]
            print(f"错误：已存在「{SEQ_HEADER}」列（{col}）且已有数据，请勿重复建档。")
            print("  若顺序乱了，请运行 05 恢复行序.py。")
            sys.exit(1)
        if line.startswith("ERROR,no_data"):
            print(f"错误：第 {DATA_START_ROW} 行起没有数据（A 列为空）。")
            sys.exit(1)
        if line.startswith("ERROR,no_sheet"):
            print("错误：工作簿中没有可用的工作表")
            sys.exit(1)

    for line in output.splitlines():
        if line.startswith("RECOVER,refill_empty_seq,"):
            safe_print(f"  检测到空的 {SEQ_HEADER} 列，正在补填序号...")
        if line.startswith("STATS,"):
            _, sheet_name, seq_col, rows_str = line.split(",", 3)
            return sheet_name, seq_col, int(rows_str)

    print()
    print("错误：COM 未返回 STATS 行，日志:", LOG_FILE)
    print("  若文件里只有 __原序 表头没有数字，请关闭 Excel 后重新运行本脚本（会自动补填）。")
    sys.exit(1)


def run_openpyxl_init(path: Path) -> tuple[str, str, int]:
    print("  openpyxl：整表加载（较慢，建议 Windows 下使用 COM）...")
    wb = load_workbook(path, keep_links=False)
    if len(wb.worksheets) <= SHEET_INDEX:
        print("错误：工作簿中没有可用的工作表")
        sys.exit(1)

    ws = wb.worksheets[SHEET_INDEX]
    header_map = read_header_map(ws)
    if SEQ_HEADER in header_map:
        wb.close()
        print(f"错误：已存在「{SEQ_HEADER}」列（{format_seq_col(header_map[SEQ_HEADER])}），请勿重复建档。")
        print("  若顺序乱了，请运行 05 恢复行序.py。")
        sys.exit(1)

    last_row = find_last_data_row(ws)
    data_rows = last_row - DATA_START_ROW + 1
    if data_rows <= 0:
        wb.close()
        print(f"错误：第 {DATA_START_ROW} 行起没有数据（A 列为空）。")
        sys.exit(1)

    seq_col = find_header_last_col(ws) + 1
    ws.cell(row=HEADER_ROW, column=seq_col, value=SEQ_HEADER)
    sheet_name = ws.title
    for offset, row in enumerate(range(DATA_START_ROW, last_row + 1)):
        ws.cell(row=row, column=seq_col, value=offset + 1)

    wb.save(path)
    wb.close()
    from openpyxl.utils import get_column_letter

    return sheet_name, get_column_letter(seq_col), data_rows


def copy_baseline(path: Path) -> Path:
    BASELINE_DIR.mkdir(parents=True, exist_ok=True)
    baseline_path = BASELINE_DIR / path.name
    if baseline_path.exists():
        baseline_path.unlink()
    shutil.copy2(path, baseline_path)
    return baseline_path


def run_init() -> None:
    t0 = time.time()
    path = require_excel_file("选择待建档报表")

    print(f"读取: {path}")

    if sys.platform == "win32":
        sheet_name, seq_col, data_rows = run_com_init(path)
    else:
        sheet_name, seq_col, data_rows = run_openpyxl_init(path)

    baseline_path = copy_baseline(path)

    print()
    print("=" * 50)
    print("建档完成")
    print("=" * 50)
    print(f"  工作表: {sheet_name}")
    print(f"  {SEQ_HEADER} 列: {seq_col}（{SEQ_HEADER}）")
    print(f"  数据行: {data_rows}（序号 1～{data_rows}）")
    print(f"  已保存: {path}")
    print(f"  基准副本: {baseline_path}")
    print(f"  总耗时: {time.time() - t0:.1f} 秒")
    print()
    print("提示：可将 __原序 列隐藏；顺序乱了时运行 05 恢复行序.py。")


if __name__ == "__main__":
    run_init()
