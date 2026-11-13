"""
列映射填表脚本（代码库 10）

桌面上的中介报表（文件名：XXX平台XXX类型报表-时间.xlsx）为列结构模板；
弹窗多选各源 Excel，按对应关系填入中介文件。

流程：
  1. 弹窗选择中介文件（XXX平台XXX类型报表-时间.xlsx）
  2. 弹窗多选要处理的中介文件页签（默认选中第一个）
  3. 多选源 Excel 文件
  4. 对第一个页签：对每个源文件弹窗选择「源文件页签」（默认第一个页签，可切换，切换后自动
     重新检测表头），再指定「源列 → 中介列」（源文件默认第 1 行表头，可改；中介文件默认第 6 行）；
     后续页签若表头与第一个完全相同，可复用同一套列映射或单独配置
  5. 每个中介页签独立询问写入位置；写完一个源页签后可继续追加同一源文件的下一个页签
  6. 备份中介文件为 .bak 后写回（也可另存为）；所有页签写入同一工作文件，最后统一保存一次

写入规则：A 列以文本写入（避免科学计数法）；I 列统一为 yyyy/mm/dd，并取该月最后一天
         （如源为 202307 或 Excel 日期序列号，均转为 2023/07/31）。

中介文件默认第 6 行表头（与其他报表脚本一致）。运行时多选页签处理范围。运行前请关闭 Excel。
运行：python "10 列映射填表.py"
"""

from __future__ import annotations

import os
import re
import calendar
import shutil
import sys
import time
import traceback
from datetime import date, datetime, timedelta


class _StderrFilter:
    """过滤 tkinter 弹窗触发的 libpng iCCP 警告，不影响真实报错。"""

    def __init__(self, stream) -> None:
        self._stream = stream

    def write(self, msg: str) -> None:
        if "libpng warning" in msg or "iCCP" in msg:
            return
        self._stream.write(msg)

    def flush(self) -> None:
        self._stream.flush()

    def __getattr__(self, name: str):
        return getattr(self._stream, name)


sys.stderr = _StderrFilter(sys.stderr)

import tkinter as tk
from dataclasses import dataclass, field
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from excel_common import list_sheet_names_fast, tk_mono_font, tk_ui_font
from file_picker_common import get_initial_dir, pick_excel_file, pick_excel_files
from row_order_common import HEADER_ROW, find_last_data_row

# ---------------------------------------------------------------------------
# 可配置参数
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent


def get_desktop_dir() -> Path:
    profile = Path.home()
    for name in ("Desktop", "桌面"):
        candidate = profile / name
        if candidate.is_dir():
            return candidate
    userprofile = os.environ.get("USERPROFILE", "")
    if userprofile:
        for name in ("Desktop", "桌面"):
            candidate = Path(userprofile) / name
            if candidate.is_dir():
                return candidate
    return profile / "Desktop"


DESKTOP_DIR = get_desktop_dir()

SHEET_INDEX = 0
INTERMEDIATE_HEADER_ROW = HEADER_ROW
DEFAULT_SOURCE_HEADER_ROW = 1
HEADER_SCAN_LIMIT = 500

# 表头自动匹配：归一化后相等，或互为别名
HEADER_ALIASES: dict[str, tuple[str, ...]] = {
    "歌曲名": ("歌曲名称", "歌曲名", "曲名", "track", "trackname"),
    "歌曲名称": ("歌曲名", "曲名", "track"),
    "track": ("track", "歌曲名", "歌曲名称", "曲名"),
    "演唱者": ("歌手", "艺人", "artist", "artistname"),
    "artist": ("artist", "演唱者", "歌手", "艺人"),
    "专辑名": ("专辑", "album", "albumname"),
    "album": ("album", "专辑", "专辑名"),
    "isrc/upc": ("isrc", "upc", "isrc/upc", "isrc / upc"),
    "isrc / upc": ("isrc/upc", "isrc", "upc"),
    "type": ("type", "类型"),
    "quantity": ("quantity", "数量", "播放量"),
    "revenue": ("revenue", "收益", "报表收益"),
    "report revenue": ("report revenue", "reportrevenue", "报表收益(未税)"),
    "date": ("date", "日期", "结算月份"),
    "country": ("country", "国家", "地区"),
    "retailer": ("retailer", "平台", "渠道"),
}

SOURCE_HEADER_SCAN_ROWS = 30

INTERMEDIATE_NAME_TEMPLATE = "XXX平台XXX类型报表-时间.xlsx"
# 文件名格式：{平台名}平台{类型名}类型报表-{时间}.xlsx
INTERMEDIATE_NAME_PATTERN = re.compile(r"^.+平台.+类型报表-.+$")

EXCEL_FILETYPES = [("Excel 文件", "*.xlsx;*.xlsm;*.xltx;*.xltm"), ("所有文件", "*.*")]
SKIP_VALUE = "（不填）"
FIXED_VALUE = "（填固定值）"

COL_A = 1
COL_I = 9
TEXT_FORMAT = "@"
I_DATE_FORMAT = "yyyy/mm/dd"
EXCEL_DATE_ORIGIN = datetime(1899, 12, 30)
EXCEL_SERIAL_MIN = 30000
EXCEL_SERIAL_MAX = 60000


@dataclass(frozen=True)
class ColumnInfo:
    index: int
    letter: str
    header: str
    display: str


@dataclass
class FileMapping:
    source_path: Path
    header_row: int
    pairs: list[tuple[int, int]]  # (source_col_idx, target_col_idx)
    sheet: int | str = SHEET_INDEX  # 源文件页签（序号或名称，默认第一个）
    fixed_values: dict[int, str] = field(default_factory=dict)


def safe_print(text: str) -> None:
    try:
        print(text)
    except UnicodeEncodeError:
        print(
            text.encode(sys.stdout.encoding or "utf-8", errors="replace").decode(
                sys.stdout.encoding or "utf-8", errors="replace"
            )
        )


def is_intermediate_name(name: str) -> bool:
    return bool(INTERMEDIATE_NAME_PATTERN.match(Path(name).stem))


def find_intermediate_files() -> list[Path]:
    if not DESKTOP_DIR.is_dir():
        return []
    results: list[Path] = []
    candidates: list[Path] = [DESKTOP_DIR]
    for sub in DESKTOP_DIR.iterdir():
        if sub.is_dir() and not sub.name.startswith("."):
            candidates.append(sub)
    for folder in candidates:
        for f in folder.iterdir():
            if (
                f.is_file()
                and f.suffix.lower() == ".xlsx"
                and not f.name.startswith("~$")
                and is_intermediate_name(f.name)
            ):
                results.append(f)
    return sorted(set(results), key=lambda p: p.name.lower())


def pick_intermediate_fallback() -> Path | None:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    messagebox.showinfo(
        "未找到中介文件",
        f"桌面上未找到符合命名规则的中介文件：\n  {INTERMEDIATE_NAME_TEMPLATE}\n\n请手动选择。",
        parent=root,
    )
    path = filedialog.askopenfilename(
        title=f"选择中介文件（{INTERMEDIATE_NAME_TEMPLATE}）",
        initialdir=str(get_initial_dir()),
        filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
    )
    root.destroy()
    if not path:
        return None
    picked = Path(path)
    if not is_intermediate_name(picked.name):
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        messagebox.showwarning(
            "文件名不符",
            f"所选文件不符合中介文件命名：\n  {INTERMEDIATE_NAME_TEMPLATE}",
            parent=root,
        )
        root.destroy()
        return None
    return picked


def pick_save_path(default: Path) -> Path | None:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.asksaveasfilename(
        title="保存合并结果（默认写回中介文件）",
        initialdir=str(default.parent),
        initialfile=default.name,
        defaultextension=".xlsx",
        filetypes=EXCEL_FILETYPES,
    )
    root.destroy()
    return Path(path) if path else None


def backup_file(path: Path) -> Path:
    backup_path = path.with_suffix(path.suffix + ".bak")
    if backup_path.exists():
        backup_path.unlink()
    shutil.copy2(path, backup_path)
    return backup_path


def header_text(value: object | None, col_idx: int) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return text if text else get_column_letter(col_idx)


def resolve_sheet(wb, sheet: int | str):
    """按页签序号（int）或页签名（str）取工作表；默认第一个页签。"""
    if isinstance(sheet, str):
        return wb[sheet]
    return wb.worksheets[sheet]


def read_columns(path: Path, header_row: int, sheet: int | str = SHEET_INDEX) -> list[ColumnInfo]:
    """读取中介文件指定页签的表头列；sheet 可为页签序号（默认 0）或页签名。"""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = resolve_sheet(wb, sheet)
        header_cells = next(
            ws.iter_rows(
                min_row=header_row,
                max_row=header_row,
                max_col=HEADER_SCAN_LIMIT,
                values_only=True,
            ),
            (),
        )
        columns: list[ColumnInfo] = []
        for col_idx, raw in enumerate(header_cells, start=1):
            text = header_text(raw, col_idx)
            letter = get_column_letter(col_idx)
            display = f"{letter} · {text}" if text != letter else letter
            columns.append(ColumnInfo(index=col_idx, letter=letter, header=text, display=display))
        while columns and columns[-1].header == get_column_letter(columns[-1].index):
            last_val = header_cells[columns[-1].index - 1] if columns[-1].index - 1 < len(header_cells) else None
            if last_val is None or str(last_val).strip() == "":
                columns.pop()
            else:
                break
        return columns
    finally:
        wb.close()


def normalize_header_key(name: str) -> str:
    return re.sub(r"\s+", "", str(name).strip().lower())


def headers_match(target_name: str, source_name: str) -> bool:
    t = normalize_header_key(target_name)
    s = normalize_header_key(source_name)
    if not t or not s:
        return False
    if t == s:
        return True
    if t in s or s in t:
        return True
    for canonical, aliases in HEADER_ALIASES.items():
        group = {normalize_header_key(canonical), *(normalize_header_key(a) for a in aliases)}
        if t in group and s in group:
            return True
    return False


def header_signature(columns: list[ColumnInfo]) -> list[str]:
    """把一列 ColumnInfo 归一为表头签名：列数 + 每列归一后表头文字（保持顺序）。"""
    return [normalize_header_key(col.header) for col in columns]


def sheet_headers_identical(cols_a: list[ColumnInfo], cols_b: list[ColumnInfo]) -> bool:
    """两个页签表头是否完全一致：列数相同且逐列归一后表头文字相同。

    用于判断后续页签能否安全复用第一页签的列映射（目标列序号按列对齐）。
    """
    return header_signature(cols_a) == header_signature(cols_b)


def auto_pick_source(target: ColumnInfo, source_columns: list[ColumnInfo]) -> str:
    target_name = target.header.strip()
    if not target_name:
        return SKIP_VALUE
    for src in source_columns:
        if headers_match(target_name, src.header):
            return src.display
    return SKIP_VALUE


def detect_source_header_row(
    source_path: Path, target_columns: list[ColumnInfo], sheet: int | str = SHEET_INDEX
) -> int:
    """在前若干行里找与中介表头最匹配的一行（网易云等导出常在第 6 行）。"""
    target_names = [col.header.strip() for col in target_columns if col.header.strip()]
    if not target_names:
        return DEFAULT_SOURCE_HEADER_ROW

    wb = load_workbook(source_path, read_only=True, data_only=True)
    try:
        ws = resolve_sheet(wb, sheet)
        best_row = DEFAULT_SOURCE_HEADER_ROW
        best_score = 0
        for row in range(1, SOURCE_HEADER_SCAN_ROWS + 1):
            header_cells = next(
                ws.iter_rows(
                    min_row=row,
                    max_row=row,
                    max_col=HEADER_SCAN_LIMIT,
                    values_only=True,
                ),
                (),
            )
            score = 0
            for raw in header_cells:
                if raw is None:
                    continue
                text = str(raw).strip()
                if not text:
                    continue
                for target_name in target_names:
                    if headers_match(target_name, text):
                        score += 1
            if score > best_score:
                best_score = score
                best_row = row
        return best_row if best_score >= 2 else DEFAULT_SOURCE_HEADER_ROW
    finally:
        wb.close()


def pick_write_start_row(header_row: int, last_data_row: int) -> int | None:
    data_start = header_row + 1
    if last_data_row < data_start:
        return data_start

    append_row = last_data_row + 1
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    choice = messagebox.askyesnocancel(
        "选择写入位置",
        f"中介文件第 {data_start}～{last_data_row} 行已有内容。\n\n"
        f"【是】追加到第 {append_row} 行起（在文件末尾，需向下滚动查看）\n"
        f"【否】从第 {data_start} 行起覆盖写入（常用：填入空模板或替换示例数据）\n"
        f"【取消】退出",
        parent=root,
        icon="question",
    )
    root.destroy()
    if choice is None:
        return None
    return append_row if choice else data_start


def iter_source_data_rows(source_path: Path, mapping: FileMapping):
    src_cols = [src for src, _tgt in mapping.pairs]
    if not src_cols and not mapping.fixed_values:
        return
    wb = load_workbook(source_path, data_only=True)
    try:
        ws = resolve_sheet(wb, mapping.sheet)
        for row in ws.iter_rows(
            min_row=mapping.header_row + 1,
            max_row=ws.max_row,
            max_col=max(src_cols, default=ws.max_column),
            values_only=True,
        ):
            has_data = row_has_mapped_data(row, src_cols) if src_cols else any(
                value is not None and str(value).strip() for value in row
            )
            if has_data:
                yield row
    finally:
        wb.close()


def count_source_data_rows(source_path: Path, mapping: FileMapping) -> int:
    return sum(1 for _row in iter_source_data_rows(source_path, mapping))


def describe_mapping(mapping: FileMapping, target_columns: list[ColumnInfo]) -> str:
    target_by_idx = {col.index: col for col in target_columns}
    lines: list[str] = [f"  源页签: {mapping.sheet}"]
    for src_col, tgt_col in mapping.pairs:
        tgt = target_by_idx.get(tgt_col)
        tgt_label = f"{tgt.letter}·{tgt.header}" if tgt else get_column_letter(tgt_col)
        lines.append(f"  源 {get_column_letter(src_col)} → 中介 {tgt_label}")
    for tgt_col, value in mapping.fixed_values.items():
        tgt = target_by_idx.get(tgt_col)
        tgt_label = f"{tgt.letter}·{tgt.header}" if tgt else get_column_letter(tgt_col)
        lines.append(f"  固定值 {value!r} → 中介 {tgt_label}")
    return "\n".join(lines) if lines else "  （未映射任何列）"


def find_last_data_row_in_cols(sheet, data_start_row: int, cols: list[int]) -> int:
    if not cols:
        return find_last_data_row(sheet, data_start_row=data_start_row)

    if getattr(sheet.parent, "read_only", False):
        last = data_start_row - 1
        max_col = max(cols)
        for idx, row in enumerate(
            sheet.iter_rows(min_row=data_start_row, max_col=max_col, values_only=True)
        ):
            if any(
                col - 1 < len(row)
                and row[col - 1] is not None
                and str(row[col - 1]).strip()
                for col in cols
            ):
                last = data_start_row + idx
        return max(last, data_start_row - 1)

    max_row = sheet.max_row
    if max_row < data_start_row:
        return data_start_row - 1
    for row in range(max_row, data_start_row - 1, -1):
        for col in cols:
            val = sheet.cell(row=row, column=col).value
            if val is not None and str(val).strip():
                return row
    return data_start_row - 1


def row_has_mapped_data(row: tuple[object, ...], src_cols: list[int]) -> bool:
    for col in src_cols:
        idx = col - 1
        if idx < 0 or idx >= len(row):
            continue
        val = row[idx]
        if val is not None and str(val).strip():
            return True
    return False


def find_column_by_display(columns: list[ColumnInfo], display: str) -> ColumnInfo | None:
    for col in columns:
        if col.display == display:
            return col
    return None


class PickIntermediateDialog:
    def __init__(self, files: list[Path]) -> None:
        self.files = files
        self.result: Path | None = None
        self.header_row = INTERMEDIATE_HEADER_ROW

        self.root = tk.Tk()
        self.root.title("中介文件")
        self.root.attributes("-topmost", True)
        self.root.resizable(False, False)

        frame = ttk.Frame(self.root, padding=12)
        frame.grid(row=0, column=0, sticky="nsew")

        ttk.Label(
            frame,
            text=f"桌面中介文件（{INTERMEDIATE_NAME_TEMPLATE}）：",
            font=("Microsoft YaHei UI", 10),
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))

        self.listbox = tk.Listbox(frame, width=72, height=min(12, max(4, len(files))), font=("Consolas", 10))
        self.listbox.grid(row=1, column=0, columnspan=2, sticky="ew")
        for path in files:
            self.listbox.insert(tk.END, path.name)
        self.listbox.selection_set(0)
        if len(files) == 1:
            self.listbox.configure(state="disabled")

        ttk.Label(frame, text="中介文件表头行：").grid(row=2, column=0, sticky="w", pady=(10, 0))
        self.header_var = tk.StringVar(value=str(INTERMEDIATE_HEADER_ROW))
        ttk.Spinbox(frame, from_=1, to=999, width=8, textvariable=self.header_var).grid(
            row=2, column=1, sticky="w", pady=(10, 0)
        )

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=(12, 0), sticky="e")
        ttk.Button(btn_frame, text="确定", command=self.confirm, width=10).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(btn_frame, text="取消", command=self.cancel, width=10).grid(row=0, column=1)

        self.root.bind("<Return>", lambda _e: self.confirm())
        self.root.bind("<Escape>", lambda _e: self.cancel())

    def confirm(self) -> None:
        sel = self.listbox.curselection()
        if not sel:
            messagebox.showwarning("提示", "请先选择一个中介文件。", parent=self.root)
            return
        try:
            row = int(self.header_var.get())
            if row < 1:
                raise ValueError
        except ValueError:
            messagebox.showwarning("提示", "表头行必须是大于 0 的整数。", parent=self.root)
            return
        self.result = self.files[sel[0]]
        self.header_row = row
        self.root.destroy()

    def cancel(self) -> None:
        self.root.destroy()

    def run(self) -> tuple[Path, int] | None:
        self.root.mainloop()
        if self.result is None:
            return None
        return self.result, self.header_row


def resolve_intermediate() -> tuple[Path, int] | None:
    picked = pick_excel_file(
        f"选择中介文件（{INTERMEDIATE_NAME_TEMPLATE}）",
        initial_dir=get_initial_dir(),
    )
    if picked is None:
        return None
    safe_print(f"  已选中介文件: {picked.name}")
    picker = PickIntermediateDialog([picked])
    return picker.run()


def list_intermediate_sheet_names(path: Path) -> list[str]:
    """读取中介文件全部页签名（zip 直读，不加载单元格）。"""
    return list_sheet_names_fast(path)


def list_source_sheet_names(path: Path) -> list[str]:
    """读取源文件全部页签名；zip 直读失败时回退 openpyxl。"""
    try:
        names = list_sheet_names_fast(path)
        if names:
            return names
    except Exception:
        pass
    wb = load_workbook(path, read_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def pick_intermediate_sheets(sheet_names: list[str]) -> list[str] | None:
    """弹窗多选要处理的页签；默认选中第一个。

    返回按文件顺序排列的选中页签名列表；用户取消或未选任何项时返回 None。
    仅有一个页签时不弹窗，直接返回 [那一个页签]。
    """
    if not sheet_names:
        return None
    if len(sheet_names) == 1:
        safe_print(f"  页签（仅一个）: {sheet_names[0]}")
        return [sheet_names[0]]

    selected: list[str] = []

    root = tk.Tk()
    root.title("选择要处理的页签（可多选）")
    root.attributes("-topmost", True)
    root.resizable(False, False)

    frame = ttk.Frame(root, padding=12)
    frame.grid(row=0, column=0, sticky="nsew")

    ttk.Label(
        frame,
        text="选择要写入的中介页签。点击可切换选中状态。",
        font=tk_ui_font(10),
    ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))

    listbox = tk.Listbox(
        frame,
        width=56,
        height=min(14, max(4, len(sheet_names))),
        font=tk_mono_font(10),
        selectmode=tk.MULTIPLE,
    )
    listbox.grid(row=1, column=0, columnspan=2, sticky="ew")
    for name in sheet_names:
        listbox.insert(tk.END, name)
    listbox.selection_set(0)

    def select_first() -> None:
        listbox.selection_clear(0, tk.END)
        listbox.selection_set(0)

    def confirm() -> None:
        sel = listbox.curselection()  # 升序返回，即文件中的先后顺序
        if not sel:
            messagebox.showwarning("提示", "请至少选择一个页签。", parent=root)
            return
        selected[:] = [sheet_names[i] for i in sel]
        root.destroy()

    def cancel() -> None:
        root.destroy()

    btn = ttk.Frame(frame)
    btn.grid(row=2, column=0, columnspan=2, pady=(10, 0), sticky="e")
    ttk.Button(btn, text="只选第一个", command=select_first).grid(row=0, column=0, padx=(0, 6))
    ttk.Button(btn, text="全选", command=lambda: listbox.selection_set(0, tk.END)).grid(row=0, column=1, padx=(0, 6))
    ttk.Button(btn, text="确定", command=confirm, width=10).grid(row=0, column=2, padx=(0, 6))
    ttk.Button(btn, text="取消", command=cancel, width=10).grid(row=0, column=3)

    root.bind("<Return>", lambda _e: confirm())
    root.bind("<Escape>", lambda _e: cancel())
    root.mainloop()

    return selected or None


def ask_reuse_mapping(sheet_name: str) -> bool:
    """后续页签表头与第一个相同时，询问是否复用同一套列映射。是=复用，否=单独配置。"""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    ok = messagebox.askyesno(
        "复用列映射",
        f"页签「{sheet_name}」的表头与第一个页签完全相同。\n\n"
        f"是否复用第一页签的同一套列映射？\n"
        f"【是】复用  【否】为该页签单独配置",
        parent=root,
    )
    root.destroy()
    return bool(ok)


def ask_continue_next_sheet(next_name: str) -> bool:
    """一个页签写完后，询问是否继续处理下一个选中页签。是=继续，否=停止。"""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    ok = messagebox.askyesno(
        "继续下一个页签",
        f"本页签已完成。\n\n是否继续处理下一个页签「{next_name}」？\n"
        f"【是】继续  【否】停止并输出已完成的页签",
        parent=root,
    )
    root.destroy()
    return bool(ok)


def ask_continue_source_sheet(source_path: Path, next_sheet: str, target_sheet: str) -> bool:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        return messagebox.askyesno(
            "继续下一个源页签",
            f"要把源文件「{source_path.name}」的下一个页签「{next_sheet}」"
            f"也追加到中介页签「{target_sheet}」吗？\n\n"
            "选「否」会保留已写入的数据，并跳过这个源文件剩余的页签。",
            parent=root,
        )
    finally:
        root.destroy()


class MappingDialog:
    def __init__(
        self,
        source_path: Path,
        target_columns: list[ColumnInfo],
        default_header_row: int = DEFAULT_SOURCE_HEADER_ROW,
        source_sheet: str | None = None,
    ) -> None:
        self.source_path = source_path
        self.target_columns = target_columns
        self.result: FileMapping | None = None
        self.skipped = False
        self.cancelled = False

        self.sheet_names = list_source_sheet_names(source_path)
        if not self.sheet_names:
            raise ValueError(f"源文件没有可用页签: {source_path.name}")

        initial_sheet = source_sheet if source_sheet is not None else self.sheet_names[0]
        detected_header = detect_source_header_row(source_path, target_columns, sheet=initial_sheet)

        self.root = tk.Tk()
        self.root.title(f"列映射 — {source_path.name}")
        self.root.attributes("-topmost", True)
        self.root.geometry("760x520")
        self.root.minsize(640, 420)

        outer = ttk.Frame(self.root, padding=12)
        outer.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(4, weight=1)

        ttk.Label(
            outer,
            text=f"源文件：{source_path.name}",
            font=("Microsoft YaHei UI", 10, "bold"),
        ).grid(row=0, column=0, sticky="w")

        sheet_row = ttk.Frame(outer)
        sheet_row.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        ttk.Label(sheet_row, text="源文件页签：").grid(row=0, column=0, sticky="w")
        self.sheet_var = tk.StringVar(value=initial_sheet)
        sheet_combo = ttk.Combobox(
            sheet_row,
            textvariable=self.sheet_var,
            values=self.sheet_names,
            state="readonly" if source_sheet is None else "disabled",
            width=40,
        )
        sheet_combo.grid(row=0, column=1, padx=(4, 8), sticky="w")
        ttk.Label(
            sheet_row,
            text="（切换后自动重新检测表头）" if source_sheet is None else "（正在配置下一个源页签）",
            foreground="#666",
        ).grid(row=0, column=2, sticky="w")
        sheet_combo.bind("<<ComboboxSelected>>", lambda _e: self.on_sheet_changed())

        top = ttk.Frame(outer)
        top.grid(row=2, column=0, sticky="ew", pady=(8, 8))
        ttk.Label(top, text="源文件表头行：").grid(row=0, column=0, sticky="w")
        self.header_var = tk.StringVar(value=str(detected_header))
        ttk.Spinbox(top, from_=1, to=999, width=8, textvariable=self.header_var).grid(row=0, column=1, padx=(4, 12))
        ttk.Button(top, text="刷新列名", command=self.reload_source_columns).grid(row=0, column=2)
        ttk.Label(
            top,
            text=f"（已自动检测，常见为 1 或 6；网易云类导出多为 6）",
            foreground="#666",
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 0))

        ttk.Label(
            outer,
            text="为每个中介列选择源列、填固定值或不填（同名表头已自动匹配）。",
        ).grid(row=3, column=0, sticky="nw")

        canvas = tk.Canvas(outer, highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        self.scroll_frame = ttk.Frame(canvas)
        self.scroll_frame.bind(
            "<Configure>",
            lambda _e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=4, column=0, sticky="nsew")
        scrollbar.grid(row=4, column=1, sticky="ns")

        self.source_columns: list[ColumnInfo] = []
        self.combo_vars: list[tk.StringVar] = []
        self.fixed_vars: list[tk.StringVar] = []
        self.reload_source_columns()

        btn_frame = ttk.Frame(outer)
        btn_frame.grid(row=5, column=0, columnspan=2, pady=(12, 0), sticky="e")
        ttk.Button(btn_frame, text="确定", command=self.confirm, width=10).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(btn_frame, text="跳过此文件", command=self.skip, width=12).grid(row=0, column=1, padx=(0, 6))
        ttk.Button(btn_frame, text="取消全部", command=self.cancel_all, width=10).grid(row=0, column=2)

        self.root.bind("<Escape>", lambda _e: self.cancel_all())

    def on_sheet_changed(self) -> None:
        """切换源文件页签后：重新检测表头行并刷新列名。"""
        try:
            detected = detect_source_header_row(
                self.source_path, self.target_columns, sheet=self.sheet_var.get()
            )
        except Exception:
            detected = DEFAULT_SOURCE_HEADER_ROW
        self.header_var.set(str(detected))
        self.reload_source_columns()

    def reload_source_columns(self) -> None:
        try:
            header_row = int(self.header_var.get())
            if header_row < 1:
                raise ValueError
        except ValueError:
            messagebox.showwarning("提示", "表头行必须是大于 0 的整数。", parent=self.root)
            return

        try:
            self.source_columns = read_columns(
                self.source_path, header_row, sheet=self.sheet_var.get()
            )
        except Exception as exc:
            messagebox.showerror("错误", f"读取源文件列失败：\n{exc}", parent=self.root)
            return

        saved_fixed = {
            target.index: value.get()
            for target, choice, value in zip(self.target_columns, self.combo_vars, self.fixed_vars)
            if choice.get() == FIXED_VALUE
        }
        for child in self.scroll_frame.winfo_children():
            child.destroy()
        self.combo_vars.clear()
        self.fixed_vars.clear()

        values = [SKIP_VALUE, FIXED_VALUE] + [col.display for col in self.source_columns]
        ttk.Label(self.scroll_frame, text="中介列", font=("Microsoft YaHei UI", 9, "bold")).grid(
            row=0, column=0, sticky="w", padx=(0, 16), pady=(0, 6)
        )
        ttk.Label(self.scroll_frame, text="取值方式", font=("Microsoft YaHei UI", 9, "bold")).grid(
            row=0, column=1, sticky="w", pady=(0, 6)
        )
        ttk.Label(self.scroll_frame, text="固定值").grid(row=0, column=2, sticky="w", padx=(8, 0))

        for row_idx, target in enumerate(self.target_columns, start=1):
            ttk.Label(
                self.scroll_frame,
                text=f"{target.letter} · {target.header}",
                font=("Consolas", 10),
            ).grid(row=row_idx, column=0, sticky="w", padx=(0, 16), pady=2)

            choice = FIXED_VALUE if target.index in saved_fixed else auto_pick_source(target, self.source_columns)
            var = tk.StringVar(value=choice)
            combo = ttk.Combobox(
                self.scroll_frame,
                textvariable=var,
                values=values,
                state="readonly",
                width=32,
            )
            combo.grid(row=row_idx, column=1, sticky="ew", pady=2)
            self.combo_vars.append(var)
            fixed_var = tk.StringVar(value=saved_fixed.get(target.index, ""))
            entry = ttk.Entry(self.scroll_frame, textvariable=fixed_var, width=18)
            entry.grid(row=row_idx, column=2, sticky="ew", padx=(8, 0), pady=2)
            entry.configure(state="normal" if choice == FIXED_VALUE else "disabled")
            combo.bind(
                "<<ComboboxSelected>>",
                lambda _event, var=var, entry=entry: entry.configure(
                    state="normal" if var.get() == FIXED_VALUE else "disabled"
                ),
            )
            self.fixed_vars.append(fixed_var)

        self.scroll_frame.columnconfigure(1, weight=1)

    def confirm(self) -> None:
        try:
            header_row = int(self.header_var.get())
            if header_row < 1:
                raise ValueError
        except ValueError:
            messagebox.showwarning("提示", "表头行必须是大于 0 的整数。", parent=self.root)
            return

        pairs: list[tuple[int, int]] = []
        fixed_values: dict[int, str] = {}
        for target, var, fixed_var in zip(self.target_columns, self.combo_vars, self.fixed_vars):
            choice = var.get()
            if choice == SKIP_VALUE:
                continue
            if choice == FIXED_VALUE:
                value = fixed_var.get()
                if not value.strip():
                    messagebox.showwarning("提示", f"请输入中介 {target.letter} 列的固定值。", parent=self.root)
                    return
                fixed_values[target.index] = value
                continue
            src = find_column_by_display(self.source_columns, choice)
            if src is None:
                messagebox.showwarning("提示", f"列选择无效：{choice}", parent=self.root)
                return
            pairs.append((src.index, target.index))

        if not pairs and not fixed_values:
            messagebox.showwarning(
                "提示",
                "请至少映射一列。\n\n若下拉只有「（不填）」或列名不对，请调整「源文件表头行」，再点「刷新列名」。",
                parent=self.root,
            )
            return

        self.result = FileMapping(
            source_path=self.source_path,
            header_row=header_row,
            pairs=pairs,
            sheet=self.sheet_var.get(),
            fixed_values=fixed_values,
        )
        self.root.destroy()

    def skip(self) -> None:
        self.skipped = True
        self.root.destroy()

    def cancel_all(self) -> None:
        self.cancelled = True
        self.root.destroy()

    def run(self) -> FileMapping | None:
        self.root.mainloop()
        if self.cancelled:
            return None
        if self.skipped:
            return None
        return self.result


def int_to_year_month(value: int) -> tuple[int, int] | None:
    if 190001 <= value <= 209912:
        year, month = divmod(value, 100)
        if 1900 <= year <= 2099 and 1 <= month <= 12:
            return year, month
    return None


def excel_serial_to_year_month(value: float | int) -> tuple[int, int] | None:
    serial = int(value)
    if EXCEL_SERIAL_MIN <= serial <= EXCEL_SERIAL_MAX:
        dt = EXCEL_DATE_ORIGIN + timedelta(days=serial)
        return dt.year, dt.month
    return None


def month_end_date(year: int, month: int) -> date:
    return date(year, month, calendar.monthrange(year, month)[1])


def i_value_from_year_month(year: int, month: int) -> tuple[date, str]:
    return month_end_date(year, month), I_DATE_FORMAT


def normalize_column_a(value: object | None) -> tuple[object | None, str]:
    if value is None:
        return None, TEXT_FORMAT
    if isinstance(value, bool):
        return value, TEXT_FORMAT
    if isinstance(value, int):
        return str(value), TEXT_FORMAT
    if isinstance(value, float):
        if value == int(value) and abs(value) < 1e15:
            return str(int(value)), TEXT_FORMAT
        text = format(value, "f").rstrip("0").rstrip(".")
        return text, TEXT_FORMAT
    text = str(value).strip()
    return text or None, TEXT_FORMAT


def normalize_column_i(value: object | None) -> tuple[object | None, str]:
    if value is None:
        return None, I_DATE_FORMAT
    if isinstance(value, datetime):
        return i_value_from_year_month(value.year, value.month)
    if isinstance(value, date):
        return i_value_from_year_month(value.year, value.month)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        number = float(value)
        if number == int(number):
            as_int = int(number)
            ym = int_to_year_month(as_int)
            if ym:
                return i_value_from_year_month(*ym)
            ym = excel_serial_to_year_month(as_int)
            if ym:
                return i_value_from_year_month(*ym)
    if isinstance(value, str):
        text = value.strip()
        if len(text) == 6 and text.isdigit():
            ym = int_to_year_month(int(text))
            if ym:
                return i_value_from_year_month(*ym)
        digits = re.sub(r"\D", "", text)
        if len(digits) >= 6:
            ym = int_to_year_month(int(digits[:6]))
            if ym:
                return i_value_from_year_month(*ym)
        for fmt, size in (("%Y/%m/%d", 10), ("%Y-%m-%d", 10), ("%Y%m%d", 8)):
            try:
                parsed = datetime.strptime(text[:size], fmt).date()
                return i_value_from_year_month(parsed.year, parsed.month)
            except ValueError:
                continue
        return text or None, TEXT_FORMAT
    return str(value).strip() or None, TEXT_FORMAT


def normalize_value_for_target(tgt_col: int, value: object | None) -> tuple[object | None, str | None]:
    if tgt_col == COL_A:
        normalized, fmt = normalize_column_a(value)
        return normalized, fmt
    if tgt_col == COL_I:
        normalized, fmt = normalize_column_i(value)
        return normalized, fmt
    return value, None


def write_target_cell(
    target_ws, row_idx: int, tgt_col: int, value: object | None, *, literal: bool = False,
) -> None:
    normalized, fmt = normalize_value_for_target(tgt_col, value)
    cell = target_ws.cell(row_idx, tgt_col, normalized)
    if literal and isinstance(normalized, str):
        cell.data_type = "s"
    if fmt is not None:
        cell.number_format = fmt


def copy_mapped_rows(
    source_path: Path,
    mapping: FileMapping,
    target_ws,
    next_row: int,
) -> int:
    row_idx = next_row
    for row in iter_source_data_rows(source_path, mapping):
        for src_col, tgt_col in mapping.pairs:
            value = row[src_col - 1] if src_col - 1 < len(row) else None
            write_target_cell(target_ws, row_idx, tgt_col, value)
        for tgt_col, value in mapping.fixed_values.items():
            write_target_cell(target_ws, row_idx, tgt_col, value, literal=True)
        row_idx += 1
    return row_idx


def configure_mappings(
    source_files: list[Path], target_columns: list[ColumnInfo]
) -> tuple[list[FileMapping], bool]:
    """对一批源文件逐个弹列映射框，返回 (mappings, cancelled)。

    cancelled=True 表示用户在某个映射框点「取消全部」，调用方应整体退出全部流程；
    单个源文件被「跳过」只跳过该文件，不影响其余。
    """
    mappings: list[FileMapping] = []
    for src in source_files:
        safe_print(f"  配置映射: {src.name}")
        try:
            dialog = MappingDialog(
                src, target_columns, default_header_row=DEFAULT_SOURCE_HEADER_ROW
            )
        except Exception as exc:
            safe_print(f"    读取源文件失败，已跳过: {exc}")
            continue
        mapping = dialog.run()
        if dialog.cancelled:
            return mappings, True
        if mapping is None:
            safe_print(f"    已跳过: {src.name}")
            continue
        mappings.append(mapping)
        mapped_cols = len(mapping.pairs) + len(mapping.fixed_values)
        preview_rows = count_source_data_rows(mapping.source_path, mapping)
        safe_print(
            f"    已映射 {mapped_cols} 列，源文件可读数据约 {preview_rows} 行"
            f"（页签「{mapping.sheet}」，表头第 {mapping.header_row} 行）"
        )
    return mappings, False


def write_sheet_rows(
    work_wb,
    sheet_name: str,
    target_columns: list[ColumnInfo],
    mappings: list[FileMapping],
    header_row: int,
    ask_write_position=None,
) -> dict | None:
    """把 mappings 写入 work_wb 的指定页签，返回写入摘要 dict。

    ask_write_position 为可注入的「写入位置」询问函数（默认 pick_write_start_row）；
    用户在该框点取消时返回 None——语义为退出全部页签（与单页签行为一致）。
    本函数只改内存中的 work_wb，不负责保存；由 run() 在全部页签写完后统一 save。
    """
    ws = work_wb[sheet_name]
    target_col_indices = [col.index for col in target_columns]
    data_start_row = header_row + 1
    last_row = find_last_data_row_in_cols(ws, data_start_row, target_col_indices)
    ask = ask_write_position or pick_write_start_row
    next_row = ask(header_row, last_row)
    if next_row is None:
        return None
    first_written_row = next_row
    written_mappings = []
    for mapping in mappings:
        source_path = mapping.source_path
        source_sheets = list_source_sheet_names(source_path)
        source_index = mapping.sheet if isinstance(mapping.sheet, int) else source_sheets.index(mapping.sheet)
        while mapping is not None:
            before = next_row
            next_row = copy_mapped_rows(source_path, mapping, ws, next_row)
            written_mappings.append(mapping)
            safe_print(
                f"  写入 {source_path.name}「{source_sheets[source_index]}」："
                f"{next_row - before} 行 → 中介「{sheet_name}」"
            )
            source_index += 1
            if source_index >= len(source_sheets):
                break
            next_sheet = source_sheets[source_index]
            if not ask_continue_source_sheet(source_path, next_sheet, sheet_name):
                break
            dialog = MappingDialog(source_path, target_columns, source_sheet=next_sheet)
            mapping = dialog.run()
            if dialog.cancelled:
                return None
    return {
        "sheet_name": sheet_name,
        "written": next_row - first_written_row,
        "first_row": first_written_row,
        "last_row": last_row,
        "end_row": next_row - 1,
        "target_columns": target_columns,
        "mappings": written_mappings,
    }


def run() -> None:
    t0 = time.time()
    safe_print("=" * 60)
    safe_print("列映射填表（10）")
    safe_print("=" * 60)

    picked = resolve_intermediate()
    if picked is None:
        safe_print("已取消。")
        sys.exit(0)
    intermediate, header_row = picked

    safe_print(f"  中介文件: {intermediate}")
    safe_print(f"  表头行: {header_row}")

    # 运行时多选要处理的中介文件页签。
    sheet_names = list_intermediate_sheet_names(intermediate)
    if not sheet_names:
        safe_print("错误：中介文件中没有页签。")
        sys.exit(1)
    selected_names = pick_intermediate_sheets(sheet_names)
    if not selected_names:
        safe_print("已取消。")
        sys.exit(0)
    safe_print(f"  已选页签（{len(selected_names)} 个）: {'、'.join(selected_names)}")

    source_files = pick_excel_files("选择源 Excel 文件（可多选）")
    if not source_files:
        safe_print("已取消。")
        sys.exit(0)

    source_files = [f for f in source_files if f.resolve() != intermediate.resolve()]
    if not source_files:
        safe_print("错误：请至少选择一个与中介文件不同的源文件。")
        sys.exit(1)

    safe_print(f"  已选 {len(source_files)} 个源文件")

    # 第一个选中页签：先读表头并配置映射（与旧单页签路径一致，在建立工作文件之前）。
    first_target_columns = read_columns(intermediate, header_row, sheet=selected_names[0])
    if not first_target_columns:
        safe_print("错误：中介文件未找到有效列。")
        sys.exit(1)
    safe_print(f"  配置第一个页签「{selected_names[0]}」的列映射…")
    first_mappings, cancelled = configure_mappings(source_files, first_target_columns)
    if cancelled:
        safe_print("已取消。")
        sys.exit(0)
    if not first_mappings:
        safe_print("没有可处理的源文件映射，已退出。")
        sys.exit(0)

    out_path = pick_save_path(intermediate)
    if out_path is None:
        safe_print("已取消保存。")
        sys.exit(0)

    if out_path.resolve() == intermediate.resolve():
        backup = backup_file(intermediate)
        safe_print(f"  已备份: {backup.name}")

    work_path = out_path.with_name(f"{out_path.stem}.列映射.work{out_path.suffix}")
    if work_path.exists():
        work_path.unlink()
    shutil.copy2(intermediate, work_path)

    processed: list[dict] = []
    grand_total = 0

    wb = load_workbook(work_path)
    try:
        for idx, sheet_name in enumerate(selected_names):
            if idx == 0:
                target_columns = first_target_columns
                mappings = first_mappings
            else:
                target_columns = read_columns(intermediate, header_row, sheet=sheet_name)
                if not target_columns:
                    safe_print(f"  页签「{sheet_name}」未找到有效列，跳过。")
                    if idx < len(selected_names) - 1 and not ask_continue_next_sheet(selected_names[idx + 1]):
                        break
                    continue

                if sheet_headers_identical(first_target_columns, target_columns):
                    if ask_reuse_mapping(sheet_name):
                        mappings = first_mappings
                        safe_print(f"  页签「{sheet_name}」复用第一页签的列映射。")
                    else:
                        safe_print(f"  页签「{sheet_name}」单独配置列映射…")
                        mappings, cancelled = configure_mappings(source_files, target_columns)
                        if cancelled:
                            wb.close()
                            work_path.unlink(missing_ok=True)
                            safe_print("已取消。")
                            sys.exit(0)
                else:
                    safe_print(f"  页签「{sheet_name}」表头与第一个不同，单独配置列映射…")
                    mappings, cancelled = configure_mappings(source_files, target_columns)
                    if cancelled:
                        wb.close()
                        work_path.unlink(missing_ok=True)
                        safe_print("已取消。")
                        sys.exit(0)

                if not mappings:
                    safe_print(f"  页签「{sheet_name}」没有可用映射，跳过写入。")
                    if idx < len(selected_names) - 1 and not ask_continue_next_sheet(selected_names[idx + 1]):
                        break
                    continue

            result = write_sheet_rows(wb, sheet_name, target_columns, mappings, header_row)
            if result is None:  # 写入位置框点取消 = 退出全部页签
                wb.close()
                work_path.unlink(missing_ok=True)
                safe_print("已取消。")
                sys.exit(0)
            processed.append(result)
            grand_total += result["written"]
            safe_print(
                f"  页签「{sheet_name}」写入 {result['written']} 行"
                f"（第 {result['first_row']}～{result['end_row']} 行；表头第 {header_row} 行，已有数据到第 {result['last_row']} 行）"
            )

            # 后面还有选中页签：写完一个就问是否继续下一个。
            if idx < len(selected_names) - 1:
                next_name = selected_names[idx + 1]
                if not ask_continue_next_sheet(next_name):
                    safe_print("  已按选择停止，未处理后续页签。")
                    break

        if grand_total == 0:
            wb.close()
            work_path.unlink(missing_ok=True)
            details = []
            for st in processed:
                for mapping in st["mappings"]:
                    preview = count_source_data_rows(mapping.source_path, mapping)
                    details.append(
                        f"页签「{st['sheet_name']}」\n"
                        f"{mapping.source_path.name}\n"
                        f"  表头行: {mapping.header_row}，数据从第 {mapping.header_row + 1} 行起\n"
                        f"  扫描到可写入行: {preview}\n"
                        f"{describe_mapping(mapping, st['target_columns'])}"
                    )
            detail_text = "\n\n".join(details) if details else "（无）"
            safe_print("-" * 60)
            safe_print("未写入任何行。诊断：")
            safe_print(detail_text)
            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            messagebox.showwarning(
                "未写入数据",
                "没有从源文件读到可写入的数据行。\n\n"
                "请检查：\n"
                "  1. 源文件表头行（默认第 1 行）是否选对\n"
                "  2. 列映射是否选对（不要只留一列，其余都「不填」）\n"
                "  3. 源文件对应列第 2 行起是否有数据\n"
                "  4. 若单元格是公式，请先在 Excel 里打开并保存一次（写入缓存值）\n\n"
                f"{detail_text}",
                parent=root,
            )
            root.destroy()
            sys.exit(1)

        # 全部选中（或已确认继续的）页签写完，统一保存一次。
        wb.save(work_path)
    finally:
        wb.close()

    if work_path.resolve() != out_path.resolve():
        if out_path.exists():
            out_path.unlink()
        shutil.move(str(work_path), str(out_path))
    else:
        work_path.replace(out_path)

    work_path = out_path

    elapsed = time.time() - t0
    safe_print("-" * 60)
    safe_print(f"完成。共写入 {grand_total} 行，涉及 {len(processed)} 个页签。")
    safe_print(f"输出: {work_path}")
    for st in processed:
        safe_print(f"  页签「{st['sheet_name']}」: 第 {st['first_row']} ～ {st['end_row']} 行（{st['written']} 行）")
    safe_print(f"耗时: {elapsed:.1f} 秒")

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    messagebox.showinfo(
        "写入完成",
        f"共写入 {grand_total} 行，涉及 {len(processed)} 个页签。\n\n"
        + "\n".join(
            f"  · 「{st['sheet_name']}」第 {st['first_row']}～{st['end_row']} 行"
            for st in processed
        ),
        parent=root,
    )
    root.destroy()


if __name__ == "__main__":
    try:
        run()
    except Exception as exc:
        safe_print("-" * 60)
        safe_print(f"运行出错: {exc}")
        traceback.print_exc()
        sys.exit(1)
