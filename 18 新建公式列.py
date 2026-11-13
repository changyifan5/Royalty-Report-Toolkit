"""
在指定页签、指定列后面插入一列或多列。可写 Excel 公式，也可只插空白列。

流程：
  1. 弹窗选 Excel
  2. 选择页签
  3. 选择表头行（默认第 1 行）
  4. 选择插在哪一列后面；添加要新建的列（每列：列名、公式、颜色）
  5. 备份为 .bak 后写回
  6. 完成后可：留在本页签继续 / 换页签继续 / 结束

公式：
  · 按第 2 行来写即可，例如 =G2/H2（脚本会自动套到每一行）
  · 公式留空 = 空白列（只写表头）
  · 不要用 $G$2（绝对引用，行号不会变）

Windows 使用 Excel COM；Mac 使用 openpyxl。运行前请关闭 Excel。
运行：python "18 新建公式列.py"
"""

from __future__ import annotations

import json
import shutil
import sys
import time
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import colorchooser, messagebox, ttk

from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter

from file_picker_common import require_excel_file
from row_order_common import find_last_data_row

BASE_DIR = Path(__file__).resolve().parent
PS1_SCRIPT = BASE_DIR / "18_insert_formula_col_com.ps1"
MANIFEST_FILE = BASE_DIR / "18_insert_formula_col_manifest.json"
LOG_FILE = BASE_DIR / "18_insert_formula_col_com.log"

DEFAULT_HEADER_ROW = 1
HEADER_SCAN_LIMIT = 500
HEADER_PREVIEW_ROWS = 20
HEADER_PREVIEW_COLS = 8
APPEND_VALUE = "（最右侧追加）"
DEFAULT_FILL_HEX = "FFF2CC"
FILL_PRESETS: tuple[tuple[str, str], ...] = (
    ("浅黄", "FFF2CC"),
    ("橙", "F8CBAD"),
    ("绿", "C6EFCE"),
    ("蓝", "BDD7EE"),
    ("无", ""),
)

FORMULA_HINT = (
    "公式按第 2 行写即可，例如 =G2/H2；留空 = 空白列。每列可单独选颜色。"
)


@dataclass
class ColumnInfo:
    index: int
    letter: str
    header: str
    display: str


@dataclass
class ColumnPlan:
    new_header: str
    formula: str = ""
    fill_hex: str = DEFAULT_FILL_HEX


@dataclass
class InsertSpec:
    after_col: int
    after_letter: str
    header_row: int
    columns: list[ColumnPlan]
    append: bool = False


@dataclass
class SheetJob:
    sheet_name: str
    after_col: int
    after_letter: str
    new_header: str
    formula: str
    data_start: int
    data_end: int
    formula_first: str
    fill_hex: str


def safe_print(text: str) -> None:
    try:
        print(text, flush=True)
    except UnicodeEncodeError:
        print(
            text.encode(sys.stdout.encoding or "utf-8", errors="replace").decode(
                sys.stdout.encoding or "utf-8", errors="replace"
            ),
            flush=True,
        )


def list_sheet_names(path: Path) -> list[str]:
    wb = load_workbook(path, read_only=True, keep_links=False)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def _columns_from_header_cells(header_cells) -> list[ColumnInfo]:
    columns: list[ColumnInfo] = []
    for col_idx, raw in enumerate(header_cells, start=1):
        text = "" if raw is None else str(raw).strip()
        letter = get_column_letter(col_idx)
        display = f"{letter} · {text}" if text else letter
        columns.append(ColumnInfo(index=col_idx, letter=letter, header=text, display=display))
    while columns and not columns[-1].header:
        columns.pop()
    return columns


def read_columns(
    path: Path,
    sheet_name: str,
    header_row: int,
    *,
    force_full: bool = False,
) -> list[ColumnInfo]:
    """读当前表头。不用 data_only，才能看到刚写入、尚未算出缓存的新列。"""
    wb = load_workbook(path, read_only=not force_full, data_only=False, keep_links=False)
    try:
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"找不到页签「{sheet_name}」")
        ws = wb[sheet_name]
        header_cells = next(
            ws.iter_rows(
                min_row=header_row,
                max_row=header_row,
                max_col=HEADER_SCAN_LIMIT,
                values_only=True,
            ),
            (),
        )
        return _columns_from_header_cells(header_cells)
    finally:
        wb.close()


def sheet_data_end(path: Path, sheet_name: str, data_start: int) -> int:
    """数据最后一行：先看 A 列，A 为空则扫前 20 列。"""
    wb = load_workbook(path, read_only=True, keep_links=False)
    try:
        if sheet_name not in wb.sheetnames:
            return data_start - 1
        ws = wb[sheet_name]
        last = find_last_data_row(ws, data_start_row=data_start)
        if last >= data_start:
            return last
        last = data_start - 1
        for idx, row in enumerate(
            ws.iter_rows(min_row=data_start, max_col=20, values_only=True)
        ):
            if any(v is not None and str(v).strip() for v in row):
                last = data_start + idx
        return last
    finally:
        wb.close()


def pick_sheet(path: Path, *, force_dialog: bool = False, current: str | None = None) -> str | None:
    names = list_sheet_names(path)
    if not names:
        safe_print("错误：工作簿中没有页签。")
        sys.exit(1)
    if len(names) == 1 and not force_dialog:
        return names[0]

    result: dict[str, str | None] = {"name": None}
    root = tk.Tk()
    root.title("选择页签")
    root.attributes("-topmost", True)
    root.resizable(False, False)
    frame = ttk.Frame(root, padding=12)
    frame.pack(fill="both", expand=True)
    ttk.Label(frame, text=f"{path.name}：选择要操作的页签").pack(anchor="w", pady=(0, 8))
    listbox = tk.Listbox(frame, width=48, height=min(14, max(4, len(names))), font=("Consolas", 10))
    listbox.pack()
    for name in names:
        listbox.insert(tk.END, name)
    listbox.selection_set(0)
    listbox.see(0)
    if current and current in names:
        idx = names.index(current)
        listbox.selection_clear(0, tk.END)
        listbox.selection_set(idx)
        listbox.see(idx)

    def confirm() -> None:
        sel = listbox.curselection()
        if not sel:
            messagebox.showwarning("提示", "请选择一个页签。", parent=root)
            return
        result["name"] = names[sel[0]]
        root.destroy()

    def cancel() -> None:
        root.destroy()

    btns = ttk.Frame(frame)
    btns.pack(fill="x", pady=(12, 0))
    ttk.Button(btns, text="确定", command=confirm, width=10).pack(side="right")
    ttk.Button(btns, text="取消", command=cancel, width=10).pack(side="right", padx=(0, 6))
    listbox.bind("<Double-Button-1>", lambda _e: confirm())
    root.bind("<Return>", lambda _e: confirm())
    root.bind("<Escape>", lambda _e: cancel())
    root.mainloop()
    return result["name"]


def preview_sheet_rows(path: Path, sheet_name: str, max_rows: int = HEADER_PREVIEW_ROWS) -> list[tuple[int, str]]:
    wb = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        if sheet_name not in wb.sheetnames:
            return [(DEFAULT_HEADER_ROW, "（无法读取）")]
        ws = wb[sheet_name]
        items: list[tuple[int, str]] = []
        for row_idx, row in enumerate(
            ws.iter_rows(min_row=1, max_row=max_rows, max_col=HEADER_PREVIEW_COLS, values_only=True),
            start=1,
        ):
            parts: list[str] = []
            for raw in row:
                text = "" if raw is None else str(raw).replace("\n", " ").strip()
                if text:
                    parts.append(text[:18])
            preview = " | ".join(parts) if parts else "（空）"
            items.append((row_idx, preview))
        if not items:
            items.append((DEFAULT_HEADER_ROW, "（空）"))
        return items
    finally:
        wb.close()


def pick_header_row(path: Path, sheet_name: str) -> int | None:
    rows = preview_sheet_rows(path, sheet_name)
    result: dict[str, int | None] = {"row": None}
    root = tk.Tk()
    root.title("选择表头行")
    root.attributes("-topmost", True)
    root.resizable(False, False)
    frame = ttk.Frame(root, padding=12)
    frame.pack(fill="both", expand=True)
    ttk.Label(
        frame,
        text=f"{path.name} / {sheet_name}\n请选择哪一行是表头（默认第 {DEFAULT_HEADER_ROW} 行，数据从下一行起）",
        justify="left",
    ).pack(anchor="w", pady=(0, 8))
    listbox = tk.Listbox(frame, width=72, height=min(16, max(6, len(rows))), font=("Consolas", 10))
    listbox.pack()
    default_index = 0
    for idx, (row_idx, preview) in enumerate(rows):
        mark = "（默认）" if row_idx == DEFAULT_HEADER_ROW else "        "
        listbox.insert(tk.END, f"第 {row_idx} 行 {mark} {preview}")
        if row_idx == DEFAULT_HEADER_ROW:
            default_index = idx
    listbox.selection_set(default_index)
    listbox.see(default_index)

    def confirm() -> None:
        sel = listbox.curselection()
        if not sel:
            messagebox.showwarning("提示", "请选择表头所在行。", parent=root)
            return
        result["row"] = rows[sel[0]][0]
        root.destroy()

    def cancel() -> None:
        root.destroy()

    btns = ttk.Frame(frame)
    btns.pack(fill="x", pady=(12, 0))
    ttk.Button(btns, text="确定", command=confirm, width=10).pack(side="right")
    ttk.Button(btns, text="取消", command=cancel, width=10).pack(side="right", padx=(0, 6))
    listbox.bind("<Double-Button-1>", lambda _e: confirm())
    root.bind("<Return>", lambda _e: confirm())
    root.bind("<Escape>", lambda _e: cancel())
    root.mainloop()
    return result["row"]


def normalize_formula(raw: str) -> str:
    text = raw.strip()
    if not text:
        return ""
    if not text.startswith("="):
        text = "=" + text
    return text


def formula_for_data_start(template: str, data_start: int) -> str:
    """用户按第 2 行写公式，平移到实际数据起始行。"""
    text = normalize_formula(template)
    if not text:
        return ""
    if data_start == 2:
        return text
    try:
        return Translator(text, origin="A2").translate_formula(f"A{data_start}")
    except Exception:
        return text


def preview_row2_formula(template: str) -> tuple[str, str]:
    """预览：第 2 行写法 → 第 3 行自动结果。"""
    first = normalize_formula(template)
    if not first:
        return "", ""
    try:
        nxt = Translator(first, origin="A2").translate_formula("A3")
    except Exception:
        nxt = first
    return first, nxt


def hex_to_bgr(hex_color: str) -> int:
    text = hex_color.strip().lstrip("#")
    if len(text) != 6:
        raise ValueError(hex_color)
    red = int(text[0:2], 16)
    green = int(text[2:4], 16)
    blue = int(text[4:6], 16)
    return red + (green << 8) + (blue << 16)


def normalize_fill_hex(raw: str) -> str:
    text = raw.strip().lstrip("#").upper()
    if not text:
        return ""
    if len(text) != 6:
        return DEFAULT_FILL_HEX
    int(text, 16)
    return text


def backup_file(path: Path) -> Path:
    backup_path = path.with_suffix(path.suffix + ".bak")
    if backup_path.exists():
        backup_path.unlink()
    shutil.copy2(path, backup_path)
    return backup_path


def ask_next_step(sheet_name: str, summary: str) -> str:
    """完成后询问下一步。返回 same / other / done。"""
    result = {"action": "done"}
    root = tk.Tk()
    root.title("是否下一步")
    root.attributes("-topmost", True)
    root.resizable(False, False)
    frame = ttk.Frame(root, padding=16)
    frame.pack(fill="both", expand=True)
    ttk.Label(frame, text=summary, justify="left", wraplength=560).pack(anchor="w")
    ttk.Label(
        frame,
        text=f"当前页签：{sheet_name}\n是否继续下一步？",
        justify="left",
    ).pack(anchor="w", pady=(10, 0))

    def choose(action: str) -> None:
        result["action"] = action
        root.destroy()

    btns = ttk.Frame(frame)
    btns.pack(fill="x", pady=(12, 0))
    ttk.Button(btns, text="是，留在本页签", command=lambda: choose("same"), width=16).pack(
        side="left"
    )
    ttk.Button(btns, text="是，换页签", command=lambda: choose("other"), width=14).pack(
        side="left", padx=(8, 0)
    )
    ttk.Button(btns, text="否", command=lambda: choose("done"), width=8).pack(side="right")
    root.bind("<Escape>", lambda _e: choose("done"))
    root.mainloop()
    return result["action"]



class OperationDialog:
    def __init__(
        self,
        path: Path,
        sheet_name: str,
        columns: list[ColumnInfo],
        header_row: int,
        data_start: int,
        data_end: int,
        prefer_after_letter: str | None = None,
    ) -> None:
        self.path = path
        self.sheet_name = sheet_name
        self.columns = columns
        self.data_start = data_start
        self.data_end = data_end
        self.prefer_after_letter = (prefer_after_letter or "").strip().upper()
        self.result: InsertSpec | None = None
        self.column_rows: list[dict] = []

        self.root = tk.Tk()
        self.root.title("新建公式列")
        self.root.attributes("-topmost", True)
        self.root.geometry("860x640")
        self.root.minsize(700, 520)

        outer = ttk.Frame(self.root, padding=12)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(4, weight=1)

        ttk.Label(
            outer,
            text=f"{path.name}  /  {sheet_name}",
            font=("Microsoft YaHei UI", 10, "bold"),
        ).grid(row=0, column=0, sticky="w")
        self.columns_preview_var = tk.StringVar()
        ttk.Label(
            outer,
            textvariable=self.columns_preview_var,
            foreground="#1a5276",
            wraplength=780,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        top = ttk.Frame(outer)
        top.grid(row=2, column=0, sticky="ew", pady=(8, 6))
        ttk.Label(top, text="表头行", font=("Microsoft YaHei UI", 10, "bold")).pack(side="left")
        self.header_var = tk.StringVar(value=str(header_row))
        header_spin = ttk.Spinbox(
            top, from_=1, to=999, width=6, textvariable=self.header_var, command=self._on_header_changed
        )
        header_spin.pack(side="left", padx=(6, 12))
        header_spin.bind("<FocusOut>", lambda _e: self._on_header_changed())
        header_spin.bind("<Return>", lambda _e: self._on_header_changed())
        self.info_var = tk.StringVar()
        ttk.Label(top, textvariable=self.info_var, foreground="#1a5276").pack(side="left")

        after_row = ttk.Frame(outer)
        after_row.grid(row=3, column=0, sticky="ew", pady=(0, 6))
        ttk.Label(after_row, text="插在哪一列后面").pack(side="left")
        values = [col.display for col in columns] + [APPEND_VALUE]
        self.after_var = tk.StringVar(value=self._default_after_value(values))
        self.after_combo = ttk.Combobox(
            after_row, textvariable=self.after_var, values=values, state="readonly", width=40
        )
        self.after_combo.pack(side="left", padx=(8, 0))

        list_wrap = ttk.LabelFrame(outer, text="要新建的列（可多列；公式按第2行写，可留空）", padding=8)
        list_wrap.grid(row=4, column=0, sticky="nsew", pady=(4, 0))
        list_wrap.columnconfigure(0, weight=1)
        list_wrap.rowconfigure(0, weight=1)

        canvas = tk.Canvas(list_wrap, highlightthickness=0)
        scroll = ttk.Scrollbar(list_wrap, orient="vertical", command=canvas.yview)
        self.rows_frame = ttk.Frame(canvas)
        self.rows_frame.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.rows_frame, anchor="nw")
        canvas.configure(yscrollcommand=scroll.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")

        head = ttk.Frame(self.rows_frame)
        head.pack(fill="x", pady=(0, 4))
        ttk.Label(head, text="列名", width=14).pack(side="left")
        ttk.Label(head, text="公式（第2行写法，可留空）", width=36).pack(side="left", padx=(6, 0))
        ttk.Label(head, text="颜色").pack(side="left", padx=(6, 0))

        ttk.Label(outer, text=FORMULA_HINT, foreground="#666").grid(row=5, column=0, sticky="w", pady=(6, 0))

        self.preview_var = tk.StringVar()
        ttk.Label(
            outer, textvariable=self.preview_var, foreground="#1a5276", justify="left", wraplength=780
        ).grid(row=6, column=0, sticky="w", pady=(6, 0))

        btns = ttk.Frame(outer)
        btns.grid(row=7, column=0, sticky="ew", pady=(12, 0))
        ttk.Button(btns, text="添加一列", command=self._add_row, width=10).pack(side="left")
        ttk.Button(btns, text="取消", command=self.cancel, width=10).pack(side="right")
        ttk.Button(btns, text="确定并写入", command=self.confirm, width=12).pack(side="right", padx=(0, 6))

        self.root.bind("<Escape>", lambda _e: self.cancel())
        self.after_var.trace_add("write", lambda *_a: self._update_formula_preview())
        self.header_var.trace_add("write", lambda *_a: self._on_header_changed())
        self._refresh_columns_preview()
        self._refresh_example_labels()
        self._add_row()
        self._update_formula_preview()

    def _default_after_value(self, values: list[str]) -> str:
        if not values:
            return APPEND_VALUE
        if self.prefer_after_letter:
            for col in self.columns:
                if col.letter.upper() == self.prefer_after_letter:
                    return col.display
        return values[0]

    def _refresh_columns_preview(self) -> None:
        self.columns_preview_var.set("当前表头：" + format_headers_line(self.columns))

    def _current_header_row(self) -> int:
        try:
            row = int(self.header_var.get())
            return row if row >= 1 else DEFAULT_HEADER_ROW
        except (TypeError, ValueError):
            return DEFAULT_HEADER_ROW

    def _refresh_example_labels(self) -> None:
        example = self.data_start
        last = self.data_end if self.data_end >= example else example
        self.info_var.set(f"数据第 {example}～{last} 行 · 公式请按第 2 行写")

    def _reload_after_combo(self) -> None:
        values = [col.display for col in self.columns] + [APPEND_VALUE]
        current = self.after_var.get()
        self.after_combo.configure(values=values)
        if current in values:
            return
        self.after_var.set(self._default_after_value(values))

    def _on_header_changed(self) -> None:
        if not hasattr(self, "rows_frame"):
            return
        header_row = self._current_header_row()
        self.data_start = header_row + 1
        try:
            self.columns = read_columns(self.path, self.sheet_name, header_row)
        except Exception:
            return
        self._refresh_columns_preview()
        self._reload_after_combo()
        self.data_end = sheet_data_end(self.path, self.sheet_name, self.data_start)
        self._refresh_example_labels()
        self._update_formula_preview()

    def _add_row(self, name: str = "", formula: str = "", fill_hex: str = DEFAULT_FILL_HEX) -> None:
        row = {
            "frame": ttk.Frame(self.rows_frame),
            "name": tk.StringVar(value=name),
            "formula": tk.StringVar(value=formula),
            "fill": tk.StringVar(value=fill_hex),
        }
        row["frame"].pack(fill="x", pady=3)
        ttk.Entry(row["frame"], textvariable=row["name"], width=16).pack(side="left")
        ttk.Entry(row["frame"], textvariable=row["formula"], width=38).pack(side="left", padx=(6, 0))
        swatch = tk.Canvas(row["frame"], width=22, height=16, highlightthickness=1, highlightbackground="#888")
        swatch.pack(side="left", padx=(8, 4))
        row["swatch"] = swatch

        def paint() -> None:
            color = row["fill"].get().strip() or "FFFFFF"
            try:
                swatch.configure(bg=f"#{color}")
            except tk.TclError:
                swatch.configure(bg="#FFF2CC")

        def set_fill(hex_color: str) -> None:
            row["fill"].set(hex_color)
            paint()
            self._update_formula_preview()

        for label, hex_color in FILL_PRESETS:
            tk.Button(
                row["frame"],
                text=label,
                width=3,
                relief="groove",
                bg=f"#{hex_color}" if hex_color else "#FFFFFF",
                command=lambda h=hex_color, sf=set_fill: sf(h),
            ).pack(side="left", padx=1)

        def pick() -> None:
            current = row["fill"].get().strip() or DEFAULT_FILL_HEX
            picked = colorchooser.askcolor(color=f"#{current}", parent=self.root, title="选择表头填充色")
            if picked and picked[1]:
                set_fill(str(picked[1]).lstrip("#").upper())

        ttk.Button(row["frame"], text="自选", width=4, command=pick).pack(side="left", padx=(2, 0))

        def remove() -> None:
            if len(self.column_rows) <= 1:
                messagebox.showinfo("提示", "至少保留一列。", parent=self.root)
                return
            row["frame"].destroy()
            self.column_rows.remove(row)
            self._update_formula_preview()

        ttk.Button(row["frame"], text="删", width=3, command=remove).pack(side="left", padx=(6, 0))
        row["name"].trace_add("write", lambda *_a: self._update_formula_preview())
        row["formula"].trace_add("write", lambda *_a: self._update_formula_preview())
        paint()
        self.column_rows.append(row)
        self._update_formula_preview()

    def cancel(self) -> None:
        self.root.destroy()

    def _resolve_after_col(self) -> tuple[int, str] | None:
        choice = self.after_var.get().strip()
        if choice == APPEND_VALUE:
            if not self.columns:
                return 0, "A"
            last = self.columns[-1]
            return last.index, last.letter
        for col in self.columns:
            if col.display == choice:
                return col.index, col.letter
        return None

    def _collect_plans(self) -> list[ColumnPlan] | None:
        plans: list[ColumnPlan] = []
        for idx, row in enumerate(self.column_rows, start=1):
            name = row["name"].get().strip()
            if not name:
                if row["formula"].get().strip():
                    messagebox.showwarning("提示", f"第 {idx} 行已填公式，请填写列名。", parent=self.root)
                    return None
                continue
            plans.append(
                ColumnPlan(
                    new_header=name,
                    formula=normalize_formula(row["formula"].get()),
                    fill_hex=normalize_fill_hex(row["fill"].get()),
                )
            )
        if not plans:
            messagebox.showwarning("提示", "请至少填写一列的列名。", parent=self.root)
            return None
        return plans

    def _update_formula_preview(self) -> None:
        if not hasattr(self, "preview_var"):
            return
        after = self._resolve_after_col()
        if after is None:
            self.preview_var.set("请先选择插在哪一列后面。")
            return
        after_col, _after_letter = after
        lines: list[str] = []
        for idx, row in enumerate(self.column_rows, start=1):
            name = row["name"].get().strip() or f"（第{idx}列未命名）"
            new_letter = get_column_letter(after_col + idx)
            formula = row["formula"].get().strip()
            color = row["fill"].get().strip()
            color_note = f"#{color}" if color else "无色"
            if formula:
                first, nxt = preview_row2_formula(formula)
                lines.append(f"{new_letter} {name} [{color_note}]: {first} -> {nxt}")
            else:
                lines.append(f"{new_letter} {name} [{color_note}]: 空白列")
        self.preview_var.set("\n".join(lines) if lines else "点击「添加一列」开始配置。")

    def confirm(self) -> None:
        after = self._resolve_after_col()
        if after is None:
            messagebox.showwarning("提示", "请选择要插在哪一列后面。", parent=self.root)
            return
        after_col, after_letter = after
        plans = self._collect_plans()
        if plans is None:
            return
        header_row = self._current_header_row()
        self.data_start = header_row + 1
        append = self.after_var.get().strip() == APPEND_VALUE
        preview_lines: list[str] = []
        for idx, plan in enumerate(plans, start=1):
            new_letter = get_column_letter(after_col + idx)
            color_note = f"#{plan.fill_hex}" if plan.fill_hex else "无色"
            if plan.formula:
                first, nxt = preview_row2_formula(plan.formula)
                preview_lines.append(f"{new_letter} 「{plan.new_header}」[{color_note}]: {first} -> {nxt}")
            else:
                preview_lines.append(f"{new_letter} 「{plan.new_header}」[{color_note}]: 空白列")
        msg = (
            f"页签：{self.sheet_name}\n"
            f"将在「{after_letter}」列后插入 {len(plans)} 列。\n"
            + "\n".join(preview_lines)
            + "\n\n继续？"
        )
        ok = messagebox.askokcancel("确认写入", msg, parent=self.root)
        if not ok:
            return
        self.result = InsertSpec(
            after_col=after_col,
            after_letter=after_letter,
            header_row=header_row,
            columns=plans,
            append=append,
        )
        self.root.destroy()

    def run(self) -> InsertSpec | None:
        self.root.mainloop()
        return self.result



def collect_jobs(
    path: Path,
    sheet_names: list[str],
    spec: InsertSpec,
    data_start: int,
) -> list[SheetJob]:
    wb = load_workbook(path, read_only=True, keep_links=False)
    jobs: list[SheetJob] = []
    try:
        for name in sheet_names:
            if name not in wb.sheetnames:
                continue
            ws = wb[name]
            if spec.append:
                header_cells = next(
                    ws.iter_rows(
                        min_row=spec.header_row,
                        max_row=spec.header_row,
                        max_col=HEADER_SCAN_LIMIT,
                        values_only=True,
                    ),
                    (),
                )
                after_col = 0
                for idx, raw in enumerate(header_cells, start=1):
                    if raw is not None and str(raw).strip():
                        after_col = idx
            else:
                after_col = spec.after_col
            data_end = find_last_data_row(ws, data_start_row=data_start)
            current_after = after_col
            for plan in spec.columns:
                after_letter = get_column_letter(max(current_after, 1))
                formula_first = (
                    formula_for_data_start(plan.formula, data_start) if plan.formula else ""
                )
                jobs.append(
                    SheetJob(
                        sheet_name=name,
                        after_col=current_after,
                        after_letter=after_letter,
                        new_header=plan.new_header,
                        formula=plan.formula,
                        data_start=data_start,
                        data_end=data_end,
                        formula_first=formula_first,
                        fill_hex=plan.fill_hex,
                    )
                )
                current_after += 1
    finally:
        wb.close()
    return jobs


def apply_formula_to_sheet(ws, job: SheetJob, header_row: int) -> int:
    new_col = job.after_col + 1
    if new_col < 1:
        new_col = 1
    new_letter = get_column_letter(new_col)
    ws.insert_cols(new_col)
    ws.cell(header_row, new_col, job.new_header)
    written = 0
    if job.formula and job.data_end >= job.data_start:
        first = job.formula_first
        origin = f"{new_letter}{job.data_start}"
        ws.cell(job.data_start, new_col, first)
        written = 1
        for row in range(job.data_start + 1, job.data_end + 1):
            dest = f"{new_letter}{row}"
            try:
                formula = Translator(first, origin=origin).translate_formula(dest)
            except Exception:
                formula = formula_for_data_start(job.formula, row)
            ws.cell(row, new_col, formula)
            written += 1
    if job.fill_hex:
        ws.cell(header_row, new_col).fill = PatternFill("solid", fgColor=job.fill_hex)
    return written


def run_com(path: Path, spec: InsertSpec, jobs: list[SheetJob]) -> tuple[int, int]:
    if not PS1_SCRIPT.exists():
        print(f"错误：找不到 COM 脚本: {PS1_SCRIPT}")
        sys.exit(1)

    from com_runner_common import stream_com

    payload = {
        "path": str(path.resolve()),
        "header_row": spec.header_row,
        "jobs": [
            {
                "sheet_name": job.sheet_name,
                "after_col": job.after_col,
                "new_header": job.new_header,
                "fill_hex": job.fill_hex,
                "fill_bgr": hex_to_bgr(job.fill_hex) if job.fill_hex else None,
                "formula": job.formula,
                "data_start_row": job.data_start,
                "data_end_row": job.data_end,
                "formula_first": job.formula_first,
            }
            for job in jobs
        ],
    }
    MANIFEST_FILE.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    cmd = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(PS1_SCRIPT),
        "-ManifestPath",
        str(MANIFEST_FILE.resolve()),
    ]
    safe_print(f"  Excel COM：插入列/写入公式（{len(jobs)} 列）...")
    returncode, lines = stream_com(
        cmd,
        log_file=LOG_FILE,
        progress_title="新建公式列 · Excel COM",
        stage_labels={"insert": "正在插入新列...", "formula": "正在写入公式..."},
        progress_labels={"sheets": "页签进度", "formula": "写入公式"},
    )
    if returncode != 0:
        print()
        print("错误：COM 执行失败，日志:", LOG_FILE)
        sys.exit(1)
    for line in lines:
        if line.startswith("STATS,"):
            parts = line.split(",")
            if len(parts) >= 4:
                try:
                    return int(parts[2]), int(parts[3])
                except ValueError:
                    break
    written = sum(
        max(0, job.data_end - job.data_start + 1)
        for job in jobs
        if job.formula and job.data_end >= job.data_start
    )
    return written, len(jobs)


def run_openpyxl(path: Path, spec: InsertSpec, jobs: list[SheetJob]) -> tuple[int, int]:
    from progress_ui import open_progress_display

    total = max(1, len(jobs))
    safe_print(f"  openpyxl：插入列/写入公式（{len(jobs)} 列）...")
    bar = open_progress_display("新建公式列 · openpyxl", total=total)
    written = 0
    done_sheets = 0
    wb = load_workbook(path)
    try:
        for idx, job in enumerate(jobs, start=1):
            if job.sheet_name not in wb.sheetnames:
                safe_print(f"  跳过：找不到页签「{job.sheet_name}」")
                continue
            ws = wb[job.sheet_name]
            added = apply_formula_to_sheet(ws, job, spec.header_row)
            written += added
            done_sheets += 1
            status = f"{job.new_header} · {added} 行" if job.formula else f"{job.new_header} · 空白列"
            bar.update(idx, status=status)
            safe_print(f"  {job.sheet_name}: 插入「{job.new_header}」完成")
        bar.update(total, status="保存")
        wb.save(path)
    finally:
        wb.close()
        bar.close()
    return written, done_sheets


def format_headers_line(columns: list[ColumnInfo], limit: int = 40) -> str:
    if not columns:
        return "（当前没有表头）"
    parts = [f"{c.letter} {c.header or '（空）'}" for c in columns[:limit]]
    text = " | ".join(parts)
    if len(columns) > limit:
        text += " | …"
    return text


def load_sheet_context(
    path: Path,
    sheet_name: str,
    header_row: int,
    *,
    expect_header: str | None = None,
) -> tuple[list[ColumnInfo], int, int]:
    data_start = header_row + 1
    data_end = sheet_data_end(path, sheet_name, data_start)
    columns = read_columns(path, sheet_name, header_row)
    expect = (expect_header or "").strip()
    if expect and not any(c.header == expect for c in columns):
        time.sleep(0.4)
        columns = read_columns(path, sheet_name, header_row, force_full=True)
    return columns, data_start, max(data_end, data_start)


def apply_one_step(path: Path, sheet_name: str, spec: InsertSpec) -> tuple[str, int, int]:
    data_start = spec.header_row + 1
    jobs = collect_jobs(path, [sheet_name], spec, data_start)
    if not jobs:
        safe_print("错误：没有可处理的页签。")
        sys.exit(1)
    safe_print(f"  页签: {sheet_name}")
    safe_print(f"  插在 {spec.after_letter} 列后，新增 {len(spec.columns)} 列")
    for plan in spec.columns:
        formula_note = plan.formula if plan.formula else "空白列"
        color_note = f"#{plan.fill_hex}" if plan.fill_hex else "无色"
        safe_print(f"    {plan.new_header} [{color_note}] -> {formula_note}")
    if sys.platform == "win32":
        written, done_sheets = run_com(path, spec, jobs)
    else:
        written, done_sheets = run_openpyxl(path, spec, jobs)
    last_letter = get_column_letter(jobs[-1].after_col + 1)
    formula_cols = sum(1 for plan in spec.columns if plan.formula)
    kind = "公式列" if formula_cols == len(spec.columns) else ("混合列" if formula_cols else "空白列")
    safe_print(f"  完成：已插入至 {last_letter} 列（{kind}×{len(spec.columns)}），共写入 {written} 行公式")
    return last_letter, written, jobs[0].data_end


def run() -> None:
    t0 = time.time()
    safe_print("=" * 60)
    safe_print("新建公式列（18）")
    safe_print("=" * 60)

    path = require_excel_file("选择要插入公式列的 Excel")
    names = list_sheet_names(path)
    if not names:
        safe_print("错误：工作簿中没有页签。")
        sys.exit(1)
    sheet_name = pick_sheet(path)
    if sheet_name is None:
        safe_print("已取消。")
        sys.exit(0)

    header_row = pick_header_row(path, sheet_name)
    if header_row is None:
        safe_print("已取消。")
        sys.exit(0)

    backup = backup_file(path)
    safe_print(f"  已备份: {backup.name}")
    step = 0
    prefer_letter: str | None = None
    expect_header: str | None = None

    while True:
        columns, data_start, data_end = load_sheet_context(
            path,
            sheet_name,
            header_row,
            expect_header=expect_header,
        )
        expect_header = None
        if data_end < header_row + 1:
            safe_print(f"  页签「{sheet_name}」第 {data_start} 行起暂未扫到数据，仍可写入表头列。")
        else:
            safe_print(f"  表头第 {header_row} 行，数据第 {data_start}～{data_end} 行")
        safe_print(f"  当前表头: {format_headers_line(columns)}")
        spec = OperationDialog(
            path,
            sheet_name,
            columns,
            header_row,
            data_start,
            data_end,
            prefer_after_letter=prefer_letter,
        ).run()
        prefer_letter = None
        if spec is None:
            if step == 0:
                safe_print("已取消。")
                sys.exit(0)
            safe_print("已结束后续步骤。")
            break

        header_row = spec.header_row
        new_letter, written, last_row = apply_one_step(path, sheet_name, spec)
        step += 1
        refreshed, _, _ = load_sheet_context(
            path,
            sheet_name,
            header_row,
            expect_header=spec.columns[-1].new_header,
        )
        safe_print(f"  更新后表头: {format_headers_line(refreshed)}")
        formula_cols = sum(1 for plan in spec.columns if plan.formula)
        n = len(spec.columns)
        names = "、".join(plan.new_header for plan in spec.columns)
        kind = "公式列" if formula_cols == n else ("混合" if formula_cols else "空白列")
        summary = (
            f"第 {step} 步完成。\n"
            f"页签「{sheet_name}」已插入 {n} 列（{kind}：{names}），末列 {new_letter}。\n"
            f"共写入 {written} 行公式（第 {spec.header_row + 1}～{last_row} 行）。\n"
            f"当前表头：{format_headers_line(refreshed)}"
        )
        action = ask_next_step(sheet_name, summary)
        if action == "done":
            break
        if action == "other":
            next_sheet = pick_sheet(path, force_dialog=True, current=sheet_name)
            if next_sheet is None:
                break
            sheet_name = next_sheet
            next_header = pick_header_row(path, sheet_name)
            if next_header is None:
                break
            header_row = next_header
            prefer_letter = None
            expect_header = None
            safe_print(f"  切换到页签「{sheet_name}」，表头第 {header_row} 行")
        else:
            prefer_letter = new_letter
            expect_header = spec.columns[-1].new_header
            safe_print(f"  继续本页签「{sheet_name}」，已刷新表头")

    elapsed = time.time() - t0
    safe_print("-" * 60)
    if step:
        safe_print(f"全部完成，共 {step} 步。")
        safe_print(f"输出: {path}")
        safe_print(f"耗时: {elapsed:.1f} 秒")
    else:
        safe_print("没有写入。")


if __name__ == "__main__":
    from launcher_common import run_with_pause

    run_with_pause(run)
