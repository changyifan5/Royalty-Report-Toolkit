"""
大报表分段 + 快速行数扫描（普适模块）

所有按 Excel 数据行处理的脚本均应使用：
  - 每段最多 CHUNK_SIZE（10 万）行
  - 运行前确认
  - 超过一段时询问操作哪一段（第 1～10 万、第 100001～20 万 …）
  - 每段完成后弹窗询问是否继续下一段（可暂停，下次再跑）
  - 多段时写入 .chunkwork 工作文件，**全部段完成后**才输出最终文件

行数优先读 xlsx zip 内 dimension（毫秒级），失败再只读扫 A 列。
"""

from __future__ import annotations

import json
import math
import re
import shutil
import sys
import xml.etree.ElementTree as ET
import zipfile
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import messagebox, ttk

from excel_common import tk_mono_font, tk_ui_font

CHUNK_SIZE = 100_000
CHUNKWORK_MARKER = ".chunkwork"


def timestamp_backup_path(source: Path) -> Path:
    """时间戳备份路径：report.xlsx → report.bak-YYYYMMDD.xlsx（与 excel_common 同名规则）。"""
    stamp = datetime.now().strftime("%Y%m%d")
    return source.with_name(f"{source.stem}.bak-{stamp}{source.suffix}")


@dataclass
class ChunkProgress:
    script_key: str
    source: str
    chunk_total: int
    completed: list[int] = field(default_factory=list)
    visible_only: bool = False


def chunkwork_path(source: Path) -> Path:
    return source.with_name(f"{source.stem}{CHUNKWORK_MARKER}{source.suffix}")


def chunk_progress_path(source: Path, script_key: str) -> Path:
    return source.with_name(f"{source.stem}.{script_key}.chunkprogress.json")


def load_chunk_progress(source: Path, script_key: str) -> ChunkProgress | None:
    path = chunk_progress_path(source, script_key)
    if not path.exists():
        return None
    data = json.loads(path.read_text(encoding="utf-8"))
    return ChunkProgress(
        script_key=data["script_key"],
        source=data["source"],
        chunk_total=int(data["chunk_total"]),
        completed=[int(x) for x in data.get("completed", [])],
        visible_only=data.get("visible_only", False),
    )


def save_chunk_progress(progress: ChunkProgress, source: Path) -> None:
    path = chunk_progress_path(source, progress.script_key)
    payload = {
        "script_key": progress.script_key,
        "source": progress.source,
        "chunk_total": progress.chunk_total,
        "completed": sorted(set(progress.completed)),
    }
    if progress.visible_only:
        payload["visible_only"] = True
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def mark_chunk_completed(source: Path, script_key: str, chunk: RowChunk) -> ChunkProgress:
    progress = load_chunk_progress(source, script_key)
    if progress is None:
        progress = ChunkProgress(
            script_key=script_key,
            source=str(source.resolve()),
            chunk_total=chunk.chunk_total,
        )
    elif progress.chunk_total != chunk.chunk_total:
        print(
            f"错误：分段总数与上次不一致（上次 {progress.chunk_total}，本次 {chunk.chunk_total}）。"
            f"\n  请删除 {chunkwork_path(source).name} 和 {chunk_progress_path(source, script_key).name} 后从第 1 段重跑。"
        )
        sys.exit(1)

    if chunk.chunk_index not in progress.completed:
        progress.completed.append(chunk.chunk_index)
    save_chunk_progress(progress, source)
    return progress


def all_chunks_completed(progress: ChunkProgress) -> bool:
    return set(progress.completed) == set(range(1, progress.chunk_total + 1))


def remaining_chunk_indices(progress: ChunkProgress) -> list[int]:
    done = set(progress.completed)
    return [i for i in range(1, progress.chunk_total + 1) if i not in done]


def prepare_chunk_work(
    source: Path,
    script_key: str,
    chunk: RowChunk,
    backup_fn: Callable[[Path], Path],
    *,
    visible_only: bool = False,
) -> tuple[Path, Path]:
    """返回 (读写路径, 备份路径)。多段时中间段只写 .chunkwork，不碰原文件。"""
    if chunk.chunk_total == 1:
        if chunk.is_first:
            backup_path = backup_fn(source)
        else:
            backup_path = timestamp_backup_path(source)
        return source, backup_path

    work = chunkwork_path(source)
    if chunk.is_first and not work.exists():
        backup_path = backup_fn(source)
        shutil.copy2(source, work)
        save_chunk_progress(
            ChunkProgress(
                script_key=script_key,
                source=str(source.resolve()),
                chunk_total=chunk.chunk_total,
                visible_only=visible_only,
            ),
            source,
        )
        print(f"  已创建分段工作文件: {work.name}")
        return work, backup_path

    if not work.exists():
        print(f"错误：找不到分段工作文件 {work.name}，请先运行第 1 段。")
        sys.exit(1)

    progress = load_chunk_progress(source, script_key)
    if progress is None:
        print(f"错误：找不到进度文件，请删除 {work.name} 后从第 1 段重跑。")
        sys.exit(1)
    if progress.visible_only != visible_only:
        raise ValueError("本次处理范围与已有分段进度不一致，请重新打开脚本续跑。")
    if progress.chunk_total != chunk.chunk_total:
        print(
            f"错误：分段总数与上次不一致。请删除 {work.name} 和"
            f" {chunk_progress_path(source, script_key).name} 后从第 1 段重跑。"
        )
        sys.exit(1)

    backup_path = timestamp_backup_path(source)
    print(f"  使用分段工作文件: {work.name}")
    return work, backup_path


def finalize_chunk_output(
    source: Path,
    script_key: str,
    chunk: RowChunk,
    output_path: Path | None = None,
    *,
    remove_source: bool = False,
) -> tuple[bool, Path | None]:
    """
    标记本段完成；若全部分段已完成则输出最终文件并清理工作文件。
    返回 (是否已输出最终文件, 最终路径)。
    """
    if chunk.chunk_total == 1:
        dest = output_path or source
        if output_path and output_path.resolve() != source.resolve():
            dest.parent.mkdir(parents=True, exist_ok=True)
            if dest.exists():
                dest.unlink()
            shutil.copy2(source, dest)
            if remove_source and source.exists():
                source.unlink()
            print(f"  已输出: {dest}")
        return True, dest

    progress = mark_chunk_completed(source, script_key, chunk)
    work = chunkwork_path(source)

    if not all_chunks_completed(progress):
        print(f"  本段已写入工作文件: {work.name}")
        print(f"  已完成段: {sorted(progress.completed)}")
        print(f"  还需完成段: {remaining_chunk_indices(progress)}")
        print("  全部段完成后才会输出最终文件。")
        return False, None

    dest = output_path or source
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists():
        dest.unlink()
    shutil.copy2(work, dest)
    work.unlink(missing_ok=True)
    chunk_progress_path(source, script_key).unlink(missing_ok=True)
    if remove_source and source.exists() and source.resolve() != dest.resolve():
        source.unlink()
    print(f"  全部 {chunk.chunk_total} 段已完成，已输出: {dest}")
    return True, dest

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CELL_REF_RE = re.compile(r"^([A-Z]+)(\d+)$", re.IGNORECASE)


@dataclass(frozen=True)
class RowChunk:
    chunk_index: int
    chunk_total: int
    data_start_row: int
    data_end_row: int
    total_data_rows: int
    data_row_start: int
    data_row_end: int

    @property
    def is_first(self) -> bool:
        return self.chunk_index == 1

    @property
    def is_final(self) -> bool:
        return self.chunk_index == self.chunk_total


def _row_from_cell_ref(ref: str) -> int:
    m = _CELL_REF_RE.match(ref.strip())
    return int(m.group(2)) if m else 0


def _row_from_dimension(ref: str) -> int:
    ref = ref.strip().upper()
    if not ref:
        return 0
    if ":" in ref:
        return _row_from_cell_ref(ref.split(":", 1)[1])
    return _row_from_cell_ref(ref)


def fast_sheet_last_row(path: Path, sheet_index: int = 0) -> int:
    """从 xlsx/xlsm zip 读取工作表 dimension 末行。"""
    with zipfile.ZipFile(path, "r") as zf:
        names = zf.namelist()
        wb_xml = zf.read("xl/workbook.xml")
        rels_xml = zf.read("xl/_rels/workbook.xml.rels")
        wb_root = ET.fromstring(wb_xml)
        rels_root = ET.fromstring(rels_xml)

        rel_map = {
            el.get("Id"): el.get("Target")
            for el in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship")
        }

        sheets = wb_root.findall(f".//{{{MAIN_NS}}}sheet")
        if sheet_index >= len(sheets):
            return 0

        rid = sheets[sheet_index].get(f"{{{REL_NS}}}id")
        target = rel_map.get(rid or "", "")
        if target.startswith("/"):
            sheet_path = target.lstrip("/")
        elif target.startswith("xl/"):
            sheet_path = target
        else:
            sheet_path = f"xl/{target}"

        if sheet_path not in names:
            alt = sheet_path.replace("xl/", "xl/worksheets/", 1) if "/worksheets/" not in sheet_path else sheet_path
            if alt not in names:
                for n in names:
                    if n.endswith(target.split("/")[-1]):
                        sheet_path = n
                        break
            else:
                sheet_path = alt

        sheet_xml = zf.read(sheet_path)
        root = ET.fromstring(sheet_xml)
        dim = root.find(f"{{{MAIN_NS}}}dimension")
        if dim is None:
            return 0
        ref = dim.get("ref") or ""
        return _row_from_dimension(ref)


def count_data_rows_slow(path: Path, data_start_row: int, sheet_index: int = 0) -> int:
    from openpyxl import load_workbook

    from row_order_common import find_last_data_row

    wb = load_workbook(path, read_only=True, keep_links=False)
    ws = wb.worksheets[sheet_index]
    last_row = find_last_data_row(ws, data_start_row=data_start_row)
    wb.close()
    if last_row < data_start_row:
        return 0
    return last_row - data_start_row + 1


def count_data_rows(path: Path, data_start_row: int = 7, sheet_index: int = 0) -> int:
    try:
        last_row = fast_sheet_last_row(path, sheet_index)
    except Exception:
        last_row = 0
    if last_row < data_start_row:
        return count_data_rows_slow(path, data_start_row, sheet_index)
    return last_row - data_start_row + 1


def compute_chunks(
    total_data_rows: int, excel_data_start_row: int, chunk_size: int = CHUNK_SIZE
) -> list[RowChunk]:
    if total_data_rows <= 0:
        return []
    chunk_total = math.ceil(total_data_rows / chunk_size)
    chunks: list[RowChunk] = []
    for i in range(chunk_total):
        data_row_start = i * chunk_size + 1
        data_row_end = min((i + 1) * chunk_size, total_data_rows)
        row_offset_start = data_row_start - 1
        row_offset_end = data_row_end - 1
        chunks.append(
            RowChunk(
                chunk_index=i + 1,
                chunk_total=chunk_total,
                data_start_row=excel_data_start_row + row_offset_start,
                data_end_row=excel_data_start_row + row_offset_end,
                total_data_rows=total_data_rows,
                data_row_start=data_row_start,
                data_row_end=data_row_end,
            )
        )
    return chunks


def format_chunk_label(chunk: RowChunk) -> str:
    return (
        f"[{chunk.chunk_index}] 第 {chunk.data_row_start:,}～{chunk.data_row_end:,} 条 "
        f"（Excel 第 {chunk.data_start_row:,}～{chunk.data_end_row:,} 行）"
    )


def print_chunk_banner(script_name: str, chunk: RowChunk) -> None:
    print()
    print("-" * 50)
    print(f"{script_name} — 本段 {chunk.chunk_index}/{chunk.chunk_total}")
    print(format_chunk_label(chunk))
    print(f"报表数据总行数: {chunk.total_data_rows:,}")
    print("-" * 50)


def row_in_chunk(row: int, chunk: RowChunk) -> bool:
    return chunk.data_start_row <= row <= chunk.data_end_row


def load_file_chunks(
    path: Path,
    excel_data_start_row: int = 7,
    sheet_index: int = 0,
) -> list[RowChunk]:
    """统计数据行并计算全部分段（不弹窗）。"""
    print(f"  统计数据行数: {path.name} ...")
    total = count_data_rows(path, excel_data_start_row, sheet_index)
    if total <= 0:
        print(f"错误：第 {excel_data_start_row} 行起没有数据（A 列为空）。")
        sys.exit(1)

    chunks = compute_chunks(total, excel_data_start_row)
    print(f"  数据行: {total:,}，分为 {len(chunks)} 段（每段最多 {CHUNK_SIZE:,} 行）")
    return chunks


def next_pending_chunk(
    chunks: list[RowChunk],
    path: Path,
    script_key: str,
) -> RowChunk | None:
    progress = load_chunk_progress(path, script_key)
    done = set(progress.completed) if progress else set()
    for ch in chunks:
        if ch.chunk_index not in done:
            return ch
    return None


def prompt_continue_next_chunk(
    script_name: str,
    completed_chunk: RowChunk,
    next_chunk: RowChunk,
) -> bool:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    ok = messagebox.askyesno(
        script_name,
        f"本段 {completed_chunk.chunk_index}/{completed_chunk.chunk_total} 已完成。\n\n"
        f"是否继续处理下一段？\n"
        f"{format_chunk_label(next_chunk)}",
        parent=root,
    )
    root.destroy()
    return ok


def prompt_next_chunk_after(
    script_name: str,
    path: Path,
    script_key: str,
    chunks: list[RowChunk],
    completed_chunk: RowChunk,
    finalized: bool,
) -> RowChunk | None:
    """本段完成后询问是否继续；返回下一段或 None。"""
    if finalized or len(chunks) <= 1:
        return None

    next_chunk = next_pending_chunk(chunks, path, script_key)
    if next_chunk is None:
        return None

    if prompt_continue_next_chunk(script_name, completed_chunk, next_chunk):
        print_chunk_banner(script_name, next_chunk)
        return next_chunk

    progress = load_chunk_progress(path, script_key)
    if progress:
        print()
        print(f"已暂停。已完成段: {sorted(progress.completed)}")
        print(f"还需完成段: {remaining_chunk_indices(progress)}")
        print("下次运行脚本可继续选择未完成段。")
    return None


def prompt_visible_rows_only(script_name: str, path: Path, script_key: str) -> bool | None:
    progress = load_chunk_progress(path, script_key)
    if progress is not None:
        scope = "仅处理可见行" if progress.visible_only else "处理全部行"
        print(f"  继续上次分段，处理范围：{scope}")
        return progress.visible_only

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        return messagebox.askyesnocancel(
            script_name,
            "是否仅处理第一个页签的可见行？\n\n"
            "是：仅处理可见行，跳过筛选或手动隐藏的行。\n"
            "否：处理全部数据行。\n\n"
            "以 Excel 文件上次保存的行隐藏状态为准。",
            parent=root,
        )
    finally:
        root.destroy()


def select_chunk_rows(sheet, chunk: RowChunk, visible_only: bool) -> range | list[int]:
    rows = range(chunk.data_start_row, chunk.data_end_row + 1)
    if not visible_only:
        return rows
    hidden = {row for row, dimension in sheet.row_dimensions.items() if dimension.hidden}
    return [row for row in rows if row not in hidden]


def prompt_row_chunk(
    script_name: str,
    path: Path,
    excel_data_start_row: int = 7,
    sheet_index: int = 0,
    script_key: str = "",
    chunks: list[RowChunk] | None = None,
) -> RowChunk | None:
    if chunks is None:
        chunks = load_file_chunks(path, excel_data_start_row, sheet_index)

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    if not messagebox.askyesno(
        script_name,
        f"文件: {path.name}\n"
        f"数据行: {chunks[0].total_data_rows:,}\n"
        f"分段: {len(chunks)} 段（每段最多 {CHUNK_SIZE:,} 行）\n\n"
        f"是否继续？",
        parent=root,
    ):
        root.destroy()
        return None

    if len(chunks) == 1:
        root.destroy()
        return chunks[0]

    selected: list[RowChunk] = []

    picker = tk.Toplevel(root)
    picker.title(f"{script_name} — 选择段")
    picker.attributes("-topmost", True)
    picker.resizable(False, False)

    frame = ttk.Frame(picker, padding=12)
    frame.grid(row=0, column=0)

    ttk.Label(
        frame,
        text=f"请选择要处理的 10 万行段（共 {len(chunks)} 段）：",
        font=tk_ui_font(10),
    ).grid(row=0, column=0, sticky="w", pady=(0, 8))

    listbox = tk.Listbox(frame, width=72, height=min(14, len(chunks)), font=tk_mono_font(10))
    listbox.grid(row=1, column=0)
    progress = load_chunk_progress(path, script_key) if script_key else None
    done_set = set(progress.completed) if progress else set()
    default_sel = 0
    for i, ch in enumerate(chunks):
        mark = " ✓" if ch.chunk_index in done_set else ""
        listbox.insert(tk.END, format_chunk_label(ch) + mark)
        if ch.chunk_index not in done_set and default_sel == 0 and done_set:
            default_sel = i
    listbox.selection_set(default_sel)

    def confirm() -> None:
        sel = listbox.curselection()
        if not sel:
            messagebox.showwarning("提示", "请先选择一段。", parent=picker)
            return
        selected.append(chunks[sel[0]])
        picker.destroy()
        root.destroy()

    def cancel() -> None:
        picker.destroy()
        root.destroy()

    btn = ttk.Frame(frame)
    btn.grid(row=2, column=0, pady=(10, 0), sticky="e")
    ttk.Button(btn, text="确定", command=confirm, width=10).grid(row=0, column=0, padx=(0, 6))
    ttk.Button(btn, text="取消", command=cancel, width=10).grid(row=0, column=1)

    picker.bind("<Return>", lambda _e: confirm())
    picker.bind("<Escape>", lambda _e: cancel())
    picker.grab_set()
    root.wait_window(picker)

    return selected[0] if selected else None


def print_chunk_next_hint(chunk: RowChunk, finalized: bool) -> None:
    if finalized:
        return
    print()
    if chunk.chunk_total == 1:
        return
    print(f"提示：本段（{chunk.chunk_index}/{chunk.chunk_total}）已写入工作文件，请继续处理其余段。")
    print("      全部段完成后才会输出最终文件。")
