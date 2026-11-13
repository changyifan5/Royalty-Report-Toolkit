"""
多文件 F3 页签汇总脚本（代码库 19）

多选 Excel，对每个文件汇总全部页签（页签名称为「总」的除外）的 F3 数值。
F3 为公式时读取计算结果（优先 openpyxl 读缓存；无缓存时 Windows 用 Excel COM 求值）。
结果仅在控制台显示，不写文件。

Windows 使用 Excel COM；Mac 使用 openpyxl。
运行：python "19 F3汇总.py"
"""

from __future__ import annotations

import json
import sys
import time
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import load_workbook

from file_picker_common import pick_excel_files

BASE_DIR = Path(__file__).resolve().parent
PS1_SCRIPT = BASE_DIR / "19_f3_sheet_sum_com.ps1"
MANIFEST_FILE = BASE_DIR / "19_f3_manifest.json"
RESULT_FILE = BASE_DIR / "19_f3_result.json"
LOG_FILE = BASE_DIR / "19_com.log"
SKIP_SHEET_NAME = "总"
CELL_ROW = 3
CELL_COL = 6  # F


@dataclass
class SheetItem:
    sheet_name: str
    f3_value: float
    raw_display: str
    has_formula: bool = False


@dataclass
class FileSummary:
    path: Path
    sheets: list[SheetItem] = field(default_factory=list)
    skipped_sheets: list[str] = field(default_factory=list)
    invalid_sheets: list[str] = field(default_factory=list)

    @property
    def total(self) -> float:
        return sum(item.f3_value for item in self.sheets)


def to_number(value) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    if text.startswith("#"):
        return None
    try:
        return float(text)
    except ValueError:
        return None


def format_full_number(value: float) -> str:
    """完整显示数值，不用 :g 缩短有效位。"""
    text = repr(value)
    if "e" in text.lower():
        return format(Decimal(str(value)), "f")
    return text


def display_from_raw(raw) -> str:
    if raw is None:
        return "(空)"
    if isinstance(raw, str):
        text = raw.strip().replace(",", "")
        if not text:
            return "(空)"
        if to_number(text) is not None:
            return text
        return text
    if isinstance(raw, (int, float)):
        return format_full_number(float(raw))
    text = str(raw).strip()
    return text or "(空)"


def sum_full_values(sheets: list[SheetItem], invalid: set[str]) -> str:
    total = Decimal(0)
    for sheet in sheets:
        if sheet.sheet_name in invalid:
            continue
        try:
            total += Decimal(sheet.raw_display.replace(",", ""))
        except (InvalidOperation, AttributeError):
            total += Decimal(str(sheet.f3_value))
    return format(total, "f")


def cell_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return str(value).strip()


def add_sheet_item(
    summary: FileSummary,
    sheet_name: str,
    raw,
    *,
    has_formula: bool = False,
) -> None:
    num = to_number(raw)
    if num is None:
        summary.invalid_sheets.append(sheet_name)
        summary.sheets.append(
            SheetItem(
                sheet_name=sheet_name,
                f3_value=0.0,
                raw_display=cell_text(raw) or "(空)",
                has_formula=has_formula,
            )
        )
        return

    summary.sheets.append(
        SheetItem(
            sheet_name=sheet_name,
            f3_value=num,
            raw_display=display_from_raw(raw),
            has_formula=has_formula,
        )
    )


def summarize_file_openpyxl(path: Path) -> tuple[FileSummary, bool]:
    """返回 (汇总, 是否需要 COM 计算公式)。有公式但无缓存值时需要 COM。"""
    summary = FileSummary(path=path)
    needs_com = False
    wb_formula = load_workbook(path, read_only=True, data_only=False, keep_links=False)
    wb_values = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        for ws in wb_formula.worksheets:
            name = ws.title
            if name == SKIP_SHEET_NAME:
                summary.skipped_sheets.append(name)
                continue

            raw_formula = ws.cell(row=CELL_ROW, column=CELL_COL).value
            has_formula = isinstance(raw_formula, str) and raw_formula.startswith("=")
            raw_value = wb_values[name].cell(row=CELL_ROW, column=CELL_COL).value
            raw = raw_value if has_formula else raw_formula
            if has_formula and to_number(raw_value) is None:
                needs_com = True
            add_sheet_item(summary, name, raw, has_formula=has_formula)
    finally:
        wb_formula.close()
        wb_values.close()
    return summary, needs_com


def write_manifest(paths: list[Path]) -> None:
    payload = {
        "paths": [str(p.resolve()) for p in paths],
        "skip_sheet_name": SKIP_SHEET_NAME,
        "cell_row": CELL_ROW,
        "cell_col": CELL_COL,
        "result_path": str(RESULT_FILE.resolve()),
    }
    MANIFEST_FILE.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")


def normalize_com_file_items(payload) -> list[dict]:
    """PowerShell ConvertTo-Json 在仅 1 个文件时返回对象而非数组。"""
    if isinstance(payload, dict):
        if "path" in payload or "sheets" in payload:
            return [payload]
        return [payload]
    if isinstance(payload, list):
        return [item for item in payload if isinstance(item, dict)]
    print("错误：COM 返回格式无法识别")
    sys.exit(1)


def load_com_payload(lines: list[str]):
    for line in lines:
        if line.startswith("RESULT_FILE,"):
            result_path = Path(line.split(",", 1)[1].strip())
            if not result_path.exists():
                print(f"错误：找不到 COM 结果文件: {result_path}")
                sys.exit(1)
            return json.loads(result_path.read_text(encoding="utf-8"))
        if line.startswith("RESULT_JSON,"):
            return json.loads(line[len("RESULT_JSON,") :])
    return None


def parse_com_summaries(lines: list[str], paths: list[Path]) -> list[FileSummary]:
    payload = load_com_payload(lines)

    if payload is None:
        print("错误：COM 未返回汇总结果，日志:", LOG_FILE)
        sys.exit(1)

    path_by_resolved = {str(p.resolve()): p for p in paths}
    summaries: list[FileSummary] = []

    for idx, item in enumerate(normalize_com_file_items(payload)):
        resolved = str(item.get("path", ""))
        path = path_by_resolved.get(resolved)
        if path is None and idx < len(paths):
            path = paths[idx]
        elif path is None:
            path = Path(resolved)
        summary = FileSummary(path=path)
        summary.skipped_sheets = list(item.get("skipped") or [])

        for sheet in item.get("sheets") or []:
            name = str(sheet.get("name", ""))
            has_formula = bool(sheet.get("has_formula"))
            if sheet.get("invalid"):
                summary.invalid_sheets.append(name)
                summary.sheets.append(
                    SheetItem(
                        sheet_name=name,
                        f3_value=0.0,
                        raw_display=str(sheet.get("raw") or "(空)"),
                        has_formula=has_formula,
                    )
                )
                continue

            value = float(sheet.get("value") or 0.0)
            raw_text = sheet.get("raw")
            summary.sheets.append(
                SheetItem(
                    sheet_name=name,
                    f3_value=value,
                    raw_display=display_from_raw(raw_text if raw_text not in (None, "") else value),
                    has_formula=has_formula,
                )
            )
        summaries.append(summary)

    return summaries


def summarize_files_com(paths: list[Path]) -> list[FileSummary]:
    if not PS1_SCRIPT.exists():
        print(f"错误：找不到 COM 脚本: {PS1_SCRIPT}")
        sys.exit(1)

    from com_runner_common import stream_com

    write_manifest(paths)
    cmd = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(PS1_SCRIPT),
        "-ManifestPath",
        str(MANIFEST_FILE),
    ]
    print("  使用 Excel COM 计算公式（仅无缓存文件）...")
    returncode, lines = stream_com(
        cmd,
        log_file=LOG_FILE,
        progress_title="F3 页签汇总",
        stage_labels={"open": "正在启动 Excel..."},
        progress_labels={"files": "文件进度"},
        timing_labels={"read_ms": "读取 F3"},
    )
    if returncode != 0:
        print()
        print("错误：COM 读取失败，日志:", LOG_FILE)
        sys.exit(1)
    return parse_com_summaries(lines, paths)


def print_summaries(summaries: list[FileSummary]) -> str:
    grand_parts: list[Decimal] = []
    for summary in summaries:
        print()
        print(f"【{summary.path.name}】")
        invalid = set(summary.invalid_sheets)
        for sheet in summary.sheets:
            tag = "（公式）" if sheet.has_formula else ""
            invalid_note = "（非数值，按 0 计）" if sheet.sheet_name in invalid else ""
            print(f"  {sheet.sheet_name}: {sheet.raw_display}{tag}{invalid_note}")
        for name in summary.skipped_sheets:
            print(f"  {name}: （已跳过，页签名=总）")
        file_total = sum_full_values(summary.sheets, invalid)
        print(f"  小计: {file_total}")
        try:
            grand_parts.append(Decimal(file_total.replace(",", "")))
        except InvalidOperation:
            grand_parts.append(Decimal(0))

    grand_total = format(sum(grand_parts, Decimal(0)), "f")
    return grand_total


def run_summarize() -> None:
    t0 = time.time()
    paths = pick_excel_files("选择要汇总 F3 的 Excel（可多选）")
    if not paths:
        print("已取消。")
        sys.exit(0)

    print(f"已选 {len(paths)} 个文件")
    summaries_by_path: dict[Path, FileSummary] = {}
    com_paths: list[Path] = []

    print("  使用 openpyxl 快速读取...")
    for idx, path in enumerate(paths, start=1):
        summary, needs_com = summarize_file_openpyxl(path)
        if needs_com and sys.platform == "win32":
            com_paths.append(path)
            print(f"  [{idx}/{len(paths)}] {path.name}：公式无缓存，待 Excel 计算")
        else:
            summaries_by_path[path] = summary
            if needs_com:
                print(f"  [{idx}/{len(paths)}] {path.name}：公式无缓存（非 Windows，结果可能为空）")
            else:
                print(f"  [{idx}/{len(paths)}] {path.name}：已读取")

    if com_paths:
        print(f"  共 {len(com_paths)} 个文件需 Excel COM 计算公式...")
        for summary in summarize_files_com(com_paths):
            summaries_by_path[summary.path] = summary

    summaries = [summaries_by_path[path] for path in paths]
    grand_total = print_summaries(summaries)

    print()
    print("=" * 50)
    print("F3 页签汇总完成")
    print("=" * 50)
    print(f"  文件数: {len(summaries)}")
    print(f"  全部 F3 合计: {grand_total}")
    if sys.platform != "win32":
        print("  提示：Mac 上公式若显示为空，请用 Windows 运行或在 Excel 中打开保存后再试。")
    print(f"  总耗时: {time.time() - t0:.1f} 秒")


if __name__ == "__main__":
    from launcher_common import run_with_pause

    run_with_pause(run_summarize)
