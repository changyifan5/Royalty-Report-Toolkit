"""Excel 报表脚本公共基础设施（01/02/04/07/09/14 等共用）。

本模块收编原先散落在各编号脚本里的重复函数与业务常量，并提供：
  - 业务常量：表头行 / 数据起始行 / Excel 后缀 / A～D 颜色约定
  - 公共函数：normalize_key / find_header_last_col / read_header_map /
    backup_file（带时间戳命名）/ safe_print / resolve_excel_path
  - 统一助手：写日志（沿用 08 的追加写法）、错误弹窗（无 GUI 时静默降级）、
    tkinter 字体 fallback、Windows PowerShell(.ps1) 集中调用

注意：find_last_data_row 仍在 row_order_common.py（供 05 家族使用），本模块不重复定义。
"""

from __future__ import annotations

import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------------------
# 业务常量
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent

# 标准报表布局：第 6 行表头，第 7 行起为数据行。
# 注意：01/14 的「比对文件 B」表头/数据行不同（见各脚本本地常量），不可强行套用。
HEADER_ROW = 6
DATA_START_ROW = 7
HEADER_SCAN_LIMIT = 500
KEY_COL = 1

# 脚本统一处理的前 N 列（A～D）。
COLOR_CHECK_COLS = 4

EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm"}

# 写回时使用的红 / 蓝字体色值（04 退回修改标记在用；不含 alpha 通道）。
COLOR_RETURN = "FF0000"   # 需退回
COLOR_RATIO = "0000FF"    # 需修改词曲比例

# 是否 Windows 平台（Windows 走 Excel COM / PowerShell，Mac 走 openpyxl）。
IS_WINDOWS = sys.platform == "win32"


# ---------------------------------------------------------------------------
# tkinter 字体 fallback（Windows 有 YaHei/Consolas，Mac 替换为 PingFang/Menlo）
# ---------------------------------------------------------------------------
if IS_WINDOWS:
    UI_FONT_FAMILY = "Microsoft YaHei UI"
    MONO_FONT_FAMILY = "Consolas"
else:
    UI_FONT_FAMILY = "PingFang SC"
    MONO_FONT_FAMILY = "Menlo"


def tk_ui_font(size: int, weight: str | None = None) -> tuple:
    """构造 tkinter 界面字体，如 tk_ui_font(10) / tk_ui_font(9, "bold")。"""
    return (UI_FONT_FAMILY, size) + ((weight,) if weight else ())


def tk_mono_font(size: int) -> tuple:
    """构造 tkinter 等宽字体（列表 / 数据显示）。"""
    return (MONO_FONT_FAMILY, size)


# ---------------------------------------------------------------------------
# 公共函数
# ---------------------------------------------------------------------------
def normalize_key(value: object) -> str | None:
    """单元格值去首尾空格后作为查找键；空串归一为 None。"""
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def find_header_last_col(sheet, header_row: int = HEADER_ROW, limit: int = HEADER_SCAN_LIMIT) -> int:
    """扫描 header_row 行，返回最后一个非空表头所在列号（1-based）；无表头返回 0。"""
    last = 0
    for col in range(1, limit + 1):
        val = sheet.cell(row=header_row, column=col).value
        if val is not None and str(val).strip():
            last = col
    return last


def read_header_map(sheet, header_row: int = HEADER_ROW) -> dict[str, int]:
    """表头文字（去空格）→ 列号（1-based）。"""
    mapping: dict[str, int] = {}
    for col in range(1, HEADER_SCAN_LIMIT + 1):
        val = sheet.cell(row=header_row, column=col).value
        if val is not None and str(val).strip():
            mapping[str(val).strip()] = col
    return mapping


def resolve_excel_path(path: Path, label: str) -> Path:
    """路径是文件就校验后缀；是文件夹则要求里面只有一个 Excel。否则打印错误并退出。"""
    if not path.exists():
        print(f"错误：{label} 路径不存在: {path}")
        sys.exit(1)

    if path.is_file():
        if path.suffix.lower() not in EXCEL_SUFFIXES:
            print(f"错误：{label} 不是支持的 Excel 格式: {path}")
            sys.exit(1)
        return path

    candidates = sorted(
        f
        for f in path.iterdir()
        if f.is_file() and f.suffix.lower() in EXCEL_SUFFIXES and not f.name.startswith("~$")
    )
    if not candidates:
        print(f"错误：{label} 文件夹内没有找到 Excel 文件: {path}")
        sys.exit(1)
    if len(candidates) > 1:
        print(f"错误：{label} 文件夹内有多个 Excel 文件，请只保留一个:")
        for f in candidates:
            print(f"  - {f}")
        sys.exit(1)

    print(f"  已自动选用 {label}: {candidates[0].name}")
    return candidates[0]


def timestamp_backup_path(path: Path, when: datetime | None = None) -> Path:
    """生成带时间戳的备份路径：report.xlsx → report.bak-YYYYMMDD.xlsx。

    命名理由：仓库 .gitignore 同时有 *.bak 与 *.xlsx 规则；*.bak-YYYYMMDD 后缀不会被
    *.bak 命中、又以 .xlsx 结尾落在 *.xlsx 忽略规则内，避免业务备份被 git 误提交。
    """
    stamp = (when or datetime.now()).strftime("%Y%m%d")
    return path.with_name(f"{path.stem}.bak-{stamp}{path.suffix}")


def backup_file(path: Path) -> Path:
    """把 path 复制为带时间戳的备份文件（同日重复运行覆盖同名备份），返回备份路径。"""
    backup_path = timestamp_backup_path(path)
    shutil.copy2(path, backup_path)
    return backup_path


def safe_print(text: str) -> None:
    """打印文本；Windows 控制台编码异常时用 replace 容错，避免中断长任务。"""
    try:
        print(text)
    except UnicodeEncodeError:
        enc = sys.stdout.encoding or "utf-8"
        print(text.encode(enc, errors="replace").decode(enc, errors="replace"))


# ---------------------------------------------------------------------------
# 日志（沿用 08 清理Chrome下载.py 的追加写法）
# ---------------------------------------------------------------------------
def append_log(lines: list[str], log_file: Path) -> None:
    """向 log_file 追加一段带时间戳的日志块（UTF-8，追加模式）。"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with log_file.open("a", encoding="utf-8") as f:
        f.write(f"\n=== {timestamp} ===\n")
        f.write("\n".join(lines))
        f.write("\n")


def log_error(message: str, log_file: Path | None = None) -> None:
    """记录一条错误到日志文件（若给定）；同时打印到控制台。"""
    safe_print(f"错误：{message}")
    if log_file is not None:
        append_log([f"[ERROR] {message}"], log_file)


# ---------------------------------------------------------------------------
# 错误弹窗（无显示环境时静默降级，不抛异常）
# ---------------------------------------------------------------------------
def show_error_popup(title: str, message: str) -> None:
    """弹出错误对话框；无 GUI / 弹窗失败时静默跳过（不影响命令行日志）。"""
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        messagebox.showerror(title, message, parent=root)
        root.destroy()
    except Exception:
        # 无显示环境 / tkinter 不可用时降级为只打印。
        safe_print(f"[{title}] {message}")


# ---------------------------------------------------------------------------
# Windows PowerShell(.ps1) 集中调用助手
# ---------------------------------------------------------------------------
def powershell_encoding() -> str:
    """Windows 下 COM 脚本控制台输出为 gbk，其余平台为 utf-8。"""
    return "gbk" if IS_WINDOWS else "utf-8"


def run_powershell(
    ps1: Path,
    extra_args: list[str],
    log_file: Path,
    *,
    encoding: str | None = None,
) -> str:
    """集中调用 Excel COM PowerShell 脚本并把 stdout 落日志。

    extra_args 形如 ["-ManifestPath", str(manifest.resolve())]（由调用方组装，保持各自
    原命令行逐字等价）。返回 PowerShell 标准输出文本。
    """
    cmd = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(ps1),
        *extra_args,
    ]
    proc_encoding = encoding or powershell_encoding()
    import subprocess

    proc = subprocess.run(
        cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        encoding=proc_encoding, errors="replace",
    )
    log_file.write_text(proc.stdout or "", encoding="utf-8")
    return proc.stdout or ""


# ---------------------------------------------------------------------------
# 快速读取页签名（zip 直读 workbook.xml，避免 load_workbook 只为看页签）
# ---------------------------------------------------------------------------
_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def list_sheet_names_fast(path: Path) -> list[str]:
    """只读取 xlsx 内 xl/workbook.xml 提取全部页签名（毫秒级，不解析单元格）。"""
    import xml.etree.ElementTree as ET

    with zipfile.ZipFile(path, "r") as zf:
        wb_xml = zf.read("xl/workbook.xml")
    root = ET.fromstring(wb_xml)
    return [el.get("name") for el in root.findall(f".//{{{_NS}}}sheet")]
