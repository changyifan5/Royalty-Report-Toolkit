"""
按页签全量同步脚本（代码库 06）

不存在单独的汇总文件 A；只提交 A_1 或 A-1（首屏页签对应的文件）。
文件名中 base 与页签名之间可用 _ 或 -（如 A_1.xlsx、A-1.xlsx，选定后同批统一）。
从 A_1 读取全部页签，同步到同目录每个 base_页签名 / base-页签名.xlsx：
  - 每个 A_i 含全部页签，内容一致
  - 本文件对应页签排第一，其余按 A_1 内页签顺序排列
  - 例：A_1 内 [1,2,3] → A_2 为 [2,1,3]

运行时：弹窗选择 A_1 或 A-1 文件（同目录其他页签文件会一并同步）。
目标不存在则新建。

Windows 使用 Excel COM；Mac 使用 openpyxl（文件复制 + 重排页签）。
运行：python "06 按页签另存.py"
"""

from __future__ import annotations

import json
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path

from file_picker_common import pick_excel_file

from openpyxl import load_workbook

BASE_DIR = Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from excel_sheet_copy import (
    copy_sheet_between_workbooks,
    reorder_workbook_sheets,
)
PS1_SCRIPT = BASE_DIR / "06_export_sheets_com.ps1"
MANIFEST_FILE = BASE_DIR / "06_export_manifest.json"
TEMP_SOURCE = BASE_DIR / "06_sync_source_temp.xlsx"
LOG_FILE = BASE_DIR / "06_com.log"

EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm"}
INVALID_SHEET_CHARS = re.compile(r'[\\/:*?"<>|]')
NAME_SEPARATORS = ("_", "-")
SKIP_FILENAMES = frozenset(
    {
        TEMP_SOURCE.name,
        "06_sync_source_temp.xlsx",
    }
)
XLS_OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def safe_print(text: str) -> None:
    try:
        print(text)
    except UnicodeEncodeError:
        print(text.encode(sys.stdout.encoding or "utf-8", errors="replace").decode(sys.stdout.encoding or "utf-8", errors="replace"))


def pick_a1_file(initial_dir: Path | None = None) -> Path | None:
    return pick_excel_file("请选择 A_1 或 A-1（首屏页签对应的文件）", initial_dir=initial_dir)


def list_excels(folder: Path) -> list[Path]:
    return sorted(
        f
        for f in folder.iterdir()
        if f.is_file()
        and f.suffix.lower() in EXCEL_SUFFIXES
        and not f.name.startswith("~$")
        and f.name not in SKIP_FILENAMES
        and ".chunkwork" not in f.stem
    )


def describe_invalid_excel(path: Path) -> str:
    try:
        head = path.read_bytes()[:8]
    except OSError as exc:
        return f"无法读取文件: {exc}"
    if len(head) < 4:
        return "文件为空或过小，不是有效的 Excel"
    if head[:2] == b"PK":
        return "扩展名是 Excel，但内部结构损坏或不完整（可尝试用 Excel 打开后另存为 .xlsx）"
    if head[:8] == XLS_OLE_MAGIC:
        return "这是旧版 .xls 格式，请用 Excel 另存为 .xlsx 后再运行"
    return "不是有效的 Excel 文件（可能损坏、改过后缀，或下载未完成）"


def try_read_sheet_names(path: Path) -> tuple[list[str] | None, str | None]:
    """读取页签名；(names, None) 成功，(None, 原因) 失败。"""
    import zipfile

    try:
        with zipfile.ZipFile(path) as archive:
            if "[Content_Types].xml" not in archive.namelist():
                return None, describe_invalid_excel(path)
    except zipfile.BadZipFile:
        return None, describe_invalid_excel(path)
    except OSError as exc:
        return None, f"无法打开文件: {exc}"

    try:
        wb = load_workbook(path, read_only=True, keep_links=False)
        names = list(wb.sheetnames)
        wb.close()
    except Exception as exc:
        return None, f"{describe_invalid_excel(path)}（{exc}）"

    if not names:
        return None, "工作簿中没有页签"
    return names, None


def read_sheet_names(path: Path) -> list[str]:
    names, err = try_read_sheet_names(path)
    if err:
        print()
        print("=" * 50)
        print("错误：无法读取 Excel 文件")
        print("=" * 50)
        print(f"  文件: {path.name}")
        print(f"  原因: {err}")
        sys.exit(1)
    assert names is not None
    return names


def sanitize_sheet_name(name: str) -> str:
    cleaned = INVALID_SHEET_CHARS.sub("_", name.strip())
    return cleaned or "Sheet"


def suffix_tokens(primary_sheet: str) -> list[tuple[str, str]]:
    """文件名末尾「分隔符+页签名」候选，用于匹配与解析 base。"""
    primary_safe = sanitize_sheet_name(primary_sheet)
    tokens: list[tuple[str, str]] = []
    for name in (primary_sheet, primary_safe):
        for sep in NAME_SEPARATORS:
            tokens.append((f"{sep}{name}", sep))
    return tokens


def stem_matches_primary(stem: str, primary_sheet: str) -> bool:
    for token, _sep in suffix_tokens(primary_sheet):
        if stem.endswith(token) and len(stem) > len(token):
            return True
    return False


def parse_base_name(stem: str, primary_sheet: str) -> tuple[str, str]:
    for token, sep in suffix_tokens(primary_sheet):
        if stem.endswith(token) and len(stem) > len(token):
            return stem[: -len(token)], sep
    print("错误：文件名与首屏页签不匹配，请选 A_1 / A-1（文件名应以 _或- + 首屏页签名 结尾）。")
    print(f"  文件: {stem}")
    print(f"  首屏页签: {primary_sheet!r}")
    print(f"  期望文件名类似: {{base}}_{primary_sheet}.xlsx 或 {{base}}-{primary_sheet}.xlsx")
    sys.exit(1)


def target_filename(
    base: str, sheet_name: str, suffix: str, sep: str, used_safe: dict[str, int]
) -> str:
    safe = sanitize_sheet_name(sheet_name)
    if safe in used_safe:
        used_safe[safe] += 1
        file_part = f"{safe}_{used_safe[safe]}"
    else:
        used_safe[safe] = 1
        file_part = safe
    return f"{base}{sep}{file_part}{suffix}"


def validate_a1(path: Path, sheet_names: list[str]) -> tuple[str, str, str]:
    primary = sheet_names[0]
    base, sep = parse_base_name(path.stem, primary)
    return base, primary, sep


def resolve_a1_in_folder(folder: Path) -> Path:
    candidates = list_excels(folder)
    if not candidates:
        print(f"错误：文件夹内没有找到 Excel 文件:\n  {folder}")
        sys.exit(1)

    valid: list[Path] = []
    skipped: list[tuple[str, str]] = []
    readable: list[Path] = []
    for path in candidates:
        names, err = try_read_sheet_names(path)
        if err:
            skipped.append((path.name, err))
            continue
        readable.append(path)
        assert names is not None
        primary = names[0]
        if stem_matches_primary(path.stem, primary):
            valid.append(path)

    if skipped:
        print("以下文件已跳过（无法作为 Excel 读取）:")
        for name, reason in skipped:
            print(f"  - {name}: {reason}")
        print()

    if not valid and not skipped:
        print(f"错误：文件夹内没有找到 Excel 文件:\n  {folder}")
        sys.exit(1)
    if not valid and skipped:
        print("错误：文件夹内的 Excel 文件均无法读取，请检查文件是否损坏或为旧版 .xls。")
        sys.exit(1)

    if len(valid) == 1:
        print(f"  已选用 A_1: {valid[0].name}")
        return valid[0]

    if len(valid) > 1:
        print("找到多个符合 A_1 命名规则的文件，请选择:")
        for f in valid:
            print(f"    - {f.name}")
        chosen = pick_a1_file(folder)
        if chosen is None:
            print("已取消。")
            sys.exit(0)
        read_sheet_names(chosen)
        return chosen

    print("未自动识别 A_1 / A-1，请手动选择（文件名须以 _或- + 首屏页签名 结尾）:")
    for f in readable:
        print(f"    - {f.name}")
    chosen = pick_a1_file(folder)
    if chosen is None:
        print("已取消。")
        sys.exit(0)
    read_sheet_names(chosen)
    return chosen


def build_manifest(path_a1: Path, folder: Path, sheet_names: list[str], base: str, sep: str) -> dict:
    used_safe: dict[str, int] = {}
    suffix = path_a1.suffix
    targets: list[dict] = []

    for sheet in sheet_names:
        name = target_filename(base, sheet, suffix, sep, used_safe)
        dest = folder / name
        mode = "overwrite" if dest.exists() else "create"
        targets.append(
            {
                "primary": sheet,
                "dest": str(dest.resolve()),
                "mode": mode,
            }
        )

    return {
        "source": str(TEMP_SOURCE.resolve()),
        "source_sheets": sheet_names,
        "base": base,
        "targets": targets,
    }


def write_manifest(payload: dict) -> None:
    MANIFEST_FILE.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def run_com_sync() -> str:
    if not PS1_SCRIPT.exists():
        print(f"错误：找不到 COM 脚本: {PS1_SCRIPT}")
        sys.exit(1)

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

    encoding = "gbk" if sys.platform == "win32" else "utf-8"
    print("  Excel COM：从 A_1 全量同步到各 A_i...")
    proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, encoding=encoding, errors="replace")
    LOG_FILE.write_text(proc.stdout or "", encoding="utf-8")
    for line in (proc.stdout or "").splitlines():
        safe_print(f"  {line}")

    if proc.returncode != 0:
        print()
        print("错误：COM 执行失败，日志:", LOG_FILE)
        sys.exit(1)

    return proc.stdout or ""


def run_openpyxl_sync(manifest: dict) -> str:
    source_path = Path(manifest["source"])
    source_sheets: list[str] = manifest["source_sheets"]
    targets: list[dict] = manifest["targets"]

    print("  openpyxl：复制文件 + 重排页签（Mac 快速模式）...")
    t0 = time.time()
    lines: list[str] = []

    for tgt in targets:
        primary = tgt["primary"]
        dest_path = Path(tgt["dest"])
        mode = tgt["mode"]
        order = desired_order(primary, source_sheets)

        extra_sheet_names: list[str] = []
        old_backup: Path | None = None
        if mode == "overwrite" and dest_path.exists():
            old_wb = load_workbook(dest_path, read_only=True, keep_links=False)
            extra_sheet_names = [n for n in old_wb.sheetnames if n not in source_sheets]
            old_wb.close()
            if extra_sheet_names:
                old_backup = dest_path.with_suffix(dest_path.suffix + ".oldtmp")
                shutil.copy2(dest_path, old_backup)

        shutil.copy2(source_path, dest_path)
        wb = load_workbook(dest_path, keep_links=False)
        reorder_workbook_sheets(wb, order)

        if extra_sheet_names and old_backup is not None:
            old_wb = load_workbook(old_backup, keep_links=False)
            try:
                for name in extra_sheet_names:
                    dest_sheet = wb.create_sheet(name)
                    copy_sheet_between_workbooks(old_wb[name], dest_sheet)
            finally:
                old_wb.close()
            old_backup.unlink(missing_ok=True)

        wb.save(dest_path)
        wb.close()

        tag = "CREATED" if mode == "create" else "SYNCED"
        lines.append(f"{tag},{primary},{dest_path}")
        safe_print(f"  {tag},{primary},{dest_path.name}")

    output_dir = Path(targets[0]["dest"]).parent
    lines.append(f"STATS,{len(targets)},{output_dir}")
    safe_print(f"  STATS,{len(targets)},{output_dir}")
    safe_print(f"  openpyxl 耗时: {time.time() - t0:.1f}s")
    return "\n".join(lines)


def desired_order(primary: str, sheet_names: list[str]) -> list[str]:
    return [primary] + [s for s in sheet_names if s != primary]


def run_export() -> None:
    t0 = time.time()
    print("请选择 A_1 或 A-1 文件...")
    path_a1 = pick_a1_file()
    if path_a1 is None:
        print("已取消。")
        sys.exit(0)

    folder = path_a1.parent
    print(f"  文件: {path_a1.name}")
    print(f"  目录: {folder}")
    sheet_names = read_sheet_names(path_a1)
    base, primary, sep = validate_a1(path_a1, sheet_names)

    shutil.copy2(path_a1, TEMP_SOURCE)
    manifest = build_manifest(path_a1, folder, sheet_names, base, sep)
    write_manifest(manifest)

    print(f"输入 A_1: {path_a1.name}")
    print(f"  base: {base!r}，分隔符: {sep!r}，首屏页签: {primary!r}")
    print(f"  页签 ({len(sheet_names)}): {sheet_names}")
    print(f"  将同步 {len(manifest['targets'])} 个文件")

    if sys.platform == "win32":
        com_output = run_com_sync()
    else:
        com_output = run_openpyxl_sync(manifest)

    synced: list[tuple[str, str, list[str]]] = []
    created: list[tuple[str, str, list[str]]] = []
    stats_count = 0

    for line in com_output.splitlines():
        if line.startswith("SYNCED,"):
            parts = line.split(",", 2)
            if len(parts) >= 3:
                p, dest = parts[1], parts[2]
                synced.append((p, dest, desired_order(p, sheet_names)))
        elif line.startswith("CREATED,"):
            parts = line.split(",", 2)
            if len(parts) >= 3:
                p, dest = parts[1], parts[2]
                created.append((p, dest, desired_order(p, sheet_names)))
        elif line.startswith("STATS,"):
            parts = line.split(",", 2)
            if len(parts) >= 2:
                stats_count = int(parts[1])

    for p in (TEMP_SOURCE, MANIFEST_FILE):
        if p.exists():
            p.unlink()

    print()
    print("=" * 50)
    print("按页签全量同步完成")
    print("=" * 50)
    print(f"  输入: {path_a1.name}")
    print(f"  共处理: {stats_count} 个文件")
    if synced:
        print(f"  覆盖同步: {len(synced)} 个")
        for p, dest, order in synced:
            print(f"    {Path(dest).name}  页签顺序: {order}")
    if created:
        print(f"  新建: {len(created)} 个")
        for p, dest, order in created:
            print(f"    {Path(dest).name}  页签顺序: {order}")
    print(f"  总耗时: {time.time() - t0:.1f} 秒")


if __name__ == "__main__":
    run_export()
