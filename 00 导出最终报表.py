"""
导出最终版报表（代码库 00）

弹窗多选 Excel 文件，复制到指定导出文件夹（默认建议「00-最终导出/」）。

运行前请关闭 Excel 中已打开的相关文件。
运行：python "00 导出最终报表.py"
"""

from __future__ import annotations

import shutil
import sys
import time
from pathlib import Path

from file_picker_common import get_initial_dir, pick_excel_files, pick_folder

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_EXPORT_DIR = BASE_DIR / "00-最终导出"


def copy_to_export(source: Path, dest: Path) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists():
        dest.unlink()
    shutil.copy2(source, dest)


def run_export() -> None:
    t0 = time.time()
    print("请选择要导出的 Excel 文件（可多选）...")
    sources = pick_excel_files("选择要导出的 Excel 文件（可多选）")
    if not sources:
        print("已取消。")
        sys.exit(0)

    print(f"  已选 {len(sources)} 个文件")
    print("请选择导出目标文件夹...")
    export_dir = pick_folder("选择导出目标文件夹", initial_dir=DEFAULT_EXPORT_DIR if DEFAULT_EXPORT_DIR.is_dir() else get_initial_dir())
    if export_dir is None:
        print("已取消。")
        sys.exit(0)

    print(f"导出目录: {export_dir}")
    print()

    copied = 0
    failed = 0
    seen_names: dict[str, Path] = {}

    for source in sources:
        dest = export_dir / source.name
        if source.name in seen_names:
            print(f"警告：同名文件 {source.name!r} 出现多次，后者将覆盖前者")
            print(f"  先前: {seen_names[source.name]}")
            print(f"  当前: {source}")
        seen_names[source.name] = source

        print(f"复制: {source.name}")
        print(f"  来自: {source.parent}")
        try:
            copy_to_export(source, dest)
            print(f"  → {dest}")
            copied += 1
        except PermissionError:
            print("  失败：文件被占用，请关闭 Excel 后重试")
            failed += 1
        except Exception as exc:
            print(f"  失败：{exc}")
            failed += 1

    print()
    print("=" * 50)
    print("完成")
    print("=" * 50)
    print(f"  复制成功: {copied}")
    print(f"  失败: {failed}")
    print(f"  导出目录: {export_dir}")
    print(f"  总耗时: {time.time() - t0:.1f} 秒")


if __name__ == "__main__":
    run_export()
