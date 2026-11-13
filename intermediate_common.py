"""桌面中介文件（XXX平台XXX类型报表-时间.xlsx）查找，供 10 / 12 等脚本共用。"""

from __future__ import annotations

import os
import re
from pathlib import Path

INTERMEDIATE_NAME_TEMPLATE = "XXX平台XXX类型报表-时间.xlsx"
INTERMEDIATE_NAME_PATTERN = re.compile(r"^.+平台.+类型报表-.+$")
EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm"}


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


def is_intermediate_name(name: str) -> bool:
    return bool(INTERMEDIATE_NAME_PATTERN.match(Path(name).stem))


def find_intermediate_files(desktop: Path | None = None) -> list[Path]:
    root = desktop or get_desktop_dir()
    if not root.is_dir():
        return []
    results: list[Path] = []
    folders: list[Path] = [root]
    for sub in root.iterdir():
        if sub.is_dir() and not sub.name.startswith("."):
            folders.append(sub)
    for folder in folders:
        for path in folder.iterdir():
            if (
                path.is_file()
                and path.suffix.lower() in EXCEL_SUFFIXES
                and not path.name.startswith("~$")
                and is_intermediate_name(path.name)
            ):
                results.append(path)
    return sorted(set(results), key=lambda p: p.name.lower())
