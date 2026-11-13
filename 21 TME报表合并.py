"""Windows 使用 Excel COM；Mac 使用 openpyxl。

按 TME 文件名识别报表和平台，把每个文件的首张工作表汇入独立页签。
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path
from tempfile import TemporaryDirectory

from file_picker_common import get_initial_dir, pick_folder

PLATFORMS = (("qq", "QQ音乐"), ("kg", "酷狗音乐"), ("kw", "酷我音乐"))
EXCEL_SUFFIXES = {".xlsx", ".xlsm"}


def match_reports(folder: Path, album_prefix: str) -> list[tuple[str, Path]]:
    report_types = (
        ("mv", "MV明细"),
        ("song", "明细表-音乐平台"),
        ("single", "明细表-单曲订购"),
        ("k", "明细表-K歌及直播"),
        ("aiting", "明细表-爱听"),
        (album_prefix, "明细表-数字专辑"),
    )
    candidates = sorted(folder.iterdir(), key=lambda path: path.name.casefold())
    reports = [
        (path.stem.casefold(), path)
        for path in candidates
        if path.is_file()
        and not path.name.startswith((".", "~$"))
        and path.suffix.lower() in EXCEL_SUFFIXES
        and not path.stem.endswith("_TME合并")
    ]
    matches = []
    for prefix, label in report_types:
        platforms = (("", ""),) if prefix in ("k", "aiting") else PLATFORMS
        for suffix, platform in platforms:
            title = f"{label}（{platform}）" if platform else label
            files = [
                path for stem, path in reports
                if stem.startswith(prefix)
                and (stem.endswith(suffix) or f"_{suffix}_" in stem)
            ]
            if len(files) > 1:
                names = "、".join(path.name for path in files)
                raise ValueError(f"多个文件对应「{title}」：{names}")
            if files:
                matches.append((title, files[0]))
    return matches


def run_openpyxl(matches: list[tuple[str, Path]], output: Path) -> None:
    from openpyxl import Workbook, load_workbook
    from openpyxl.cell import WriteOnlyCell

    with output.open("xb") as stream:
        workbook = Workbook(write_only=True)
        try:
            for title, path in matches:
                source = load_workbook(path, read_only=True, data_only=True, keep_links=False)
                try:
                    target = workbook.create_sheet(title)
                    for row in source.worksheets[0].iter_rows():
                        cells = []
                        for cell in row:
                            if cell.value is None:
                                cells.append(None)
                                continue
                            target_cell = WriteOnlyCell(target, value=cell.value)
                            target_cell.data_type = cell.data_type
                            target_cell.number_format = cell.number_format
                            cells.append(target_cell)
                        target.append(cells)
                finally:
                    source.close()
            workbook.save(stream)
        finally:
            workbook.close()


def run_com(matches: list[tuple[str, Path]], output: Path, temp_folder: Path) -> None:
    manifest = temp_folder / "manifest.json"
    manifest.write_text(json.dumps({
        "reports": [{"title": title, "path": str(path)} for title, path in matches],
        "output": str(output),
    }), encoding="utf-8")
    subprocess.run([
        "powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-File", str(Path(__file__).with_name("21_tme_merge_com.ps1")), str(manifest),
    ], check=True)


def pick_source_folders() -> list[Path]:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("TME 批量合并")
    root.geometry("620x410")
    root.columnconfigure(0, weight=1)
    root.rowconfigure(2, weight=1)
    current = get_initial_dir()
    current_label = tk.StringVar()
    children = []
    selected = []

    ttk.Label(root, textvariable=current_label).grid(row=0, column=0, sticky="w", padx=12, pady=8)
    ttk.Label(root, text="点击子文件夹即可选中，再点一次取消。可同时选多个。").grid(
        row=1, column=0, sticky="w", padx=12,
    )
    listing = tk.Listbox(root, selectmode="multiple", exportselection=False)
    listing.grid(row=2, column=0, sticky="nsew", padx=(12, 0), pady=8)
    scrollbar = ttk.Scrollbar(root, command=listing.yview)
    scrollbar.grid(row=2, column=1, sticky="ns", padx=(0, 12), pady=8)
    listing.configure(yscrollcommand=scrollbar.set)

    def show_directory(folder: Path) -> None:
        nonlocal current, children
        try:
            children = sorted(
                (path for path in folder.iterdir() if path.is_dir() and not path.name.startswith(".")),
                key=lambda path: path.name.casefold(),
            )
        except OSError as exc:
            messagebox.showerror("无法打开文件夹", str(exc), parent=root)
            return
        current = folder
        current_label.set(str(folder))
        listing.delete(0, tk.END)
        for child in children:
            listing.insert(tk.END, child.name)

    def change_directory() -> None:
        folder = filedialog.askdirectory(parent=root, initialdir=str(current), mustexist=True)
        if folder:
            show_directory(Path(folder))

    def confirm_selection(use_current: bool = False) -> None:
        folders = [current] if use_current else [children[i] for i in listing.curselection()]
        if folders:
            selected.extend(folders)
            root.destroy()
        else:
            messagebox.showinfo("选择文件夹", "请先选中要处理的子文件夹。", parent=root)

    buttons = ttk.Frame(root)
    buttons.grid(row=3, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 12))
    ttk.Button(buttons, text="更换上级目录", command=change_directory).pack(side="left")
    ttk.Button(buttons, text="处理当前目录", command=lambda: confirm_selection(True)).pack(side="left")
    ttk.Button(buttons, text="处理选中的文件夹", command=confirm_selection).pack(side="right")
    show_directory(current)
    root.mainloop()
    return selected


def merge_folder(folder: Path, album_prefix: str, output_dir: Path) -> None:
    matches = match_reports(folder, album_prefix)
    if not matches:
        raise ValueError("未找到符合命名规则的 .xlsx 或 .xlsm 文件")
    output = output_dir / f"{folder.name}_TME合并.xlsx"
    if output.exists():
        raise FileExistsError(f"输出文件已存在，请先移走或重命名：{output}")
    for title, path in matches:
        print(f"{title} ← {path.name}", flush=True)
    with TemporaryDirectory(prefix=".tme-", dir=output_dir) as temp_dir:
        temp_folder = Path(temp_dir)
        temp_output = temp_folder / output.name
        if sys.platform == "win32":
            run_com(matches, temp_output, temp_folder)
        else:
            run_openpyxl(matches, temp_output)
        # 独占创建目标，避免两个进程同时运行时覆盖对方的结果。
        output.touch(exist_ok=False)
        try:
            temp_output.replace(output)
        except OSError:
            output.unlink()
            raise
    print(f"已生成 {len(matches)} 个页签：{output}", flush=True)


def main() -> int:
    parser = argparse.ArgumentParser(description="按报表和平台顺序合并 TME Excel 文件")
    parser.add_argument("folders", nargs="*", type=Path, help="一个或多个 TME 报表文件夹")
    parser.add_argument("--output-dir", type=Path, help="统一输出目录")
    parser.add_argument("--album-prefix", default="album", help="数字专辑文件名前缀")
    args = parser.parse_args()
    album_prefix = args.album_prefix.strip().casefold()
    if not album_prefix or any(
        album_prefix.startswith(prefix) or prefix.startswith(album_prefix)
        for prefix in ("mv", "song", "single", "k", "aiting")
    ):
        parser.error("数字专辑前缀必须非空，且与其他报表前缀可区分")

    try:
        folders = list(dict.fromkeys(
            folder.expanduser().resolve() for folder in (args.folders or pick_source_folders())
        ))
        if not folders:
            return 0
        output_dir = args.output_dir
        if output_dir is None and not args.folders:
            output_dir = pick_folder("选择汇总文件的输出目录", initial_dir=folders[0].parent)
            if output_dir is None:
                return 0
        if output_dir is not None:
            output_dir = output_dir.expanduser().resolve()
            if not output_dir.is_dir():
                raise NotADirectoryError(f"输出目录不存在：{output_dir}")
    except Exception as exc:
        print(f"无法开始合并：{exc}", file=sys.stderr)
        return 1

    failed = 0
    for index, folder in enumerate(folders, start=1):
        print(f"\n[{index}/{len(folders)}] {folder}", flush=True)
        try:
            merge_folder(folder, album_prefix, output_dir or folder)
        except Exception as exc:
            failed += 1
            print(f"合并失败：{exc}", file=sys.stderr, flush=True)
    print(f"\n处理完成：成功 {len(folders) - failed} 个，失败 {failed} 个。")
    return int(failed > 0)


if __name__ == "__main__":
    sys.exit(main())
