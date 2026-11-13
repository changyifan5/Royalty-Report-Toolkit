"""手动登记除错轮次、核对送审版本和汇总报表进度。"""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import tkinter as tk
import webbrowser
from collections import deque
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from tempfile import NamedTemporaryFile, TemporaryDirectory, mkdtemp
from tkinter import messagebox, simpledialog, ttk

from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries

from file_picker_common import pick_excel_file

STATE = Path(__file__).resolve().with_name(".report-workflow.json")
WINDOW_SIZE = STATE.with_name(".report-desk-window-size")
AUTO_BACKUP_DIR = STATE.with_name(".report-desk-backups")
EXPORT, SUBMIT, SPLIT = 0, 3, 6
STEPS = {
    EXPORT: "登记系统汇出",
    SUBMIT: "登记送入系统",
    SPLIT: "登记拆分结果并核对总金额",
}
PROJECT_COLORS = (
    ("#D9EAF7", "#EEF6FB"), ("#D7EEE5", "#EEF8F4"),
    ("#F5E5CC", "#FBF5EB"), ("#E8E0F4", "#F5F0FA"),
    ("#F5DDE6", "#FCF0F4"), ("#D9EEEE", "#EDF8F8"),
    ("#EFEBCC", "#FAF8E9"), ("#E3E5F6", "#F1F2FB"),
)


def restore_task(task: dict) -> dict:
    """兼容旧版除错步骤，保留已有历史和对账记录。"""
    if task["step"] == 4:
        return {**task, "step": EXPORT, "round": task["round"] + 1}
    if task["step"] in (1, 2, 5):
        return {**task, "step": SUBMIT}
    return task


def save_state(tasks: list[dict]) -> None:
    with NamedTemporaryFile(dir=STATE.parent, delete=False) as stream:
        temporary = Path(stream.name)
    try:
        temporary.write_text(json.dumps(tasks, ensure_ascii=False, indent=2), encoding="utf-8")
        temporary.replace(STATE)
    finally:
        temporary.unlink(missing_ok=True)


def report_folder(root: Path, number: str) -> Path:
    endings = ("系统导出",) if number == "04" else ("送进系统", "送入系统")
    matches = [
        child for child in root.iterdir()
        if child.is_dir() and "".join(child.name.split()) in {number + name for name in endings}
    ]
    return matches[0] if len(matches) == 1 else root


def report_root(path: Path) -> Path:
    folder = path.parent
    if "".join(folder.name.split()) in {
        "01原始文件", "02工作文件", "03送进系统", "03送入系统", "04系统导出",
    }:
        return folder.parent
    return folder


def choose_report(title: str, directory: Path, parent: tk.Tk) -> tuple[Path, str] | None:
    path = pick_excel_file(title, initial_dir=directory)
    if path is None:
        return None
    workbook = load_workbook(path, read_only=True)
    try:
        names = [sheet.title for sheet in workbook.worksheets]
    finally:
        workbook.close()
    if len(names) == 1:
        return path, names[0]
    default = names.index("总") + 1 if "总" in names else 1
    number = simpledialog.askinteger(
        "选择明细页签", "\n".join(f"{i}  {name}" for i, name in enumerate(names, 1)),
        initialvalue=default, minvalue=1, maxvalue=len(names), parent=parent,
    )
    return (path, names[number - 1]) if number is not None else None


def read_f3(path: Path, sheet: str) -> dict:
    wb = load_workbook(path, read_only=True, data_only=False, keep_links=False)
    try:
        cell = wb[sheet]["F3"]
        formula = cell.value if cell.data_type == "f" else None
        value = cell.value
    finally:
        wb.close()
    if formula is not None:
        wb = load_workbook(path, read_only=True, data_only=True, keep_links=False)
        try:
            ws = wb[sheet]
            value = ws["F3"].value
            match = re.fullmatch(
                r"=SUM\(\s*(\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+)\s*\)",
                formula, re.IGNORECASE,
            ) if isinstance(formula, str) else None
            if match:
                left, top, right, bottom = range_boundaries(match[1].upper())
                if not (left <= 6 <= right and top <= 3 <= bottom):
                    options = dict(
                        min_row=top, max_row=min(bottom, ws.max_row or bottom),
                        min_col=left, max_col=right,
                    )
                    total = Decimal(0)
                    source_book = load_workbook(path, read_only=True, keep_links=False)
                    try:
                        source_rows = source_book[sheet].iter_rows(**options)
                        missing_formula = False
                        for row, originals in zip(ws.iter_rows(**options), source_rows):
                            for source, original in zip(row, originals):
                                if source.data_type == "e":
                                    raise ValueError(f"{path.name} / {sheet} 的总金额求和范围内有错误值。")
                                if type(source.value) in (int, float):
                                    total += Decimal(str(source.value))
                                elif source.value is None and original.data_type == "f":
                                    missing_formula = True
                        if not missing_formula:
                            value = total
                    finally:
                        source_book.close()
        finally:
            wb.close()
    try:
        amount = Decimal(str(value))
        if not amount.is_finite():
            raise InvalidOperation
    except InvalidOperation as exc:
        raise ValueError(f"{path.name} / {sheet} 的总金额没有可读数值。若是公式，请在 Excel 中重新计算并保存。") from exc
    return {"path": str(path.resolve()), "sheet": sheet, "value": format(amount, "f")}


def f3_line(item: dict) -> str:
    return f'{item["path"]} / {item["sheet"]}   总金额 = {item["value"]}'


def compare_f3(baseline: dict, items: list[dict]) -> tuple[Decimal, Decimal]:
    total = sum((Decimal(item["value"]) for item in items), Decimal(0))
    return total, total - Decimal(baseline["value"])


def latest_report_path(task: dict) -> Path | None:
    for event in reversed(task["history"]):
        if event.get("path"):
            return Path(event["path"])
        if event.get("split_items"):
            return Path(event["split_items"][-1]["path"])
    return Path(task["report_path"]) if task.get("report_path") else None


def numbered_tools(directory: Path, platform: str) -> dict[str, Path]:
    pattern = "运行-*.command" if platform == "darwin" else "[0-9][0-9] *.py"
    result = {}
    for path in sorted(directory.glob(pattern)):
        label = path.stem.removeprefix("运行-")
        if re.match(r"^[0-9]{2}(?![0-9])", label) and not label.startswith("24"):
            result[label] = path
    return result


def with_backup(task: dict, path: str, round_no: int, stage: str) -> dict:
    backups = task.get("backups", [])
    if any(item["path"] == path for item in backups):
        return task
    entry = {
        "path": path, "round": round_no, "stage": stage,
        "at": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    return {**task, "backups": [*backups, entry]}


def split_branches(task: dict) -> list[dict]:
    if "branches" in task:
        return task["branches"]
    for event in reversed(task["history"]):
        if "split_items" in event:
            return event["split_items"]
    return []


def branch_versions(branch: dict) -> list[dict]:
    if "versions" in branch:
        return branch["versions"]
    return [{key: branch[key] for key in ("path", "sheet", "value", "at", "note") if key in branch}]


def activate_version(branch: dict, versions: list[dict], index: int) -> dict:
    version = versions[index]
    return {**branch, **version, "note": version.get("note", ""),
            "versions": versions, "active_version": index, "working_source": None}


def backup_export(source: Path) -> Path:
    AUTO_BACKUP_DIR.mkdir(exist_ok=True)
    folder = Path(mkdtemp(prefix="export-", dir=AUTO_BACKUP_DIR))
    copy = folder / source.name
    try:
        shutil.copy2(source, copy)
    except OSError:
        copy.unlink(missing_ok=True)
        folder.rmdir()
        raise
    return copy


def remove_export_backup(path: Path) -> None:
    folder = path.parent
    if folder.is_symlink() or folder.resolve().parent != AUTO_BACKUP_DIR.resolve() or not folder.name.startswith("export-"):
        raise ValueError("只能删除进度台生成的自动备份。")
    path.unlink(missing_ok=True)
    if not any(folder.iterdir()):
        folder.rmdir()


def registration_details(parent, title: str, detail: str, *, name: str | None = None, automatic_backup=False):
    class RegistrationDialog(simpledialog.Dialog):
        def body(self, master):
            ttk.Label(master, text=detail, wraplength=580, justify="left").grid(
                row=0, column=0, columnspan=3, sticky="w", padx=8, pady=8,
            )
            self.name = tk.StringVar(master=self, value=name or "")
            if name is not None:
                ttk.Label(master, text="子项目名称").grid(row=1, column=0, sticky="w", padx=8)
                ttk.Entry(master, textvariable=self.name, width=35).grid(row=1, column=1, sticky="ew", padx=8)
            self.note = tk.StringVar(master=self)
            self.backup = tk.BooleanVar(master=self)
            ttk.Label(master, text="修改备注").grid(row=2, column=0, sticky="w", padx=8, pady=8)
            entry = ttk.Entry(master, textvariable=self.note, width=48)
            entry.grid(row=2, column=1, sticky="ew", padx=8, pady=8)
            if automatic_backup:
                ttk.Label(master, text="保存登记时自动备份").grid(row=2, column=2, sticky="w", padx=8)
            else:
                ttk.Checkbutton(master, text="标记为备份", variable=self.backup).grid(row=2, column=2, sticky="w", padx=8)
            return entry

        def buttonbox(self):
            buttons = ttk.Frame(self, padding=8)
            ttk.Button(buttons, text="保存登记", command=self.ok, default="active").pack(side="left", padx=4)
            ttk.Button(buttons, text="取消", command=self.cancel).pack(side="left", padx=4)
            self.bind("<Return>", self.ok)
            self.bind("<Escape>", self.cancel)
            buttons.pack()

        def validate(self):
            if name is None or self.name.get().strip():
                return True
            messagebox.showwarning("子项目名称", "请输入子项目名称。", parent=self)
            return False

        def apply(self):
            self.result = self.name.get().strip(), self.note.get().strip(), self.backup.get()

    return RegistrationDialog(parent, title).result


def load_window_size() -> tuple[int, int]:
    try:
        width, height = WINDOW_SIZE.read_text(encoding="utf-8").strip().split("x")
        size = int(width), int(height)
        if min(size) > 0:
            return size
    except (OSError, ValueError):
        pass
    return 1030, 650


def main() -> None:
    saved = json.loads(STATE.read_text(encoding="utf-8")) if STATE.exists() else []
    tasks = [restore_task(task) for task in saved]
    root = tk.Tk()
    root.title("24 报表进度台")
    saved_width, saved_height = load_window_size()
    max_width = max(1, root.winfo_screenwidth() - 48)
    max_height = max(1, root.winfo_screenheight() - 100)
    width = min(max(920, saved_width), max_width)
    height = min(max(500, saved_height), max_height)
    root.geometry(f"{width}x{height}+24+24")
    root.minsize(min(920, max_width), min(500, max_height))
    root.columnconfigure(0, weight=1)
    root.rowconfigure(2, weight=1)
    root.configure(background="#EEF3F7")
    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure("Desk.Treeview", rowheight=32, background="#FFFFFF", foreground="#153147", borderwidth=0)
    style.configure("Desk.Treeview.Heading", padding=(10, 9), background="#E2EBF2", foreground="#153147")
    style.map("Desk.Treeview", background=[("selected", "#17626A")], foreground=[("selected", "#FFFFFF")])
    style.configure("Desk.TButton", padding=(10, 6))
    style.configure("Stage.TButton", padding=(14, 7), background="#17626A", foreground="#FFFFFF")
    style.map("Stage.TButton", background=[("active", "#124F55"), ("disabled", "#CCD8DE")])
    banner = tk.Frame(root, background="#153147", padx=18, pady=12)
    banner.grid(row=0, column=0, sticky="ew")
    tk.Label(banner, text="报表进度台", font=("TkDefaultFont", 20, "bold"), background="#153147",
             foreground="#FFFFFF").pack(side="left")
    overview = tk.StringVar()
    tk.Label(banner, textvariable=overview, background="#153147", foreground="#C6E1E5").pack(side="right")
    setup_bar = tk.Frame(root, background="#EEF3F7")
    setup_bar.grid(row=1, column=0, sticky="ew", padx=18, pady=(10, 0))
    table = tk.Frame(root, background="#EEF3F7")
    table.grid(row=2, column=0, sticky="nsew", padx=18, pady=(14, 8))
    table.columnconfigure(0, weight=1)
    table.rowconfigure(0, weight=1)
    tree = ttk.Treeview(table, columns=("round", "step", "time"), show="tree headings",
                        selectmode="browse", style="Desk.Treeview", height=1)
    scrollbar = ttk.Scrollbar(table, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.grid(row=0, column=1, sticky="ns")
    for index, (project_color, branch_color) in enumerate(PROJECT_COLORS):
        tree.tag_configure(f"project-{index}", background=project_color)
        tree.tag_configure(f"branch-{index}", background=branch_color)
    tree.tag_configure("completed", foreground="#17626A")
    for key, label, width in (
        ("#0", "项目 / 子项目", 310), ("round", "轮次", 85),
        ("step", "下一步", 180), ("time", "上次确认", 170),
    ):
        tree.heading(key, text=label)
        tree.column(key, width=width)
    tree.grid(row=0, column=0, sticky="nsew")
    details = tk.StringVar()
    tk.Label(root, textvariable=details, background="#FFFFFF", foreground="#153147",
             wraplength=850, anchor="w", justify="left", padx=12, pady=9).grid(
                 row=3, column=0, sticky="ew", padx=18,
             )
    bar = tk.Frame(root, background="#EEF3F7")
    bar.grid(row=4, column=0, sticky="ew", padx=18, pady=8)
    stage_bar = tk.Frame(root, background="#E2EBF2", padx=12, pady=8)
    stage_bar.grid(row=5, column=0, sticky="ew", padx=18)
    stage_label = tk.StringVar()
    tk.Label(stage_bar, textvariable=stage_label, background="#E2EBF2", foreground="#153147").pack(side="left")
    tools_bar = tk.Frame(root, background="#EEF3F7")
    tools_bar.grid(row=6, column=0, sticky="ew", padx=18, pady=9)

    packer_queue = deque()
    packer_status = {}
    packer_run = None
    packer_poll_pending = False

    def open_packer_page() -> None:
        task = current()
        if task is None:
            messagebox.showinfo("打开系统页面", "请先选中项目或分支。", parent=root)
            return
        branch = current_branch(task)
        owner = task if branch is None else branch
        report_id = owner.get("packer_id")
        if not re.fullmatch(r"[1-9][0-9]*", str(report_id or "")):
            messagebox.showinfo("打开系统页面", "请先点「设置报表 ID」。", parent=root)
            return
        url = f"https://reports.example.invalid/zh-hans/pa/report/{report_id}/view/"
        if not webbrowser.open(url, new=2):
            raise RuntimeError(f"无法打开浏览器，请手动访问：{url}")

    def run_packer() -> None:
        task = current()
        if task is None:
            messagebox.showinfo("开始比对", "请先选中项目或分支。", parent=root)
            return
        branch = current_branch(task)
        owner = task if branch is None else branch
        report_id = owner.get("packer_id")
        if not report_id:
            messagebox.showinfo("开始比对", "请先点「设置报表 ID」。", parent=root)
            return
        if any(job_id == report_id for job_id, _ in packer_queue) or (
            packer_run is not None and packer_run[0] == report_id
        ):
            messagebox.showinfo("开始比对", "这个报表已在运行或排队。", parent=root)
            return
        if branch is not None:
            source = branch.get("path")
        else:
            source = next(
                (event["path"] for event in reversed(task["history"])
                 if event.get("label") == STEPS[SUBMIT] and event.get("path")),
                None,
            )
        if not source:
            messagebox.showinfo("开始比对", "请先登记送入系统的 Excel。", parent=root)
            return
        path = Path(source)
        if not path.is_file() or path.suffix.lower() != ".xlsx":
            raise ValueError(f"找不到有效的 .xlsx 登记文件：{path}")
        if not messagebox.askyesno(
            "上传并取代报表",
            f"用这份文件取代 Packer 报表 {report_id}？\n如果该报表正在比对，将等当前比对完成后再上传。\n\n{path}",
            default=messagebox.NO, parent=root,
        ):
            return
        packer_queue.append((report_id, path))
        packer_status[report_id] = "排队中"
        if not packer_poll_pending:
            poll_packer()
        else:
            show_details()

    def poll_packer() -> None:
        nonlocal packer_run, packer_poll_pending
        packer_poll_pending = False
        if packer_run is None and packer_queue:
            report_id, path = packer_queue.popleft()
            directory = TemporaryDirectory(prefix="packer-desk-")
            status_path = Path(directory.name) / "status.json"
            output_path = Path(directory.name) / "output.log"
            try:
                with output_path.open("wb") as output:
                    process = subprocess.Popen(
                        [sys.executable, str(STATE.with_name("20 Packer比对汇出.py")),
                         report_id, "--mode", "queue-replace", "--file", str(path),
                         "--yes-download", "--status-file", str(status_path)],
                        cwd=STATE.parent, stdin=subprocess.DEVNULL, stdout=output, stderr=subprocess.STDOUT,
                    )
                packer_run = report_id, process, directory, status_path, output_path
                packer_status[report_id] = "正在启动 20 号"
            except OSError as exc:
                directory.cleanup()
                packer_status[report_id] = f"20 号启动失败：{exc}"
        if packer_run is not None:
            report_id, process, directory, status_path, output_path = packer_run
            code = process.poll()
            record = json.loads(status_path.read_text(encoding="utf-8"))[0] if status_path.exists() else {}
            result = record.get("result", {})
            if code is None:
                packer_status[report_id] = record.get("phase", "正在启动 20 号")
            else:
                if code == 0 and result.get("file"):
                    packer_status[report_id] = f'已下载：{result["file"]}'
                else:
                    lines = output_path.read_text(encoding="utf-8", errors="replace").strip().splitlines()
                    error = result.get("error") or (lines[-1] if lines else f"退出码 {code}")
                    packer_status[report_id] = f"20 号失败：{error}"
                directory.cleanup()
                packer_run = None
        show_details()
        if packer_run is not None or packer_queue:
            packer_poll_pending = True
            root.after(1000, poll_packer)

    def packer_detail(owner: dict) -> str:
        report_id = owner.get("packer_id")
        status = packer_status.get(report_id)
        return f'Packer 报表 ID  {report_id or "未登记"}' + (f'\n比对与汇出  {status}' if status else "")

    def selection() -> tuple[int, int | None] | None:
        selected = tree.selection()
        if not selected:
            return None
        parent, _, branch = selected[0].partition(":")
        return int(parent), int(branch) if branch else None

    def current() -> dict | None:
        selected = selection()
        return tasks[selected[0]] if selected is not None else None

    def current_branch(task: dict) -> dict | None:
        selected = selection()
        if selected is not None and selected[1] is not None:
            return split_branches(task)[selected[1]]
        return None

    def show_details(_event=None) -> None:
        task = current()
        if not task:
            details.set("添加报表后，选中一行查看文件和下一步。")
            stage_label.set("请选择报表")
            stage_button.state(["disabled"])
            next_button.configure(text="登记下一步", state="disabled")
            split_export_button.pack_forget()
            return
        branch = current_branch(task)
        if task["step"] == SPLIT and not task["done"]:
            split_export_button.pack(side="left", padx=(0, 5))
        else:
            split_export_button.pack_forget()
        if branch is not None:
            details.set(
                f'{task["name"]} / {branch.get("name", branch["sheet"])}\n'
                f'{f3_line(branch)}\n修改备注  {branch.get("note") or "—"}'
                f'  |  备份标记 {len(branch.get("backups", []))} 份\n'
                f"{packer_detail(branch)}"
            )
            stage_label.set("当前选择  子项目")
            stage_button.state(["disabled"])
            next_button.configure(text="登记送入系统", state="disabled" if task["done"] else "normal")
            return
        stage = "已完成" if task["done"] else ("拆分与对账" if task["step"] == SPLIT else "除错")
        stage_label.set(f"当前阶段  {stage}")
        stage_button.state(["disabled" if task["done"] else "!disabled"])
        next_button.configure(
            text="登记拆分子项目" if task["step"] == SPLIT else STEPS[task["step"]],
            state="disabled" if task["done"] else "normal",
        )
        event = task["history"][-1] if task["history"] else {}
        path = latest_report_path(task)
        details.set(
            f'{task["root"]}\n'
            f'登记文件  {task.get("report_path", "（旧记录）")}\n'
            f'最近确认  {event.get("label", "待登记")}\n'
            f'最近文件  {path or "待登记"}\n'
            f'原始总金额  {task.get("baseline", {}).get("value", "待登记")}  |  备份标记 {len(task.get("backups", []))} 份\n'
            f"{packer_detail(task)}"
        )

    def refresh(index: int | None = None, branch: int | None = None) -> None:
        today = datetime.now().strftime("%Y-%m-%d")
        overview.set(f'今天 {today}   ·   {len(tasks)} 份报表   ·   {sum(task["done"] for task in tasks)} 份完成')

        def display_time(value: str) -> str:
            return f"今天 {value[11:]}" if value.startswith(today + " ") else value

        for item in tree.get_children():
            tree.delete(item)
        for i, task in enumerate(tasks):
            last = task["history"][-1] if task["history"] else {}
            step = "已完成" if task["done"] else STEPS[task["step"]]
            color = i % len(PROJECT_COLORS)
            tags = (f"project-{color}", "completed") if task["done"] else (f"project-{color}",)
            tree.insert("", "end", iid=str(i), text=task["name"], open=True, tags=tags,
                        values=(task["round"], step, display_time(last.get("at", ""))))
            for j, child in enumerate(split_branches(task)):
                tree.insert(str(i), "end", iid=f"{i}:{j}", text=child.get("name", child["sheet"]),
                            tags=(f"branch-{color}",),
                            values=("", "已对账" if task["done"] else "待对账", display_time(child.get("at", ""))))
        if index is not None:
            tree.selection_set(str(index) if branch is None else f"{index}:{branch}")
        show_details()

    def add() -> None:
        path = pick_excel_file("选择要添加的报表 Excel")
        if path is None:
            return
        path = path.resolve()
        name = simpledialog.askstring("报表名称", "输入你想用的报表名称", initialvalue=path.stem, parent=root)
        if not name or not name.strip():
            return
        round_no = simpledialog.askinteger("当前轮次", "这份报表目前是第几轮？", minvalue=1, initialvalue=1, parent=root)
        if round_no is None:
            return
        tasks.append({
            "name": name.strip(), "report_path": str(path), "root": str(report_root(path)),
            "round": round_no, "step": 0, "done": False, "history": [],
        })
        try:
            save_state(tasks)
        except OSError:
            tasks.pop()
            raise
        refresh(len(tasks) - 1)

    def update_task(index: int, updated: dict, branch: int | None = None) -> None:
        replacement = [*tasks]
        replacement[index] = updated
        save_state(replacement)
        tasks[index] = updated
        refresh(index, branch)

    def remove_task() -> None:
        task = current()
        if task is None:
            messagebox.showinfo("删除登记", "请先选中要删除的报表。", parent=root)
            return
        index, branch_index = selection()
        if branch_index is not None:
            if task["done"]:
                messagebox.showinfo("删除登记", "已对账的子项目保留在完成记录中。", parent=root)
                return
            branches = split_branches(task)
            branch = branches[branch_index]
            if messagebox.askyesno("删除子项目", f'移除「{branch.get("name", branch["sheet"])}」的登记和备份标记？',
                                   default=messagebox.NO, parent=root):
                update_task(index, {**task, "branches": branches[:branch_index] + branches[branch_index + 1:]})
            return
        if not messagebox.askyesno(
            "删除登记",
            f'确定删除「{task["name"]}」的登记吗？\n\n'
            "该报表的进度、对账和备份标记将一并移除。此操作只影响进度台记录。",
            default=messagebox.NO,
            parent=root,
        ):
            return
        remaining = tasks[:index] + tasks[index + 1:]
        save_state(remaining)
        tasks[:] = remaining
        refresh(min(index, len(tasks) - 1) if tasks else None)

    def set_packer_id() -> None:
        task = current()
        if task is None:
            messagebox.showinfo("报表 ID", "请先选中项目或分支。", parent=root)
            return
        index, branch_index = selection()
        branch = current_branch(task)
        owner = task if branch is None else branch
        value = simpledialog.askstring(
            "Packer 报表 ID", "请输入数字报表 ID，例如 1003。留空可以清除登记。",
            initialvalue=owner.get("packer_id", ""), parent=root,
        )
        if value is None:
            return
        value = value.strip()
        if value and not re.fullmatch(r"[1-9][0-9]*", value):
            raise ValueError("报表 ID 必须是正整数。")
        updated = {**owner, "packer_id": value}
        if branch_index is None:
            update_task(index, updated)
        else:
            branches = list(split_branches(task))
            branches[branch_index] = updated
            update_task(index, {**task, "branches": branches}, branch_index)

    def set_baseline() -> None:
        task = current()
        if not task or task["done"] or current_branch(task) is not None:
            return
        selected = choose_report("选择最初未处理的报表与基准页签", Path(task["root"]), root)
        if selected is None:
            return
        baseline = read_f3(*selected)
        old = task.get("baseline")
        prompt = f'当前基准\n{f3_line(old)}\n\n' if old else ""
        if messagebox.askyesno("登记原始总金额", prompt + f'使用以下原始基准？\n{f3_line(baseline)}', parent=root):
            update_task(selection()[0], {**task, "baseline": baseline})

    def mark_backup() -> None:
        task = current()
        if not task:
            return
        index, branch_index = selection()
        branch = current_branch(task)
        owner = branch if branch is not None else task
        path = pick_excel_file(
            "选择要标记为备份的文件",
            initial_dir=Path(branch["path"]).parent if branch else Path(task["root"]),
        )
        if path is None:
            return
        path_text = str(path.resolve())
        stage = "拆分子项目" if branch else ("已完成" if task["done"] else STEPS[task["step"]])
        updated = with_backup(owner, path_text, task["round"], stage)
        if updated is owner:
            messagebox.showinfo("备份标记", "这份文件已在当前项目标记为备份。", parent=root)
            return
        name = branch.get("name", branch["sheet"]) if branch else task["name"]
        if not messagebox.askyesno("标记备份", f"标记到「{name}」？\n{path}", parent=root):
            return
        if branch_index is None:
            update_task(index, updated)
        else:
            branches = list(split_branches(task))
            branches[branch_index] = updated
            update_task(index, {**task, "branches": branches}, branch_index)

    def advance() -> None:
        task = current()
        if not task or task["done"] or current_branch(task) is not None:
            return
        index, _ = selection()
        stage = task["step"]
        directory = Path(task["root"])
        label = STEPS[stage]
        event = {"label": label, "at": datetime.now().strftime("%Y-%m-%d %H:%M"), "round": task["round"]}
        updated = dict(task)
        if stage == EXPORT:
            register_export()
            return
        if stage == SUBMIT:
            path = pick_excel_file(label, initial_dir=report_folder(directory, "03"))
            if path is None:
                return
            event["path"] = str(path.resolve())
            prompt = f'{task["name"]} · 第 {task["round"]} 轮\n{path}'
            updated.update(round=task["round"] + 1, step=EXPORT, working_source=None)
        else:
            baseline = task.get("baseline")
            if baseline is None:
                raise ValueError("请先点「登记原始总金额」，选择最开始没处理的报表。")
            branches = split_branches(task)
            if not branches:
                raise ValueError("请先登记要参与对账的拆分子项目。")
            items = []
            for branch in branches:
                item = read_f3(Path(branch["path"]), branch["sheet"])
                if (item["path"], item["sheet"]) == (baseline["path"], baseline["sheet"]):
                    raise ValueError("原始总金额的页签已登记为拆分子项目，请移除后再对账。")
                items.append({**branch, **item})
            updated["branches"] = items
            total, difference = compare_f3(baseline, items)
            prompt = "\n".join([
                f"原始总金额  {baseline['value']}",
                *(f'{item.get("name", item["sheet"])}  {f3_line(item)}' for item in items),
                f"拆分合计  {total:f}   差额  {difference:f}",
            ])
            if difference:
                messagebox.showwarning("总金额有差额", prompt, parent=root)
                return
            event.update(split_items=items, total=format(total, "f"), difference=format(difference, "f"))
            updated["done"] = True
        if stage == SUBMIT:
            registration = registration_details(root, label, prompt)
            if registration is None:
                return
            _, event["note"], backup = registration
            if backup:
                updated = with_backup(updated, event["path"], task["round"], "除错 · 送入系统")
        elif not messagebox.askyesno(label, prompt + "\n\n是否确认登记？", parent=root):
            return
        updated["history"] = [*task["history"], event]
        update_task(index, updated)

    def register_export() -> None:
        task = current()
        if not task or task["done"] or task["step"] not in (EXPORT, SPLIT):
            return
        splitting = task["step"] == SPLIT
        index, branch_index = selection()
        branch = current_branch(task)
        owner = branch if branch is not None else task
        path = pick_excel_file("选择系统汇出文件", initial_dir=report_folder(Path(task["root"]), "04"))
        if path is None:
            return
        name = branch.get("name", branch["sheet"]) if branch else task["name"]
        registration = registration_details(root, STEPS[EXPORT], f"{name}\n{path}", automatic_backup=True)
        if registration is None:
            return
        _, note, _ = registration
        records = [*owner.get("history", []), *owner.get("source_records", [])]
        previous = next(
            (Path(record["backup_copy"]) for record in reversed(records)
             if record.get("backup_copy")),
            None,
        )
        copy = backup_export(path)
        record = {
            "path": str(path.resolve()), "note": note, "label": STEPS[EXPORT], "round": task["round"],
            "at": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "backup_copy": str(copy),
        }
        updated = {**owner, "working_source": record["path"]}
        if splitting:
            updated["source_records"] = [*owner.get("source_records", []), record]
        else:
            updated.update(step=SUBMIT, history=[*task["history"], record])
        updated = with_backup(updated, str(copy), task["round"], "系统汇出 · 自动备份")

        def save_owner(new_owner):
            if branch_index is None:
                update_task(index, new_owner)
                return
            branches = list(split_branches(task))
            branches[branch_index] = new_owner
            update_task(index, {**task, "branches": branches}, branch_index)

        try:
            save_owner(updated)
        except OSError:
            remove_export_backup(copy)
            raise
        if previous and previous.is_file() and messagebox.askyesno(
            "上一份自动备份",
            f"新备份已保存。要删除「{name}」的上一份自动备份吗？\n\n{previous}",
            default=messagebox.NO, parent=root,
        ):
            remove_export_backup(previous)
            remaining = [item for item in updated["backups"] if item["path"] != str(previous)]
            save_owner({**updated, "backups": remaining})

    def register_branch_version() -> None:
        task = current()
        if not task or task["done"] or task["step"] != SPLIT:
            return
        index, branch_index = selection()
        branch = current_branch(task)
        directory = report_folder(Path(task["root"]), "03")
        selected = choose_report("选择送入系统的文件和页签", directory, root)
        if selected is None:
            return
        item = read_f3(*selected)
        branches = list(split_branches(task))
        if any(
            j != branch_index and (other["path"], other["sheet"]) == (item["path"], item["sheet"])
            for j, other in enumerate(branches)
        ):
            raise ValueError("这个文件的该页签已登记在另一个子项目。")
        title = "登记拆分子项目" if branch is None else "登记送入系统"
        detail = f'{task["name"]}\n{f3_line(item)}'
        registration = registration_details(root, title, detail, name=item["sheet"] if branch is None else None)
        if registration is None:
            return
        name, note, backup = registration
        item.update(note=note, at=datetime.now().strftime("%Y-%m-%d %H:%M"))
        if branch is None:
            branch = {"name": name}
            versions = [item]
        else:
            versions = [*branch_versions(branch), item]
        updated = activate_version(branch, versions, len(versions) - 1)
        if backup:
            updated = with_backup(updated, item["path"], task["round"], "拆分 · 送入系统")
        if branch_index is None:
            branch_index = len(branches)
            branches.append(updated)
        else:
            branches[branch_index] = updated
        update_task(index, {**task, "branches": branches}, branch_index)

    def next_action() -> None:
        task = current()
        if task and task["step"] == SPLIT:
            register_branch_version()
        elif task:
            advance()

    def finish_stage() -> None:
        task = current()
        if not task or task["done"] or current_branch(task) is not None:
            return
        if task["step"] == SPLIT:
            advance()
            return
        if not messagebox.askyesno(
            "该阶段完成",
            f'确认「{task["name"]}」已经彻底除错完成，可以进入拆分和金额对账？',
            default=messagebox.NO,
            parent=root,
        ):
            return
        event = {
            "label": "确认除错完成", "at": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "round": task["round"],
        }
        update_task(selection()[0], {
            **task, "step": SPLIT, "history": [*task["history"], event],
        })

    def open_report(path: Path | None) -> None:
        if path is None or not path.is_file():
            raise FileNotFoundError(f"找不到登记文件：{path or '尚未登记'}")
        if sys.platform == "win32":
            os.startfile(path)
        else:
            command = "open" if sys.platform == "darwin" else "xdg-open"
            subprocess.run([command, str(path)], check=True)

    def open_latest() -> None:
        task = current()
        if task:
            branch = current_branch(task)
            owner = branch if branch is not None else task
            working_source = owner.get("working_source") if not task["done"] else None
            if working_source:
                open_report(Path(working_source))
            else:
                open_report(Path(branch["path"]) if branch else latest_report_path(task))

    def show_versions() -> None:
        task = current()
        if not task:
            messagebox.showinfo("版本记录", "请先选中项目或子项目。", parent=root)
            return
        index, branch_index = selection()
        branch = current_branch(task)
        owner = branch if branch is not None else task
        sources = [{**record, "label": STEPS[EXPORT]} for record in owner.get("source_records", [])]
        versions = branch_versions(branch) if branch else [
            event for event in task["history"]
            if event.get("path") and event["label"] in (STEPS[EXPORT], STEPS[SUBMIT], "登记送审版")
        ]
        records = [*sources, *versions]
        if not records:
            messagebox.showinfo("版本记录", "当前项目还没有登记文件。", parent=root)
            return
        active = branch.get("active_version", len(versions) - 1) if branch else len(versions) - 1
        window = tk.Toplevel(root)
        window.title("版本记录")
        window.transient(root)
        window.grab_set()
        window.geometry("850x380")
        listing = tk.Listbox(window, exportselection=False)
        listing.pack(fill="both", expand=True, padx=12, pady=8)
        detail = tk.StringVar(master=window)
        ttk.Label(window, textvariable=detail, wraplength=810, justify="left").pack(fill="x", padx=12)
        for j, version in enumerate(records):
            mark = "当前 · " if versions and j == len(sources) + active else ""
            listing.insert("end", f'{mark}{version.get("at", "")}  {Path(version["path"]).name}  {version.get("sheet", version.get("label", ""))}')

        def show_version(_event=None):
            chosen = listing.curselection()
            if chosen:
                version = records[chosen[0]]
                copy = version.get("backup_copy")
                text = f'{version["path"]}\n修改备注  {version.get("note") or "—"}'
                available = copy is not None and Path(copy).is_file()
                if copy:
                    text += f'\n自动备份  {copy if available else "已删除或移走"}'
                detail.set(text)
                backup_button.state(["!disabled" if available else "disabled"])

        def open_version():
            chosen = listing.curselection()
            if chosen:
                open_report(Path(records[chosen[0]]["path"]))

        def open_version_backup():
            chosen = listing.curselection()
            if chosen:
                copy = records[chosen[0]].get("backup_copy")
                if copy:
                    open_report(Path(copy))

        def switch_version():
            chosen = listing.curselection()
            if not chosen:
                return
            new_index = chosen[0] - len(sources)
            if new_index < 0:
                messagebox.showinfo("版本记录", "系统导出可以打开查看。请选择一份送入系统的版本切回。", parent=window)
                return
            if new_index == active:
                return
            current_task = tasks[index]
            branches = list(split_branches(current_task))
            current_branch_record = branches[branch_index]
            if current_task["done"] or branch_versions(current_branch_record) != versions:
                raise ValueError("项目记录已更新，请重新打开版本记录。")
            version = versions[new_index]
            item = read_f3(Path(version["path"]), version["sheet"])
            if any(
                j != branch_index and (other["path"], other["sheet"]) == (item["path"], item["sheet"])
                for j, other in enumerate(branches)
            ):
                raise ValueError("这个页签已经是另一个子项目的当前版本。")
            if not messagebox.askyesno("切回选中版本", f'将此版本设为子项目的当前版本？\n{f3_line(item)}',
                                       default=messagebox.NO, parent=window):
                return
            updated_versions = list(versions)
            updated_versions[new_index] = {**version, **item}
            branches[branch_index] = activate_version(current_branch_record, updated_versions, new_index)
            update_task(index, {**current_task, "branches": branches}, branch_index)
            window.destroy()

        buttons = ttk.Frame(window, padding=8)
        buttons.pack(fill="x")
        ttk.Button(buttons, text="打开选中版本", command=lambda: guarded(open_version)).pack(side="left")
        backup_button = ttk.Button(buttons, text="打开自动备份", command=lambda: guarded(open_version_backup))
        backup_button.pack(side="left", padx=8)
        if branch is not None and not task["done"]:
            ttk.Button(buttons, text="切回选中版本", command=lambda: guarded(switch_version)).pack(side="left", padx=8)
        listing.bind("<<ListboxSelect>>", show_version)
        selected_index = len(sources) + active if versions else len(sources) - 1
        listing.selection_set(selected_index)
        listing.see(selected_index)
        show_version()

    def show_backups() -> None:
        task = current()
        if not task:
            messagebox.showinfo("查看备份", "请先选中一份报表。", parent=root)
            return
        branch = current_branch(task)
        owner = branch if branch is not None else task
        name = branch.get("name", branch["sheet"]) if branch else task["name"]
        backups = list(reversed(owner.get("backups", [])))
        if not backups:
            messagebox.showinfo("查看备份", f'「{name}」还没有备份记录。', parent=root)
            return
        window = tk.Toplevel(root)
        window.title(f"备份 · {name}")
        window.geometry("860x350")
        listing = ttk.Treeview(
            window, columns=("file", "round", "stage", "time"),
            show="headings", selectmode="browse", style="Desk.Treeview",
        )
        for key, label, width in (
            ("file", "文件", 350), ("round", "轮次", 65),
            ("stage", "阶段", 175), ("time", "标记时间", 145),
        ):
            listing.heading(key, text=label)
            listing.column(key, width=width)
        for index, entry in enumerate(backups):
            listing.insert("", "end", iid=str(index),
                           values=(Path(entry["path"]).name, entry["round"], entry["stage"], entry["at"]))
        listing.pack(fill="both", expand=True, padx=12, pady=8)
        selected_path = tk.StringVar()
        ttk.Label(window, textvariable=selected_path, wraplength=820).pack(fill="x", padx=12)

        def selection_changed(_event=None) -> None:
            selected = listing.selection()
            selected_path.set(backups[int(selected[0])]["path"] if selected else "")

        def open_selected() -> None:
            if selected_path.get():
                open_report(Path(selected_path.get()))

        listing.bind("<<TreeviewSelect>>", selection_changed)
        listing.bind("<Double-Button-1>", lambda _event: guarded(open_selected))
        listing.selection_set("0")
        selection_changed()
        ttk.Button(window, text="打开备份文件", style="Desk.TButton",
                   command=lambda: guarded(open_selected)).pack(anchor="e", padx=12, pady=8)

    def guarded(action) -> None:
        try:
            action()
        except Exception as exc:
            messagebox.showerror("报表进度台", str(exc), parent=root)

    tools = numbered_tools(STATE.parent, sys.platform)
    tool_picker = ttk.Combobox(tools_bar, values=list(tools), state="readonly", width=34)
    if tools:
        tool_picker.current(0)

    def launch_tool() -> None:
        path = tools.get(tool_picker.get())
        if path is None:
            return
        if sys.platform == "darwin":
            subprocess.run(["open", "-a", "Terminal", str(path)], check=True)
        elif sys.platform == "win32":
            subprocess.Popen([sys.executable, str(path)], cwd=path.parent, creationflags=subprocess.CREATE_NEW_CONSOLE)
        else:
            subprocess.Popen([sys.executable, str(path)], cwd=path.parent)

    def section_label(parent, title: str, *, side: str = "left") -> None:
        tk.Label(parent, text=title, background="#EEF3F7", foreground="#153147",
                 font=("TkDefaultFont", 10, "bold")).pack(side=side, padx=(0, 10))

    def action_button(parent, label, action, *, side="left", style_name="Desk.TButton"):
        button = ttk.Button(parent, text=label, style=style_name, command=lambda: guarded(action))
        button.pack(side=side, padx=(0, 5))
        return button

    tree.bind("<<TreeviewSelect>>", show_details)
    section_label(setup_bar, "报表设置")
    action_button(setup_bar, "添加报表", add)
    action_button(setup_bar, "登记原始总金额", set_baseline)
    action_button(setup_bar, "设置报表 ID", set_packer_id)
    action_button(setup_bar, "删除登记", remove_task, side="right")

    section_label(bar, "每轮操作")
    next_button = action_button(bar, "登记下一步", next_action, style_name="Stage.TButton")
    action_button(bar, "打开登记文件", open_latest)
    action_button(bar, "版本记录", show_versions)
    action_button(bar, "开始比对", run_packer)
    action_button(bar, "打开系统页面", open_packer_page)
    split_export_button = ttk.Button(bar, text=STEPS[EXPORT], style="Desk.TButton",
                                    command=lambda: guarded(register_export))
    stage_button = ttk.Button(stage_bar, text="该阶段完成", style="Stage.TButton",
                              command=lambda: guarded(finish_stage))
    stage_button.pack(side="right")

    archive = tk.Frame(tools_bar, background="#EEF3F7")
    archive.pack(side="left", padx=(0, 18))
    section_label(archive, "备份")
    action_button(archive, "标记备份", mark_backup)
    action_button(archive, "查看备份", show_backups)
    section_label(tools_bar, "运行其他工具")
    tool_picker.pack(side="left")
    ttk.Button(tools_bar, text="启动", style="Desk.TButton",
               command=lambda: guarded(launch_tool)).pack(side="left", padx=6)
    refresh()

    def close_window() -> None:
        try:
            WINDOW_SIZE.write_text(f"{root.winfo_width()}x{root.winfo_height()}", encoding="utf-8")
        except OSError as exc:
            print(f"窗口大小保存失败：{exc}", file=sys.stderr)
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", close_window)
    root.mainloop()


if __name__ == "__main__":
    main()
