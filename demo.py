"""Independent interview demo. All examples are synthetic."""

from dataclasses import dataclass, field
from decimal import Decimal


@dataclass(frozen=True)
class Version:
    name: str
    amount: Decimal
    note: str = ""


@dataclass
class Branch:
    name: str
    versions: list[Version] = field(default_factory=list)
    current: int = 0

    def register(self, version: Version) -> None:
        self.versions.append(version)
        self.current = len(self.versions) - 1

    def restore(self, index: int) -> None:
        if not 0 <= index < len(self.versions):
            raise ValueError("Unknown version")
        self.current = index

    @property
    def amount(self) -> Decimal:
        return self.versions[self.current].amount


@dataclass
class Project:
    name: str
    original: Decimal
    stage: str = "除错"
    rounds: list[tuple[str, str]] = field(default_factory=list)
    pending: str | None = None
    branches: dict[str, Branch] = field(default_factory=dict)
    backups: dict[str, Version] = field(default_factory=dict)

    def register_export(self, file: str) -> None:
        if self.stage != "除错" or self.pending is not None:
            raise ValueError("Cannot register an export in this state")
        self.pending = file

    def register_submission(self, file: str) -> None:
        if self.stage != "除错" or self.pending is None:
            raise ValueError("Register the export first")
        self.rounds.append((self.pending, file))
        self.pending = None

    def finish_debugging(self) -> None:
        if self.stage != "除错":
            raise ValueError("Already past debugging")
        self.stage = "拆分与对账"

    def register_branch(self, name: str, version: Version) -> None:
        if self.stage != "拆分与对账":
            raise ValueError("Split the project after debugging")
        self.branches.setdefault(name, Branch(name)).register(version)

    def reconcile(self) -> Decimal:
        if not self.branches:
            raise ValueError("Register at least one branch")
        return sum((branch.amount for branch in self.branches.values()), Decimal(0)) - self.original


def demo() -> None:
    project = Project("虚构平台 · 示例项目", Decimal("1000.00"))
    for number in (1, 2):
        project.register_export(f"demo-export-{number}.xlsx")
        project.register_submission(f"demo-submit-{number}.xlsx")
    print("完成除错轮次", len(project.rounds))
    project.finish_debugging()
    project.register_branch("分支 A", Version("a-v1.xlsx", Decimal("600.00"), "首次拆分"))
    project.register_branch("分支 B", Version("b-v1.xlsx", Decimal("400.00"), "首次拆分"))
    branch = project.branches["分支 A"]
    project.backups["分支 A / v1"] = branch.versions[0]
    project.register_branch("分支 A", Version("a-v2.xlsx", Decimal("590.00"), "模拟误差"))
    print("对账差额", project.reconcile())
    branch.restore(0)
    print("恢复旧版本后的差额", project.reconcile())
    assert len(project.rounds) == 2
    assert len(branch.versions) == 2
    assert project.reconcile() == Decimal("0.00")


if __name__ == "__main__":
    demo()
