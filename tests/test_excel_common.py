"""
excel_common.py 公共基础设施单元测试。

覆盖：
  - normalize_key / find_header_last_col / read_header_map（表头解析，现造小 xlsx）
  - timestamp_backup_path / backup_file（report.xlsx → report.bak-YYYYMMDD.xlsx）
  - tk_ui_font / tk_mono_font（Mac 下 PingFang SC / Menlo）
  - list_sheet_names_fast（zip 直读页签名，与 openpyxl sheetnames 对照）

临时 xlsx 全部在 TemporaryDirectory 内构造，用完即清。
"""

import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from openpyxl import Workbook, load_workbook  # noqa: E402

import excel_common as ec  # noqa: E402


def _make_header_xlsx(path: Path) -> Path:
    """造一个标准布局小报表：第 6 行表头，A/B/C 列有值，C 列空，E 列有值。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "数据"
    ws.cell(row=6, column=1, value="歌曲")
    ws.cell(row=6, column=2, value="演唱者")
    # 第 3 列留空（模拟表头缺口）
    ws.cell(row=6, column=5, value="专辑")
    ws.cell(row=7, column=1, value="数据行1")
    wb.save(path)
    wb.close()
    return path


class TestNormalizeKey(unittest.TestCase):
    """normalize_key：去首尾空格，空串归一为 None。"""

    def test_去空格(self):
        self.assertEqual(ec.normalize_key("  abc "), "abc")

    def test_None与空白归None(self):
        self.assertIsNone(ec.normalize_key(None))
        self.assertIsNone(ec.normalize_key("   "))

    def test数字转字符串(self):
        self.assertEqual(ec.normalize_key(123), "123")


class TestHeaderScan(unittest.TestCase):
    """find_header_last_col / read_header_map：基于现造小 xlsx。"""

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.path = _make_header_xlsx(Path(self._tmp.name) / "hdr.xlsx")
        self.wb = load_workbook(self.path)
        self.ws = self.wb.active

    def tearDown(self):
        self.wb.close()
        self._tmp.cleanup()

    def test最后非空表头列号跳过缺口(self):
        # A=1 有、B=2 有、C=3 空、E=5 有 → 最后一列是 5
        self.assertEqual(ec.find_header_last_col(self.ws, header_row=6), 5)

    def test表头映射文字到列号(self):
        mapping = ec.read_header_map(self.ws, header_row=6)
        self.assertEqual(mapping, {"歌曲": 1, "演唱者": 2, "专辑": 5})


class TestTimestampBackup(unittest.TestCase):
    """timestamp_backup_path：report.xlsx → report.bak-YYYYMMDD.xlsx。"""

    def test命名规则带固定日期(self):
        src = Path("/tmp/report.xlsx")
        out = ec.timestamp_backup_path(src, when=datetime(2026, 9, 29, 15, 30))
        self.assertEqual(out.name, "report.bak-20260929.xlsx")
        self.assertEqual(out.suffix, ".xlsx", "备份仍以原扩展名结尾，落在 git 忽略规则内")

    def testbackup_file实际复制(self):
        with tempfile.TemporaryDirectory() as td:
            src = Path(td) / "report.xlsx"
            src.write_bytes(b"fake xlsx body")
            dst = ec.backup_file(src)
            self.assertTrue(dst.exists(), "backup_file 应真的复制出备份文件")
            self.assertTrue(dst.name.startswith("report.bak-"))
            self.assertEqual(dst.read_bytes(), b"fake xlsx body")


class TestTkFont(unittest.TestCase):
    """tk_ui_font / tk_mono_font：Mac 下应为 PingFang SC / Menlo。"""

    def test界面字体Mac族(self):
        self.assertEqual(ec.tk_ui_font(10), ("PingFang SC", 10))
        self.assertEqual(ec.tk_ui_font(9, "bold"), ("PingFang SC", 9, "bold"))

    def test等宽字体Mac族(self):
        self.assertEqual(ec.tk_mono_font(10), ("Menlo", 10))


class TestListSheetNamesFast(unittest.TestCase):
    """list_sheet_names_fast：zip 直读 workbook.xml，与 openpyxl sheetnames 一致。"""

    def test多页签一致(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "multi.xlsx"
            wb = Workbook()
            wb.active.title = "主表"
            wb.create_sheet("街版")
            wb.create_sheet("临时")
            wb.save(path)
            wb.close()

            self.assertEqual(
                ec.list_sheet_names_fast(path),
                ["主表", "街版", "临时"],
                "zip 直读页签名应与 openpyxl 的 sheetnames 完全一致",
            )


if __name__ == "__main__":
    unittest.main()
