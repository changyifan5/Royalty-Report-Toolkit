"""
10 列映射填表.py —— 运行时多选中介文件页签功能的单元测试。

用 importlib 按文件路径加载带空格的编号脚本（与 test_script_functions.py 同款），
只测无副作用的纯逻辑与「单个页签完整写入」函数，绝不触发 run() 主流程。

覆盖：
  - list_intermediate_sheet_names：zip 直读页签名，顺序正确
  - read_columns 指定页签：不同页签读到不同表头
  - sheet_headers_identical / header_signature：表头一致性（含 normalize 边界）
  - write_sheet_rows：多页签循环写入，last_row/写入位置独立、一次保存生效
  - write_sheet_rows 取消语义：写入位置框返回 None 时整体返回 None

注意：弹窗交互（pick_intermediate_sheets 多选、ask_reuse_mapping 复用询问、
ask_continue_next_sheet 继续询问、pick_write_start_row、MappingDialog）涉及真实
tkinter，本文件不做自动化，只测其背后的纯逻辑函数。
"""

import importlib.util
import sys
import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from openpyxl import Workbook, load_workbook  # noqa: E402


def _load_script(module_name: str, file_name: str):
    """按文件路径加载编号脚本；先登记 sys.modules，兼容其 dataclass 定义。"""
    path = REPO_ROOT / file_name
    spec = importlib.util.spec_from_file_location(module_name, path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = mod
    spec.loader.exec_module(mod)
    return mod


s10 = _load_script("ut_10", "10 列映射填表.py")


HEADER_ROW = 6  # 与脚本默认表头行一致


def _col(index: int, header: str) -> "object":
    from openpyxl.utils import get_column_letter
    return s10.ColumnInfo(index=index, letter=get_column_letter(index), header=header, display=header)


# ---------------------------------------------------------------------------
# 页签列表读取
# ---------------------------------------------------------------------------
class TestListSheetNames(unittest.TestCase):
    """list_intermediate_sheet_names：现造多页签 xlsx，返回页签名与顺序。"""

    def test多页签顺序正确(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "multi.xlsx"
            wb = Workbook()
            wb.active.title = "主表"
            wb.create_sheet("街版")
            wb.create_sheet("临时")
            wb.save(path)
            wb.close()
            self.assertEqual(
                s10.list_intermediate_sheet_names(path),
                ["主表", "街版", "临时"],
            )


# ---------------------------------------------------------------------------
# read_columns 指定页签
# ---------------------------------------------------------------------------
class TestReadColumnsBySheet(unittest.TestCase):
    """read_columns(path, header_row, sheet=...)：不同页签读到各自表头。"""

    def test不同页签读到不同表头(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "inter.xlsx"
            wb = Workbook()
            ws1 = wb.active
            ws1.title = "甲"
            ws1.cell(row=HEADER_ROW, column=1, value="歌曲名")
            ws1.cell(row=HEADER_ROW, column=2, value="演唱者")
            ws2 = wb.create_sheet("乙")
            ws2.cell(row=HEADER_ROW, column=1, value="产品名")
            ws2.cell(row=HEADER_ROW, column=2, value="数量")
            wb.save(path)
            wb.close()

            cols_jia = s10.read_columns(path, HEADER_ROW, sheet="甲")
            cols_yi = s10.read_columns(path, HEADER_ROW, sheet="乙")
            self.assertEqual([c.header for c in cols_jia[:2]], ["歌曲名", "演唱者"])
            self.assertEqual([c.header for c in cols_yi[:2]], ["产品名", "数量"])

    def test按序号与按名字一致(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "inter.xlsx"
            wb = Workbook()
            wb.active.title = "甲"
            wb.active.cell(row=HEADER_ROW, column=1, value="歌曲名")
            wb.create_sheet("乙").cell(row=HEADER_ROW, column=1, value="产品名")
            wb.save(path)
            wb.close()
            by_index = s10.read_columns(path, HEADER_ROW, sheet=1)
            by_name = s10.read_columns(path, HEADER_ROW, sheet="乙")
            self.assertEqual([c.header for c in by_index], [c.header for c in by_name])


# ---------------------------------------------------------------------------
# 表头一致性
# ---------------------------------------------------------------------------
class TestSheetHeadersIdentical(unittest.TestCase):
    """sheet_headers_identical：列数 + 归一后表头文字全等才视为相同。"""

    def test完全相同为真(self):
        a = [_col(1, "歌曲名"), _col(2, "演唱者")]
        b = [_col(1, "歌曲名"), _col(2, "演唱者")]
        self.assertTrue(s10.sheet_headers_identical(a, b))

    def test归一化后相同为真(self):
        # 大小写/空白差异经 normalize_header_key 归一后应相等
        a = [_col(1, "Track Name"), _col(2, "Artist")]
        b = [_col(1, "trackname"), _col(2, " artist ")]
        self.assertTrue(
            s10.sheet_headers_identical(a, b),
            "大小写与首尾空白差异应被 normalize_header_key 抹平",
        )

    def test列数不同为假(self):
        a = [_col(1, "歌曲名"), _col(2, "演唱者")]
        b = [_col(1, "歌曲名")]
        self.assertFalse(s10.sheet_headers_identical(a, b))

    def test表头文字不同为假(self):
        a = [_col(1, "歌曲名"), _col(2, "演唱者")]
        b = [_col(1, "歌曲名"), _col(2, "专辑名")]
        self.assertFalse(s10.sheet_headers_identical(a, b))


# ---------------------------------------------------------------------------
# 多页签循环写入
# ---------------------------------------------------------------------------
def _build_source(td: Path) -> Path:
    """源文件：第 1 行表头，2 行数据（A=曲名整数, B=歌手, C=YYYYMM）。"""
    src = td / "源.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet"
    ws.cell(row=1, column=1, value="曲名")
    ws.cell(row=1, column=2, value="歌手")
    ws.cell(row=1, column=3, value="月份")
    ws.cell(row=2, column=1, value=20230701001)
    ws.cell(row=2, column=2, value="陈奕迅")
    ws.cell(row=2, column=3, value=202307)
    ws.cell(row=3, column=1, value=20230701002)
    ws.cell(row=3, column=2, value="周杰伦")
    ws.cell(row=3, column=3, value="202308")
    wb.save(src)
    wb.close()
    return src


def _build_workbook(td: Path) -> Path:
    """多页签中介工作文件：甲/乙表头相同，丙表头不同；甲在第 8 行预置一条旧数据。"""
    path = td / "work.xlsx"
    wb = Workbook()
    jia = wb.active
    jia.title = "甲"
    jia.cell(row=HEADER_ROW, column=1, value="歌曲名")
    jia.cell(row=HEADER_ROW, column=2, value="演唱者")
    jia.cell(row=HEADER_ROW, column=9, value="月份")
    jia.cell(row=8, column=1, value="旧数据")  # 预置：甲已有数据到第 8 行

    yi = wb.create_sheet("乙")
    yi.cell(row=HEADER_ROW, column=1, value="歌曲名")
    yi.cell(row=HEADER_ROW, column=2, value="演唱者")
    yi.cell(row=HEADER_ROW, column=9, value="月份")  # 与甲完全一致

    bing = wb.create_sheet("丙")
    bing.cell(row=HEADER_ROW, column=1, value="产品名")
    bing.cell(row=HEADER_ROW, column=2, value="数量")  # 与甲不同

    wb.save(path)
    wb.close()
    return path


class TestMultiSheetWrite(unittest.TestCase):
    """write_sheet_rows：逐页签写入，写入位置独立、一次保存后各页签结果正确。"""

    def test逐页签写入且写入位置独立(self):
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            src = _build_source(td)
            work_path = _build_workbook(td)

            cols_jia = [_col(1, "歌曲名"), _col(2, "演唱者"), _col(9, "月份")]
            cols_bing = [_col(1, "产品名"), _col(2, "数量")]

            mapping_ab = s10.FileMapping(source_path=src, header_row=1, pairs=[(1, 1), (2, 2), (3, 9)])
            mapping_c = s10.FileMapping(source_path=src, header_row=1, pairs=[(1, 1), (2, 2)])

            # 写入位置询问：模拟「追加到 last_row+1」。甲已有数据到第8行 -> 从第9行写；
            # 乙无旧数据 -> last_row=6 < data_start=7 -> 从第7行写。
            def ask(header_row, last_row):
                data_start = header_row + 1
                return data_start if last_row < data_start else last_row + 1

            wb = load_workbook(work_path)
            try:
                r_jia = s10.write_sheet_rows(wb, "甲", cols_jia, [mapping_ab], HEADER_ROW, ask_write_position=ask)
                r_yi = s10.write_sheet_rows(wb, "乙", cols_jia, [mapping_ab], HEADER_ROW, ask_write_position=ask)
                r_bing = s10.write_sheet_rows(wb, "丙", cols_bing, [mapping_c], HEADER_ROW, ask_write_position=ask)

                # 甲：last_row=8 -> 追加从第 9 行起
                self.assertEqual(r_jia["first_row"], 9)
                self.assertEqual(r_jia["written"], 2)
                # 乙：无旧数据 -> 从第 7 行起（与甲独立）
                self.assertEqual(r_yi["first_row"], 7)
                self.assertEqual(r_yi["written"], 2)
                # 丙：从第 7 行起
                self.assertEqual(r_bing["first_row"], 7)
                self.assertEqual(r_bing["written"], 2)

                wb.save(work_path)  # 全部页签写完一次保存
            finally:
                wb.close()

            # 重新打开校验三个页签都写入了
            wb = load_workbook(work_path)
            try:
                jia = wb["甲"]
                # 甲从第 9 行写：A9 文本、I9 月末日期
                self.assertEqual(jia.cell(row=9, column=1).value, "20230701001")
                self.assertEqual(jia.cell(row=9, column=1).number_format, "@")
                self.assertEqual(jia.cell(row=9, column=9).value, datetime(2023, 7, 31))
                self.assertEqual(jia.cell(row=9, column=9).number_format, "yyyy/mm/dd")
                # 旧数据第 8 行仍在
                self.assertEqual(jia.cell(row=8, column=1).value, "旧数据")

                yi = wb["乙"]
                self.assertEqual(yi.cell(row=7, column=1).value, "20230701001")
                self.assertEqual(yi.cell(row=7, column=9).value, datetime(2023, 7, 31))

                bing = wb["丙"]
                self.assertEqual(bing.cell(row=7, column=1).value, "20230701001")
                self.assertEqual(bing.cell(row=7, column=2).value, "陈奕迅")
            finally:
                wb.close()

    def test写入位置取消返回None(self):
        """写入位置框返回 None（用户取消）时，write_sheet_rows 返回 None（退出全部语义）。"""
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            src = _build_source(td)
            work_path = _build_workbook(td)
            cols = [_col(1, "歌曲名"), _col(2, "演唱者")]
            mapping = s10.FileMapping(source_path=src, header_row=1, pairs=[(1, 1), (2, 2)])

            wb = load_workbook(work_path)
            try:
                result = s10.write_sheet_rows(
                    wb, "甲", cols, [mapping], HEADER_ROW,
                    ask_write_position=lambda header_row, last_row: None,
                )
                self.assertIsNone(result, "写入位置框取消应返回 None，供调用方退出全部")
            finally:
                wb.close()


if __name__ == "__main__":
    unittest.main()
