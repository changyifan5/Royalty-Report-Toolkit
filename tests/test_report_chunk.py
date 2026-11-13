"""
report_chunk.py 大报表分段单元测试。

覆盖：
  - compute_chunks（0 行 / 单段 / 多段边界）
  - row_in_chunk / format_chunk_label / RowChunk.is_first / is_final
  - timestamp_backup_path（与 excel_common 同一命名规则）
  - fast_sheet_last_row（现造小 xlsx，dimension 末行取值；多页签索引）

不触碰弹窗 / tkinter 交互函数（prompt_*），它们需 GUI，不在纯测试范围。
"""

import sys
import tempfile
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from openpyxl import Workbook  # noqa: E402

import report_chunk as rc  # noqa: E402


class TestComputeChunks(unittest.TestCase):
    """compute_chunks：按 chunk_size 切分数据行，Excel 行号带 data_start_row 偏移。"""

    def test零行返回空(self):
        self.assertEqual(rc.compute_chunks(0, 7), [])
        self.assertEqual(rc.compute_chunks(-5, 7), [])

    def test单段边界(self):
        chunks = rc.compute_chunks(5, excel_data_start_row=7, chunk_size=100_000)
        self.assertEqual(len(chunks), 1)
        c = chunks[0]
        self.assertTrue(c.is_first)
        self.assertTrue(c.is_final)
        self.assertEqual((c.data_row_start, c.data_row_end), (1, 5), "数据条序号从 1 起")
        self.assertEqual((c.data_start_row, c.data_end_row), (7, 11), "Excel 行号从第 7 行起")

    def test多段切分与边界衔接(self):
        # 250000 行、每段 100000 → 3 段
        chunks = rc.compute_chunks(250_000, excel_data_start_row=7, chunk_size=100_000)
        self.assertEqual(len(chunks), 3)
        self.assertEqual(
            [(c.data_start_row, c.data_end_row) for c in chunks],
            [(7, 100_006), (100_007, 200_006), (200_007, 250_006)],
        )
        self.assertTrue(chunks[0].is_first)
        self.assertTrue(chunks[-1].is_final)
        self.assertFalse(chunks[1].is_first)
        self.assertFalse(chunks[1].is_final)


class TestRowInChunk(unittest.TestCase):
    """row_in_chunk：Excel 行是否落在本段区间内。"""

    def test命中与越界(self):
        chunks = rc.compute_chunks(5, excel_data_start_row=7, chunk_size=100)
        c = chunks[0]
        self.assertTrue(rc.row_in_chunk(7, c))
        self.assertTrue(rc.row_in_chunk(11, c))
        self.assertFalse(rc.row_in_chunk(6, c), "data_start_row 之前不算本段")
        self.assertFalse(rc.row_in_chunk(12, c), "data_end_row 之后不算本段")


class TestFormatChunkLabel(unittest.TestCase):
    """format_chunk_label：给用户看的段标签文案。"""

    def test文案含条号与Excel行号(self):
        c = rc.compute_chunks(5, excel_data_start_row=7)[0]
        label = rc.format_chunk_label(c)
        self.assertIn("第 1～5 条", label)
        self.assertIn("Excel 第 7～11 行", label)


class TestTimestampBackupPath(unittest.TestCase):
    """report_chunk.timestamp_backup_path：与 excel_common 同一命名。"""

    def test命名规则(self):
        out = rc.timestamp_backup_path(Path("/data/report.xlsx"))
        self.assertRegex(
            out.name, r"^report\.bak-\d{8}\.xlsx$",
            "应为 report.bak-YYYYMMDD.xlsx",
        )


class TestFastSheetLastRow(unittest.TestCase):
    """fast_sheet_last_row：读 xlsx dimension 末行，现造小 xlsx 验证。"""

    def test单页签末行(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "dim.xlsx"
            wb = Workbook()
            ws = wb.active
            for r in range(1, 13):  # 写到第 12 行
                ws.cell(row=r, column=1, value=r)
            wb.save(path)
            wb.close()
            self.assertEqual(rc.fast_sheet_last_row(path, sheet_index=0), 12)

    def test第二页签索引(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "dim2.xlsx"
            wb = Workbook()
            ws1 = wb.active
            ws1.cell(row=1, column=1, value="只有一行")
            ws2 = wb.create_sheet("第二个")
            for r in range(1, 21):  # 第二页签到第 20 行
                ws2.cell(row=r, column=1, value=r)
            wb.save(path)
            wb.close()
            self.assertEqual(rc.fast_sheet_last_row(path, sheet_index=0), 1)
            self.assertEqual(rc.fast_sheet_last_row(path, sheet_index=1), 20)


class TestChunkProgressHelpers(unittest.TestCase):
    """all_chunks_completed / remaining_chunk_indices：进度集合的纯逻辑。"""

    def test全部完成判定(self):
        chunks = rc.compute_chunks(250_000, excel_data_start_row=7, chunk_size=100_000)
        from report_chunk import ChunkProgress

        prog = ChunkProgress(script_key="t", source="x", chunk_total=3, completed=[1, 2, 3])
        self.assertTrue(rc.all_chunks_completed(prog))
        self.assertEqual(rc.remaining_chunk_indices(prog), [])

    def test未完成段列表(self):
        from report_chunk import ChunkProgress

        prog = ChunkProgress(script_key="t", source="x", chunk_total=3, completed=[1])
        self.assertFalse(rc.all_chunks_completed(prog))
        self.assertEqual(rc.remaining_chunk_indices(prog), [2, 3])


if __name__ == "__main__":
    unittest.main()
