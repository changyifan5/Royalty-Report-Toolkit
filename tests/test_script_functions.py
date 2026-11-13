"""
编号脚本与 row_order_common 内纯函数的单元测试。

用 importlib 按文件路径加载带空格/数字开头的编号脚本（02/07/11/13），
只引用其无副作用的纯函数，绝不触发 run_* 主流程（主流程均在 __main__ 守卫下）。

覆盖：
  - 02 半自动除错：classify_error / cell_text（_ERROR_RE 四类错误串）
  - 07 G列比例拆分：split_value / parse_ratio / to_number
  - 11 街版拆分：split_formula / calc_street_g_value
  - 13 B列筛选G列汇总：parse_i_date / excel_serial_to_date
  - row_order_common：parse_seq（__原序 列解析）
"""

import importlib.util
import sys
import unittest
from datetime import date, datetime
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

import row_order_common  # noqa: E402


def _load_script(module_name: str, file_name: str):
    """按文件路径加载编号脚本；先登记 sys.modules，兼容其 dataclass 定义。"""
    path = REPO_ROOT / file_name
    spec = importlib.util.spec_from_file_location(module_name, path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = mod
    spec.loader.exec_module(mod)
    return mod


s02 = _load_script("ut_02", "02 半自动除错.py")
s07 = _load_script("ut_07", "07 G列比例拆分.py")
s11 = _load_script("ut_11", "11 街版拆分.py")
s13 = _load_script("ut_13", "13 B列筛选G列汇总.py")


# ---------------------------------------------------------------------------
# 02 半自动除错
# ---------------------------------------------------------------------------
class TestClassifyError(unittest.TestCase):
    """02 classify_error：_ERROR_RE 四类命名分组逐一命中，空/正常串返回 None。"""

    def test找不到演唱人(self):
        self.assertEqual(
            s02.classify_error("找不到名稱為張三的演唱人或藝人"), "找不到演唱人")

    def test找不到歌曲(self):
        self.assertEqual(
            s02.classify_error("找不到歌曲名稱為情歌的歌曲"), "找不到歌曲")

    def test找不到专辑(self):
        self.assertEqual(
            s02.classify_error("找不到專輯名稱為精選的專輯"), "找不到专辑")

    def test收益时间错误(self):
        self.assertEqual(
            s02.classify_error("收益時間2024/01/01不符合專輯上、下架（月底）區間"),
            "收益时间错误",
        )

    def test空串与空白串返回None(self):
        self.assertIsNone(s02.classify_error(""))
        self.assertIsNone(s02.classify_error("   "))

    def test正常文本返回None(self):
        self.assertIsNone(s02.classify_error("一切正常，没有错误标记"))

    def test带前后缀仍命中(self):
        self.assertEqual(
            s02.classify_error("行号12 前缀 找不到歌曲名稱為情歌的歌曲 后缀备注"),
            "找不到歌曲",
            "错误串出现在文本中间时仍应 search 命中",
        )


class TestCellText(unittest.TestCase):
    """02 cell_text：单元格值转展示文本；整浮点去小数。"""

    def test_None空串(self):
        self.assertEqual(s02.cell_text(None), "")

    def test整浮点变整数字符串(self):
        self.assertEqual(s02.cell_text(5.0), "5")

    def test去首尾空格(self):
        self.assertEqual(s02.cell_text("  abc "), "abc")


# ---------------------------------------------------------------------------
# 07 G列比例拆分
# ---------------------------------------------------------------------------
class TestSplitValue(unittest.TestCase):
    """07 split_value：按比例拆分，v1+v2 必须严格等于 total。"""

    def test_3比7拆分(self):
        v1, v2 = s07.split_value(10, 3, 7)
        self.assertAlmostEqual(v1, 3.0, places=6)
        self.assertAlmostEqual(v2, 7.0, places=6)
        self.assertAlmostEqual(v1 + v2, 10.0, places=9, msg="拆分两数之和必须还原原值")

    def test对半拆分(self):
        v1, v2 = s07.split_value(99, 1, 1)
        self.assertAlmostEqual(v1 + v2, 99.0, places=9)


class TestParseRatio(unittest.TestCase):
    """07 parse_ratio：冒号/逗号/空白三种写法，含全角冒号；非法输入抛 ValueError。"""

    def test冒号写法(self):
        self.assertEqual(s07.parse_ratio("3:7"), (3.0, 7.0))

    def test全角冒号(self):
        self.assertEqual(s07.parse_ratio("3：7"), (3.0, 7.0), "全角冒号应归一为半角")

    def test小数逗号写法(self):
        r1, r2 = s07.parse_ratio("0.3,0.7")
        self.assertAlmostEqual(r1, 0.3)
        self.assertAlmostEqual(r2, 0.7)

    def test空白分隔两数(self):
        self.assertEqual(s07.parse_ratio("3 7"), (3.0, 7.0))

    def test负数抛错(self):
        with self.assertRaises(ValueError):
            s07.parse_ratio("-1:2")

    def test双零抛错(self):
        with self.assertRaises(ValueError):
            s07.parse_ratio("0:0")

    def test单个数抛错(self):
        with self.assertRaises(ValueError):
            s07.parse_ratio("3")


class TestToNumber07(unittest.TestCase):
    """07 to_number：数字/千分位/空/非数字/布尔。"""

    def test数字直转float(self):
        self.assertEqual(s07.to_number(5), 5.0)

    def test千分位字符串(self):
        self.assertEqual(s07.to_number("1,234"), 1234.0)

    def test空与非数字返回None(self):
        self.assertIsNone(s07.to_number(None))
        self.assertIsNone(s07.to_number(""))
        self.assertIsNone(s07.to_number("abc"))

    def test布尔不算数字(self):
        self.assertIsNone(s07.to_number(False))


# ---------------------------------------------------------------------------
# 11 街版拆分
# ---------------------------------------------------------------------------
class TestSplitFormula(unittest.TestCase):
    """11 split_formula：按列号/行号生成两条 Excel 公式字符串。"""

    def test公式内容(self):
        f1, f2 = s11.split_formula(row=8, g_col=7, pai_col=8, jie_col=9, out_col1=10)
        self.assertEqual(f1, '=IF((H8+I8)=0,"",G8*H8/(H8+I8))')
        self.assertEqual(f2, '=IF((H8+I8)=0,"",G8-J8)')


class TestCalcStreetGValue(unittest.TestCase):
    """11 calc_street_g_value：G × 词曲示例版权方 / (词曲派歌+词曲示例版权方)。"""

    def test正常比例(self):
        # 100 × 70 / (30+70) = 70
        self.assertAlmostEqual(s11.calc_street_g_value(100, 30, 70), 70.0)

    def test另一比例(self):
        # 100 × 30 / (20+30) = 60
        self.assertAlmostEqual(s11.calc_street_g_value(100, 20, 30), 60.0)

    def test分母为零返回None(self):
        self.assertIsNone(s11.calc_street_g_value(100, 0, 0))

    def testG为空返回None(self):
        self.assertIsNone(s11.calc_street_g_value(None, 10, 10))


# ---------------------------------------------------------------------------
# 13 B列筛选G列汇总
# ---------------------------------------------------------------------------
class TestExcelSerialToDate(unittest.TestCase):
    """13 excel_serial_to_date：Excel 序列日期 ↔ 日历日，越界返回 None。"""

    def test已知序列映射(self):
        # Excel 序列 44927 = 2023-01-01（1899-12-30 起算）
        self.assertEqual(s13.excel_serial_to_date(44927), date(2023, 1, 1))

    def test下界内(self):
        # 下界 30000 ≈ 1982 年，在区间内即可转换（不断言具体日）
        self.assertIsNotNone(s13.excel_serial_to_date(30000))

    def test越界返回None(self):
        self.assertIsNone(s13.excel_serial_to_date(1000), "小于最小序列不算日期")
        self.assertIsNone(s13.excel_serial_to_date(70000), "大于最大序列不算日期")


class TestParseIDate(unittest.TestCase):
    """13 parse_i_date：I 列日期解析。

    当前实际语义（照实断言，不为测试改写源码）：
      - datetime / date 实例直接转 date
      - 整数 YYYYMM（190001~209912）→ 当月月末
      - 字符串只要去掉非数字后前 6 位是合法 YYYYMM，一律归当月月末
        （因此 "2024/03/15" 也会被归成 2024-03-31）
      - 其余整数走 Excel 序列日期；不合法一律 None
    """

    def test_None返回None(self):
        self.assertIsNone(s13.parse_i_date(None))

    def test_datetime取date部分(self):
        self.assertEqual(s13.parse_i_date(datetime(2024, 3, 15)), date(2024, 3, 15))

    def test_date原样返回(self):
        d = date(2024, 3, 31)
        self.assertEqual(s13.parse_i_date(d), d)

    def test整数YYYYMM归月末(self):
        self.assertEqual(s13.parse_i_date(202403), date(2024, 3, 31))

    def test浮点整数归月末(self):
        self.assertEqual(s13.parse_i_date(202403.0), date(2024, 3, 31))

    def test浮点非整数返回None(self):
        self.assertIsNone(s13.parse_i_date(202403.5))

    def test字符串YYYYMM归月末(self):
        self.assertEqual(s13.parse_i_date("202403"), date(2024, 3, 31))

    def test带斜杠日期串仍归月末(self):
        # 数字串分支优先：前 6 位 202403 → 月末
        self.assertEqual(s13.parse_i_date("2024/03/15"), date(2024, 3, 31))

    def test带空白字符串归月末(self):
        self.assertEqual(s13.parse_i_date("  202403  "), date(2024, 3, 31))

    def test非数字文本返回None(self):
        self.assertIsNone(s13.parse_i_date("abc"))
        self.assertIsNone(s13.parse_i_date(""))


# ---------------------------------------------------------------------------
# row_order_common：__原序 列解析
# ---------------------------------------------------------------------------
class TestParseSeq(unittest.TestCase):
    """row_order_common.parse_seq：__原序列的取值，容错空/布尔/非数字。"""

    def test整数(self):
        self.assertEqual(row_order_common.parse_seq(7), 7)

    def test浮点整数(self):
        self.assertEqual(row_order_common.parse_seq(7.0), 7)

    def test浮点带小数返回None(self):
        self.assertIsNone(row_order_common.parse_seq(7.5))

    def test字符串数字(self):
        self.assertEqual(row_order_common.parse_seq("7"), 7)

    def test字符串带空白(self):
        self.assertEqual(row_order_common.parse_seq("  8 "), 8)

    def test空与None(self):
        self.assertIsNone(row_order_common.parse_seq(None))
        self.assertIsNone(row_order_common.parse_seq(""))

    def test布尔不算序号(self):
        self.assertIsNone(row_order_common.parse_seq(True))
        self.assertIsNone(row_order_common.parse_seq(False))

    def test非数字文本返回None(self):
        self.assertIsNone(row_order_common.parse_seq("abc"))


if __name__ == "__main__":
    unittest.main()
