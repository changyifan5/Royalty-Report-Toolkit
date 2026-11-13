"""
excel_font_color.py 颜色识别单元测试。

覆盖：
  - 颜色换算：hex_to_rgb / apply_tint_component / rgb_to_hex / color_hint
  - 校准阈值组（09/11 在用，读 *_font_calibration.json）：is_orange_font /
    is_red_font / is_purple_font / is_blue_font / is_gray_fill / is_gray_bold_cell
  - 严格阈值组（04 在用）：is_strict_red_font / is_strict_blue_font

标准校准色：橙 #C0504D（theme 5）、红 #FF0000、紫 #7030A0、蓝 #4BACC6（theme 8），
容差均为 12。

红蓝紫橙灰的关键边界色各至少 1 例；带样式的单元格用 openpyxl 现造的临时 xlsx
验证（见 TestGrayBoldCell），绝不触碰真实业务报表。
"""

import sys
import tempfile
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from openpyxl import Workbook  # noqa: E402
from openpyxl.styles import Color, Font, PatternFill  # noqa: E402

import excel_font_color as fc  # noqa: E402


def make_font(rgb_hex: str | None) -> Font:
    """用 openpyxl Font 构造一个显式 RGB 字体（None 表示默认黑色）。"""
    return Font(color=rgb_hex) if rgb_hex else Font()


def make_theme_font(theme: int, tint: float = 0.0) -> Font:
    """构造 theme 色字体（走校准项的 theme 匹配路径）。"""
    return Font(color=Color(theme=theme, tint=tint))


class TestHexRgb(unittest.TestCase):
    """hex_to_rgb：8 位（含 alpha）/6 位 / 非法长度解析。"""

    def test_8位带alpha取后6位(self):
        self.assertEqual(fc.hex_to_rgb("FF112233"), (17, 34, 51))

    def test_6位直接解析(self):
        self.assertEqual(fc.hex_to_rgb("FF0000"), (255, 0, 0))

    def test非法长度返回None(self):
        self.assertIsNone(fc.hex_to_rgb("11223"))
        self.assertIsNone(fc.hex_to_rgb(""))
        self.assertIsNone(fc.hex_to_rgb("1122334455"))


class TestTintComponent(unittest.TestCase):
    """apply_tint_component：theme 色的明暗分量计算。"""

    def test_tint为0原值返回(self):
        self.assertEqual(fc.apply_tint_component(100, 0.0), 100)

    def test负tint压暗(self):
        self.assertEqual(fc.apply_tint_component(255, -0.5), 128)

    def test正tint提亮(self):
        self.assertEqual(fc.apply_tint_component(0, 0.5), 128)

    def test不越界(self):
        self.assertEqual(fc.apply_tint_component(255, 0.5), 255)
        self.assertEqual(fc.apply_tint_component(0, -0.5), 0)


class TestRgbToHexAndHint(unittest.TestCase):
    """rgb_to_hex / color_hint 的展示层输出。"""

    def test_rgb_to_hex大写两位(self):
        self.assertEqual(fc.rgb_to_hex((255, 0, 0)), "#FF0000")
        self.assertEqual(fc.rgb_to_hex((16, 32, 48)), "#102030")

    def test_color_hint红色(self):
        self.assertEqual(fc.color_hint((200, 50, 50)), "红色")

    def test_color_hint蓝色(self):
        self.assertEqual(fc.color_hint((50, 50, 220)), "蓝色")

    def test_color_hint紫色(self):
        self.assertEqual(fc.color_hint((150, 80, 150)), "紫色")

    def test_color_hint黑色默认(self):
        self.assertEqual(fc.color_hint((0, 0, 0)), "黑色/默认")

    def test_color_hint其他(self):
        # 绿灰调：r/b 都不达红蓝阈值、g 又不满足紫的压低条件 → 其他
        self.assertEqual(fc.color_hint((100, 150, 100)), "其他")


class TestStrictFonts(unittest.TestCase):
    """严格组（04 在用）：r/g/b 阈值硬卡，橙/紫均不算。"""

    def test纯红命中严格红(self):
        self.assertTrue(fc.is_strict_red_font(make_font("FFFF0000")), "纯红 FFFF0000 应判为严格红")

    def test纯蓝命中严格蓝(self):
        self.assertTrue(fc.is_strict_blue_font(make_font("FF0000FF")), "纯蓝 FF0000FF 应判为严格蓝")

    def test橙不命中严格红(self):
        self.assertFalse(fc.is_strict_red_font(make_font("FFFFC000")), "橙色 g=192>100，不算严格红")

    def test深红8B0000不命中严格红(self):
        # r=139 < 180，严格组要求强红，边界示例
        self.assertFalse(fc.is_strict_red_font(make_font("FF8B0000")), "暗红 r=139 不达严格红阈值")

    def test严格红不命中严格蓝(self):
        self.assertFalse(fc.is_strict_blue_font(make_font("FFFF0000")))

    def test无色字体一律False(self):
        f = make_font(None)
        self.assertFalse(fc.is_strict_red_font(f))
        self.assertFalse(fc.is_strict_blue_font(f))


class TestLooseFonts(unittest.TestCase):
    """校准组（09/11 在用）：橙/红/紫/蓝按 *_font_calibration.json 阈值互相排除。

    标准色：橙 #C0504D（theme 5 / tint 0）、红 #FF0000、紫 #7030A0、蓝 #4BACC6
    （theme 8 / tint 0），容差均为 12。
    """

    def test校准橙命中is_orange(self):
        self.assertTrue(fc.is_orange_font(make_font("FFC0504D")), "校准橙 #C0504D 应判为橙")

    def test主题5橙命中is_orange(self):
        # 校准项 excel_theme=5 / excel_tint=0.0
        self.assertTrue(fc.is_orange_font(make_theme_font(5, 0.0)), "theme5/tint0 应判为橙")

    def test纯红命中is_red且排除橙(self):
        f = make_font("FFFF0000")
        self.assertTrue(fc.is_red_font(f), "纯红应判为红")
        self.assertFalse(fc.is_orange_font(f), "纯红不应误判为橙")

    def test校准橙不命中is_red(self):
        self.assertFalse(fc.is_red_font(make_font("FFC0504D")), "校准橙被 is_red 显式排除")

    def test校准紫命中is_purple(self):
        self.assertTrue(fc.is_purple_font(make_font("FF7030A0")), "校准紫 #7030A0 应判为紫")

    def test纯蓝命中is_blue且排除紫(self):
        f = make_font("FF0000FF")
        self.assertTrue(fc.is_blue_font(f), "纯蓝应判为蓝")
        self.assertFalse(fc.is_purple_font(f), "纯蓝 b=255/r=0 不达紫阈值")

    def test校准紫不命中is_blue(self):
        self.assertFalse(fc.is_blue_font(make_font("FF7030A0")), "校准紫被 is_blue 显式排除")

    def test主题8蓝命中is_blue(self):
        # 校准项 excel_theme=8 / excel_tint=0.0
        self.assertTrue(fc.is_blue_font(make_theme_font(8, 0.0)), "theme8/tint0 应判为蓝")

    def test无色字体校准组全False(self):
        f = make_font(None)
        for fn in (fc.is_orange_font, fc.is_red_font, fc.is_purple_font, fc.is_blue_font):
            self.assertFalse(fn(f), f"{fn.__name__} 对无色字体应返回 False")


class TestGrayFill(unittest.TestCase):
    """is_gray_fill：灰色实心背景（非纯白/纯黑/高饱和）。"""

    def test中灰命中(self):
        self.assertTrue(fc.is_gray_fill(PatternFill("solid", fgColor="FF808080")))

    def test白色不算灰(self):
        # avg=255 > 210
        self.assertFalse(fc.is_gray_fill(PatternFill("solid", fgColor="FFFFFFFF")))

    def test高饱和红底不算灰(self):
        self.assertFalse(fc.is_gray_fill(PatternFill("solid", fgColor="FFFF0000")))

    def test无填充不算灰(self):
        self.assertFalse(fc.is_gray_fill(PatternFill(fill_type=None)))


class TestGrayBoldCell(unittest.TestCase):
    """is_gray_bold_cell：灰色背景 + 加粗字体；用临时 xlsx 现造单元格验证。"""

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.path = Path(self._tmp.name) / "style_sample.xlsx"
        wb = Workbook()
        ws = wb.active
        # A1：灰底 + 加粗（应命中）
        c1 = ws.cell(row=1, column=1, value="命中")
        c1.font = Font(bold=True)
        c1.fill = PatternFill("solid", fgColor="FFBFBFBF")
        # A2：白底 + 加粗（不命中）
        c2 = ws.cell(row=2, column=1, value="白底加粗")
        c2.font = Font(bold=True)
        c2.fill = PatternFill("solid", fgColor="FFFFFFFF")
        # A3：灰底 + 不加粗（不命中）
        c3 = ws.cell(row=3, column=1, value="灰底普通")
        c3.font = Font(bold=False)
        c3.fill = PatternFill("solid", fgColor="FF808080")
        wb.save(self.path)
        wb.close()

    def tearDown(self):
        self._tmp.cleanup()

    def test灰底加粗命中(self):
        from openpyxl import load_workbook

        wb = load_workbook(self.path)
        ws = wb.active
        self.assertTrue(fc.is_gray_bold_cell(ws.cell(row=1, column=1)), "灰底+加粗应命中")
        self.assertFalse(fc.is_gray_bold_cell(ws.cell(row=2, column=1)), "白底不算灰")
        self.assertFalse(fc.is_gray_bold_cell(ws.cell(row=3, column=1)), "不加粗不算")
        wb.close()


if __name__ == "__main__":
    unittest.main()
