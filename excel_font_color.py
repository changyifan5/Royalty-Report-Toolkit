"""Excel 字体/填充颜色识别（04 / 07 / 09 / 11 共用）。"""

from __future__ import annotations

import json
from pathlib import Path

from openpyxl.styles import Font

_BASE_DIR = Path(__file__).resolve().parent
_CALIBRATION_PATH = _BASE_DIR / "orange_font_calibration.json"
_DEFAULT_ORANGE_RGB = (192, 80, 77)
_DEFAULT_ORANGE_TOLERANCE = 12


def _load_font_calibration(filename: str, default_rgb: tuple[int, int, int], default_tol: int, default_theme: int | None, default_tint: float = 0.0) -> tuple[tuple[int, int, int], int, int | None, float]:
    path = _BASE_DIR / filename
    if not path.is_file():
        return default_rgb, default_tol, default_theme, default_tint
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        rgb = tuple(int(x) for x in data["rgb"])
        tolerance = int(data.get("tolerance", default_tol))
        theme_raw = data.get("excel_theme", default_theme)
        theme = int(theme_raw) if theme_raw is not None else None
        tint = float(data.get("excel_tint", default_tint))
        if len(rgb) != 3:
            raise ValueError("rgb must have 3 components")
        return rgb, tolerance, theme, tint
    except (OSError, ValueError, KeyError, TypeError, json.JSONDecodeError):
        return default_rgb, default_tol, default_theme, default_tint


def _load_orange_calibration() -> tuple[tuple[int, int, int], int, int, float]:
    rgb, tol, theme, tint = _load_font_calibration(
        "orange_font_calibration.json", _DEFAULT_ORANGE_RGB, _DEFAULT_ORANGE_TOLERANCE, 5, 0.0
    )
    return rgb, tol, int(theme if theme is not None else 5), tint


ORANGE_FONT_RGB, ORANGE_FONT_TOLERANCE, ORANGE_FONT_THEME, ORANGE_FONT_TINT = (
    _load_orange_calibration()
)
ORANGE_FONT_HEX = f"#{ORANGE_FONT_RGB[0]:02X}{ORANGE_FONT_RGB[1]:02X}{ORANGE_FONT_RGB[2]:02X}"

RED_FONT_RGB, RED_FONT_TOLERANCE, RED_FONT_THEME, RED_FONT_TINT = _load_font_calibration(
    "red_font_calibration.json", (255, 0, 0), 12, None, 0.0
)
BLUE_FONT_RGB, BLUE_FONT_TOLERANCE, BLUE_FONT_THEME, BLUE_FONT_TINT = _load_font_calibration(
    "blue_font_calibration.json", (75, 172, 198), 12, 8, 0.0
)
PURPLE_FONT_RGB, PURPLE_FONT_TOLERANCE, PURPLE_FONT_THEME, PURPLE_FONT_TINT = _load_font_calibration(
    "purple_font_calibration.json", (112, 48, 160), 12, None, 0.0
)

try:
    from openpyxl.styles.colors import COLOR_INDEX
except ImportError:
    COLOR_INDEX = ()

DEFAULT_THEME_HEX = (
    "000000",
    "FFFFFF",
    "1F497D",
    "EEECE1",
    "4F81BD",
    "C0504D",
    "9BBB59",
    "8064A2",
    "4BACC6",
    "F79646",
    "0000FF",
    "800080",
)


def apply_tint_component(value: int, tint: float) -> int:
    if not tint:
        return value
    if tint < 0:
        return max(0, min(255, int(round(value * (1.0 + tint)))))
    return max(0, min(255, int(round(value + (255 - value) * tint))))


def hex_to_rgb(hex_rgb: str) -> tuple[int, int, int] | None:
    if len(hex_rgb) == 8:
        hex_rgb = hex_rgb[2:]
    if len(hex_rgb) != 6:
        return None
    return int(hex_rgb[0:2], 16), int(hex_rgb[2:4], 16), int(hex_rgb[4:6], 16)


def openpyxl_color_rgb(color) -> tuple[int, int, int] | None:
    if color is None:
        return None
    if color.type == "rgb" and color.rgb:
        return hex_to_rgb(str(color.rgb))
    if color.type == "indexed" and color.indexed is not None:
        idx = int(color.indexed)
        if 0 <= idx < len(COLOR_INDEX):
            return hex_to_rgb(COLOR_INDEX[idx])
    if color.type == "theme" and color.theme is not None:
        theme = int(color.theme)
        tint = float(color.tint or 0.0)
        if 0 <= theme < len(DEFAULT_THEME_HEX):
            base = hex_to_rgb(DEFAULT_THEME_HEX[theme])
            if base is None:
                return None
            return (
                apply_tint_component(base[0], tint),
                apply_tint_component(base[1], tint),
                apply_tint_component(base[2], tint),
            )
    return None


def font_rgb_triple(font: Font) -> tuple[int, int, int] | None:
    if font is None or font.color is None:
        return None
    return openpyxl_color_rgb(font.color)


def fill_rgb_triple(fill) -> tuple[int, int, int] | None:
    if fill is None:
        return None
    fill_type = getattr(fill, "fill_type", None) or getattr(fill, "patternType", None)
    if fill_type in (None, "none"):
        return None
    color = getattr(fill, "fgColor", None) or getattr(fill, "start_color", None)
    return openpyxl_color_rgb(color)


def rgb_matches_calibrated(
    rgb: tuple[int, int, int],
    ref_rgb: tuple[int, int, int],
    tolerance: int,
) -> bool:
    ref_r, ref_g, ref_b = ref_rgb
    r, g, b = rgb
    return (
        abs(r - ref_r) <= tolerance
        and abs(g - ref_g) <= tolerance
        and abs(b - ref_b) <= tolerance
    )


def rgb_matches_orange(rgb: tuple[int, int, int]) -> bool:
    """与用户校准的标准橙色 (#C0504D) 比较，允许小幅色差。"""
    return rgb_matches_calibrated(rgb, ORANGE_FONT_RGB, ORANGE_FONT_TOLERANCE)


def _is_calibrated_theme_font(font: Font, theme: int | None, tint: float, ref_rgb: tuple[int, int, int], tolerance: int) -> bool:
    if font is None or font.color is None or theme is None:
        return False
    color = font.color
    if color.type == "theme" and color.theme is not None:
        if int(color.theme) == theme:
            cell_tint = float(color.tint or 0.0)
            if abs(cell_tint - tint) < 0.001:
                return True
            rgb = openpyxl_color_rgb(color)
            return rgb is not None and rgb_matches_calibrated(rgb, ref_rgb, tolerance)
    rgb = font_rgb_triple(font)
    return rgb is not None and rgb_matches_calibrated(rgb, ref_rgb, tolerance)


def is_orange_font(font: Font) -> bool:
    """标准橙色字体（#C0504D / theme 5），匹配 orange_font_calibration.json。"""
    if font is None or font.color is None:
        return False
    color = font.color
    if color.type == "theme" and color.theme is not None:
        if int(color.theme) == ORANGE_FONT_THEME:
            tint = float(color.tint or 0.0)
            if abs(tint - ORANGE_FONT_TINT) < 0.001:
                return True
            rgb = openpyxl_color_rgb(color)
            return rgb is not None and rgb_matches_orange(rgb)
    if color.type == "rgb" and color.rgb:
        raw = str(color.rgb).upper()
        if raw.endswith(ORANGE_FONT_HEX.lstrip("#")):
            return True
    rgb = font_rgb_triple(font)
    return rgb is not None and rgb_matches_orange(rgb)


def is_automatic_font(font: Font) -> bool:
    """openpyxl 中无显式字体色（Automatic / 默认）。"""
    if font is None or font.color is None:
        return True
    color = font.color
    if color.type == "indexed" and color.indexed is not None and int(color.indexed) >= 32767:
        return True
    return False


def is_explicit_orange_font(font: Font) -> bool:
    """仅当单元格显式设置了橙色字体时返回 True。"""
    if is_automatic_font(font):
        return False
    return is_orange_font(font)


def font_rgb_label(font: Font) -> str:
    rgb = font_rgb_triple(font)
    if rgb is None:
        return "默认/无"
    r, g, b = rgb
    return f"#{r:02X}{g:02X}{b:02X}"


def is_red_font(font: Font) -> bool:
    """纯红字体；排除橙。匹配 red_font_calibration.json。"""
    if is_orange_font(font):
        return False
    if RED_FONT_THEME is not None:
        if _is_calibrated_theme_font(font, RED_FONT_THEME, RED_FONT_TINT, RED_FONT_RGB, RED_FONT_TOLERANCE):
            return True
    rgb = font_rgb_triple(font)
    if rgb is None:
        return False
    return rgb_matches_calibrated(rgb, RED_FONT_RGB, RED_FONT_TOLERANCE)


def is_purple_font(font: Font) -> bool:
    """紫色字体。仅匹配 purple_font_calibration.json。"""
    if PURPLE_FONT_THEME is not None:
        if _is_calibrated_theme_font(font, PURPLE_FONT_THEME, PURPLE_FONT_TINT, PURPLE_FONT_RGB, PURPLE_FONT_TOLERANCE):
            return True
    rgb = font_rgb_triple(font)
    if rgb is None:
        return False
    return rgb_matches_calibrated(rgb, PURPLE_FONT_RGB, PURPLE_FONT_TOLERANCE)


def is_blue_font(font: Font) -> bool:
    """蓝色字体（排除紫）。匹配 blue_font_calibration.json，并兼容常见变体。"""
    if is_purple_font(font):
        return False
    if BLUE_FONT_THEME is not None:
        if _is_calibrated_theme_font(font, BLUE_FONT_THEME, BLUE_FONT_TINT, BLUE_FONT_RGB, BLUE_FONT_TOLERANCE):
            return True
    rgb = font_rgb_triple(font)
    if rgb is not None and rgb_matches_calibrated(rgb, BLUE_FONT_RGB, BLUE_FONT_TOLERANCE):
        return True
    if rgb is None:
        return False
    r, g, b = rgb
    return b >= 150 and r <= 130 and b > r + 30 and b >= g + 20


def is_gray_fill(fill) -> bool:
    """灰色实心背景（非纯白/纯黑）。"""
    rgb = fill_rgb_triple(fill)
    if rgb is None:
        return False
    r, g, b = rgb
    if max(r, g, b) - min(r, g, b) > 45:
        return False
    avg = (r + g + b) / 3
    return 70 <= avg <= 210


def is_gray_bold_cell(cell) -> bool:
    """A～D 等：灰色背景 + 加粗字体。"""
    font = cell.font
    if font is None or not font.bold:
        return False
    return is_gray_fill(cell.fill)


def is_font_bold(font: Font | None) -> bool:
    return bool(font and font.bold)


# ---------------------------------------------------------------------------
# 07 颜色展示辅助（rgb_to_hex / color_hint）
# ---------------------------------------------------------------------------
def rgb_to_hex(rgb: tuple[int, int, int]) -> str:
    """(r,g,b) → '#RRGGBB'（07 颜色选择列表展示用）。"""
    return f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"


def color_hint(rgb: tuple[int, int, int]) -> str:
    """给用户看的颜色粗略描述（07 扫描结果展示用）。"""
    r, g, b = rgb
    if r >= 180 and g <= 100 and b <= 100:
        return "红色"
    if b >= 180 and r <= 100 and b > g:
        return "蓝色"
    if r >= 100 and b >= 100 and g <= min(r, b) - 40:
        return "紫色"
    if r <= 80 and g <= 80 and b <= 80:
        return "黑色/默认"
    return "其他"


# ---------------------------------------------------------------------------
# 严格阈值组（04 退回修改标记在用）：要求明确的纯红 / 纯蓝
# ---------------------------------------------------------------------------
def is_strict_red_font(font: Font) -> bool:
    """04 用：r>=180 且 g,b<=100 的强红（不排除橙，阈值与宽松组不同）。"""
    rgb = font_rgb_triple(font)
    if rgb is None:
        return False
    r, g, b = rgb
    return r >= 180 and g <= 100 and b <= 100 and r > g and r > b


def is_strict_blue_font(font: Font) -> bool:
    """04 用：b>=180 且 r<=100 的强蓝。"""
    rgb = font_rgb_triple(font)
    if rgb is None:
        return False
    r, g, b = rgb
    return b >= 180 and r <= 100 and b > r and b > g
