"""设计系统：色彩（OKLCH）、字体、字号阶梯、12 栏网格。

目标不是"多几个颜色"，而是让版式有**可复用的秩序**：
    - 颜色从单一强调色推导（明度/彩度在 OKLCH 里按比例走，肉眼更均匀）
    - 字号走固定比例（1.25），标题/正文/注释只有 5 档
    - 所有坐标落在 12 栏网格上，左右边距 1in，栏距 0.28in
"""
from __future__ import annotations

import math
from dataclasses import dataclass

from pptx.dml.color import RGBColor

# ---------------------------------------------------------------- 字体

SANS = "Noto Sans SC"       # 思源黑体
SERIF = "Noto Serif SC"     # 思源宋体
NUM = "Inter"               # 数字/西文用 Inter，tabular 更整齐

# ---------------------------------------------------------------- 网格

SLIDE_W, SLIDE_H = 13.333, 7.5
MARGIN_X = 1.0
MARGIN_TOP = 0.85
COL_GUTTER = 0.28
COLS = 12


def col_x(i: int) -> float:
    """第 i 栏（0 起）的左边距（英寸）。"""
    w = content_w()
    return MARGIN_X + i * (w + COL_GUTTER) / COLS


def col_w(n: int) -> float:
    """跨 n 栏的宽度（英寸）。"""
    unit = (content_w() + COL_GUTTER) / COLS
    return unit * n - COL_GUTTER


def content_w() -> float:
    return SLIDE_W - 2 * MARGIN_X


def row_y(n: float) -> float:
    """从版心顶部往下 n 英寸。"""
    return MARGIN_TOP + n


# ---------------------------------------------------------------- 色彩

def _to_hex(rgb: tuple[float, float, float]) -> str:
    return "#%02X%02X%02X" % tuple(max(0, min(255, round(c * 255))) for c in rgb)


def oklch(lightness: float, chroma: float, hue: float) -> str:
    """OKLCH → #RRGGBB（超出 sRGB 时按最近色裁剪）。"""
    try:
        from coloraide import Color
        c = Color("oklch", [lightness, chroma, hue]).fit(method="clip").convert("srgb")
        return c.to_string(hex=True)[:7].upper()
    except Exception:
        # 退化到近似：把 L 当 HSL 明度用
        import colorsys
        r, g, b = colorsys.hls_to_rgb(hue / 360.0, lightness, min(1.0, chroma * 2))
        return _to_hex((r, g, b))


def to_oklch(hex_color: str) -> tuple[float, float, float]:
    try:
        from coloraide import Color
        c = Color(hex_color).convert("oklch")
        l, ch, h = c.coords()
        return (float(l), float(ch), float(h if h == h else 0))
    except Exception:
        return (0.5, 0.05, 260.0)


def rgb(hex_color: str) -> RGBColor:
    v = hex_color.lstrip("#")
    return RGBColor(int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16))


def luminance(hex_color: str) -> float:
    v = hex_color.lstrip("#")
    chans = []
    for i in (0, 2, 4):
        c = int(v[i:i + 2], 16) / 255
        chans.append(c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4)
    return 0.2126 * chans[0] + 0.7152 * chans[1] + 0.0722 * chans[2]


def contrast(fg: str, bg: str) -> float:
    a, b = luminance(fg), luminance(bg)
    hi, lo = max(a, b), min(a, b)
    return (hi + 0.05) / (lo + 0.05)


@dataclass(frozen=True)
class Palette:
    bg: str
    ink: str
    muted: str
    accent: str
    accent_soft: str
    rule: str
    reverse: str = "#FFFFFF"

    def on_accent(self) -> str:
        """强调色上的文字颜色：按对比度自动选黑白。"""
        return self.reverse if contrast(self.reverse, self.accent) >= 3.2 else self.ink


def palette(accent: str = "#1B3A6B", *, mode: str = "paper") -> Palette:
    """从一个强调色推导整套配色。

    mode:
        paper  米白纸感（偏暖，适合分享/杂志）
        clean  纯白（适合商务/学术）
    """
    l, c, h = to_oklch(accent)
    hue = h
    chroma = min(c, 0.16)
    if mode == "paper":
        bg = oklch(0.978, 0.006, hue + 40)
        rule = oklch(0.895, 0.008, hue)
        soft = oklch(0.945, 0.028, hue)
    else:
        bg = "#FFFFFF"
        rule = oklch(0.905, 0.006, hue)
        soft = oklch(0.955, 0.022, hue)
    ink = oklch(0.235, 0.028, hue)
    muted = oklch(0.545, 0.022, hue)
    # 强调色稍微压暗一点，保证与白底对比 ≥ 6:1
    accent_c = accent
    if contrast(accent_c, bg) < 6.0:
        accent_c = oklch(max(0.30, l - 0.06), chroma, hue)
    return Palette(bg=bg, ink=ink, muted=muted, accent=accent_c,
                   accent_soft=soft, rule=rule)


# ---------------------------------------------------------------- 字号

RATIO = 1.25
BASE_PT = 18.0


def type_size(step: int = 0) -> float:
    """step=0 正文 18pt；每 +1 乘 1.25；每 -1 除 1.25。"""
    return round(BASE_PT * (RATIO ** step), 1)


T_BODY = type_size(0)      # 18
T_SMALL = type_size(-1)    # 14.4
T_CAPTION = type_size(-2)  # 11.5
T_H4 = type_size(1)        # 22.5
T_H3 = type_size(2)        # 28.1
T_H2 = type_size(3)        # 35.2
T_H1 = type_size(4)        # 44
T_DISPLAY = type_size(5)   # 55

# 行距：正文 1.55，标题 1.18
LEAD_BODY = 1.55
LEAD_TITLE = 1.18
