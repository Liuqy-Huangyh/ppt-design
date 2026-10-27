"""设计主题：配色、字体、字号。

三套主题共用一套网格与字号阶梯（都来自 kit/pptxkit/design.py）：
16:9 / 左右 1in 边距 / 12 栏 0.28in 栏距 / 1.25 字号阶梯。
配色也由 design.palette() 从单一强调色按 OKLCH 推导（academic/minimal 用
clean、editorial 用 paper），不再手写六组色值。

换主题 = 换配色和字体，版式代码不必改。
"""
from dataclasses import dataclass, replace

from pptx.dml.color import RGBColor

from . import design as D

# 幻灯片尺寸（英寸）
SLIDE_W = D.SLIDE_W
SLIDE_H = D.SLIDE_H

# 版心网格：与 design 的 12 栏网格一致（MARGIN_X=1.0 / 栏距 0.28）
MARGIN_X = D.MARGIN_X
TITLE_TOP = 0.80
CONTENT_TOP = 2.02
CONTENT_W = D.content_w()                   # 11.333
CONTENT_H = 4.42                           # 2.02 → 6.44
FOOTER_TOP = 6.72

# 字体：思源黑体/思源宋体/Inter（scripts/install-fonts.py 安装；缺失时会回退系统字体）
CJK_HEI = D.SANS
CJK_HEI_LIGHT = D.SANS
CJK_SONG = D.SERIF
LATIN = D.NUM


def rgb(value: str) -> RGBColor:
    value = value.lstrip("#")
    return RGBColor(int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))


@dataclass(frozen=True)
class Theme:
    key: str
    label: str
    bg: RGBColor
    ink: RGBColor          # 主文字
    muted: RGBColor        # 次要文字
    accent: RGBColor       # 强调色
    accent_soft: RGBColor  # 强调色浅底
    rule: RGBColor         # 分隔线
    font_title: str
    font_body: str
    font_display: str      # 大数字 / 引言
    title_size: int = 32
    body_size: int = 19
    kicker_size: int = 11

    @property
    def dark(self) -> bool:
        """深色底主题（决定页脚/分割线对比度）。"""
        r, g, b = self.bg[0], self.bg[1], self.bg[2]
        return (r * 299 + g * 587 + b * 114) / 1000 < 128


def _theme(**kw) -> Theme:
    return Theme(**kw)


def _colors(accent: str, mode: str, **kwargs) -> dict:
    """用 design.palette() 从一个强调色推导六种颜色（含 WCAG 对比度校正）。"""
    pal = D.palette(accent, mode=mode, **kwargs)
    return {
        "bg": rgb(pal.bg),
        "ink": rgb(pal.ink),
        "muted": rgb(pal.muted),
        "accent": rgb(pal.accent),
        "accent_soft": rgb(pal.accent_soft),
        "rule": rgb(pal.rule),
    }


THEMES = {
    # 论文答辩 / 学术汇报：深蓝（#1B3A6B）+ 纯白纸面（clean）
    "academic": _theme(
        key="academic",
        label="学术答辩",
        **_colors("#1B3A6B", "clean"),
        font_title=CJK_HEI,
        font_body=CJK_HEI,
        font_display=CJK_HEI_LIGHT,
        title_size=31,
        body_size=19,
    ),
    # 杂志 / 设计感 / 线上分享：克莱因蓝（#1F3BFF）+ 暖米色纸感（paper + warm）
    "editorial": _theme(
        key="editorial",
        label="杂志编辑风",
        **_colors("#1F3BFF", "paper", warm_paper=True),
        font_title=CJK_SONG,
        font_body=CJK_HEI,
        font_display=CJK_SONG,
        title_size=33,
        body_size=19,
        kicker_size=10,
    ),
    # 商务 / 产品 / 路演：纯白 + 一点橙（#E4572E，clean）
    "minimal": _theme(
        key="minimal",
        label="商务简约",
        **_colors("#E4572E", "clean"),
        font_title=CJK_HEI,
        font_body=CJK_HEI,
        font_display=CJK_HEI_LIGHT,
        title_size=30,
        body_size=18,
    ),
}

DEFAULT_THEME = "academic"

# 每个内置主题的基准标题字号；profile.apply() 的 title_scale 通过 title_size 传入，
# 版式再按「当前 title_size / 基准」等比缩放阶梯字号（默认主题比例恒为 1.0）。
TITLE_BASE = {key: theme.title_size for key, theme in THEMES.items()}


def get_theme(name: str | None) -> Theme:
    if not name:
        return THEMES[DEFAULT_THEME]
    key = name.strip().lower()
    if key in THEMES:
        return THEMES[key]
    alias = {
        "学术": "academic", "答辩": "academic", "论文": "academic",
        "杂志": "editorial", "设计": "editorial", "分享": "editorial",
        "商务": "minimal", "简约": "minimal", "产品": "minimal",
    }
    if key in alias:
        return THEMES[alias[key]]
    raise SystemExit(f"未知主题：{name}（可选：{', '.join(THEMES)}）")


def title_scale(theme: Theme) -> float:
    """当前主题标题字号相对内置基准的比例（风格档案的 title_scale 由此生效）。"""
    base = TITLE_BASE.get(theme.key)
    if not base:
        return 1.0
    return theme.title_size / base


def clone_with(theme: Theme, **changes) -> Theme:
    """复制主题并替换若干字段（Theme 是 frozen dataclass，用 replace 生成副本）。

    典型用法：风格档案（profile.apply）在某个内置主题上覆盖配色 / 字体 / 标题字号，
    版式代码里按 theme.key 分支的逻辑保持不变。
    """
    return replace(theme, **changes)
