"""设计主题：配色、字体、字号。

三个主题都遵循同一套网格：16:9 / 左右 0.95in 边距 / 标题区 0.85in 起。
换主题 = 换配色和字体，版式代码不必改。
"""
from dataclasses import dataclass, replace

from pptx.dml.color import RGBColor

# 幻灯片尺寸（英寸）
SLIDE_W = 13.333
SLIDE_H = 7.5

# 版心网格
MARGIN_X = 0.95
TITLE_TOP = 0.80
CONTENT_TOP = 2.02
CONTENT_W = SLIDE_W - 2 * MARGIN_X          # 11.433
CONTENT_H = 4.42                           # 2.02 → 6.44
FOOTER_TOP = 6.72

# 字体：思源黑体/思源宋体/Inter（scripts/install-fonts.py 安装；缺失时会回退系统字体）
CJK_HEI = "Noto Sans SC"
CJK_HEI_LIGHT = "Noto Sans SC"
CJK_SONG = "Noto Serif SC"
LATIN = "Inter"


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


THEMES = {
    # 论文答辩 / 学术汇报：深蓝 + 金，克制、信息密度高
    "academic": _theme(
        key="academic",
        label="学术答辩",
        bg=rgb("FFFFFF"),
        ink=rgb("14213D"),
        muted=rgb("5C6B82"),
        accent=rgb("1B3A6B"),
        accent_soft=rgb("E9EEF7"),
        rule=rgb("D5DCE8"),
        font_title=CJK_HEI,
        font_body=CJK_HEI,
        font_display=CJK_HEI_LIGHT,
        title_size=31,
        body_size=19,
    ),
    # 杂志 / 设计感 / 线上分享：米白 + 克莱因蓝 + 衬线大标题
    "editorial": _theme(
        key="editorial",
        label="杂志编辑风",
        bg=rgb("F7F3EA"),
        ink=rgb("141414"),
        muted=rgb("6E6A63"),
        accent=rgb("1F3BFF"),
        accent_soft=rgb("E7EBFF"),
        rule=rgb("D6CFC0"),
        font_title=CJK_SONG,
        font_body=CJK_HEI,
        font_display=CJK_SONG,
        title_size=33,
        body_size=19,
        kicker_size=10,
    ),
    # 商务 / 产品 / 路演：纯白 + 一点橙，留白最多
    "minimal": _theme(
        key="minimal",
        label="商务简约",
        bg=rgb("FFFFFF"),
        ink=rgb("111111"),
        muted=rgb("8A8A8A"),
        accent=rgb("E4572E"),
        accent_soft=rgb("FDEDE8"),
        rule=rgb("E6E6E6"),
        font_title=CJK_HEI,
        font_body=CJK_HEI,
        font_display=CJK_HEI_LIGHT,
        title_size=30,
        body_size=18,
    ),
}

DEFAULT_THEME = "academic"


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


def clone_with(theme: Theme, **changes) -> Theme:
    """复制主题并替换若干字段（Theme 是 frozen dataclass，用 replace 生成副本）。

    典型用法：风格档案（profile.apply）在某个内置主题上覆盖配色 / 字体 / 标题字号，
    版式代码里按 theme.key 分支的逻辑保持不变。
    """
    return replace(theme, **changes)
