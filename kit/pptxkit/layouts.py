"""11 种版式。所有版式共用同一套网格，只换主题配色。

cover / section / stats / quote / timeline 五个版式直接引用 design.py 的
12 栏网格（col_x / col_w / row_y）与 1.25 字号阶梯（T_H1 / T_H2 / T_BODY …）；
其余版式保持原有实现，只共享 themes 的 MARGIN_X / CONTENT_W 常量。
"""
from __future__ import annotations

import math

from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from . import design as D
from .primitives import (add_para, bg, hline, notes, picture, pick_size, rect,
                         slide_number, table, textbox, vline, weighted_len)
from .themes import (CONTENT_W, MARGIN_X, SLIDE_H, SLIDE_W, TITLE_TOP, Theme,
                     title_scale)

CENTER = PP_ALIGN.CENTER
RIGHT = PP_ALIGN.RIGHT

# ---------------------------------------------------------------- 设计系统常量
# 全部由 design.py 推导：纵向节奏 = 栏距，短线 = 一格栏宽，圆点约 0.43 栏距。
GAP = D.COL_GUTTER                 # 0.28in：栏距，同时也是版式的纵向节奏
HAIRLINE = D.col_w(1)              # 0.688in：一格栏宽的装饰短线
RULE_L = D.col_w(1) * 1.5          # 1.03in：封面标题下的强调短线
DOT = D.COL_GUTTER * 0.43          # 0.12in：时间线圆点（视觉上约 8.6pt）
PANEL_COL = 8                      # editorial 封面右侧色块从第 8 栏起（留 8 栏给文字）


def _title_pt(theme: Theme, step: int) -> float:
    """阶梯字号 × 主题标题比例（风格档案的 title_scale 由此生效）。"""
    return round(D.type_size(step) * title_scale(theme), 1)


def _skeleton(data: dict) -> dict | None:
    """builder 注入的版式骨架（profile.layout_hints() 的结果）。"""
    hints = data.get("_skeleton") if isinstance(data, dict) else None
    return hints if isinstance(hints, dict) else None


def _geometry(data: dict) -> tuple[float, float, float]:
    """(标题纵向位移, 左边距位移, 右边距位移)，单位英寸；无骨架时全为 0。"""
    hints = _skeleton(data)
    if not hints:
        return 0.0, 0.0, 0.0
    try:
        return (float(hints.get("title_shift") or 0.0),
                float(hints.get("margin_left") or 0.0),
                float(hints.get("margin_right") or 0.0))
    except (TypeError, ValueError):
        return 0.0, 0.0, 0.0


# ---------------------------------------------------------------- 公共零件

def _new_slide(prs, theme: Theme):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg(slide, theme)
    return slide


def _kicker(slide, theme: Theme, text: str, top: float = 0.92, left: float = MARGIN_X,
            width: float = CONTENT_W):
    if not text:
        return
    tf = textbox(slide, (left, top, width, 0.24))
    add_para(tf, theme, text.upper() if text.isascii() else text, first=True,
             size=theme.kicker_size, color=theme.accent, bold=True, spacing=1.6)


def _title(slide, theme: Theme, data: dict) -> float:
    """画标题 + 短横线，返回正文起始 y。可按版式骨架微调标题纵向位置与左右边距。"""
    title = data.get("title", "").strip()
    if not title:
        return 1.55
    dy, dxl, dxr = _geometry(data)
    left = MARGIN_X + dxl
    width = CONTENT_W - dxl - dxr
    base_top = 1.20 if data.get("kicker") else TITLE_TOP + 0.18
    top = base_top + dy
    if dy:
        top = min(max(top, 0.55), 2.20)      # 正文页标题带的安全区间
    if data.get("kicker"):
        _kicker(slide, theme, data["kicker"], top=top - 0.28, left=left, width=width)
    length = weighted_len(title)
    size = theme.title_size if length <= 26 else (27 if length <= 36 else 24)
    tf = textbox(slide, (left, top, width, 1.1))
    add_para(tf, theme, title, first=True, size=size, color=theme.ink, bold=True,
             name=theme.font_title, line=1.18)
    lines = 1 if length <= 24 else 2
    rule_y = top + (0.56 if lines == 1 else 0.98)
    hline(slide, left, rule_y, 0.86, theme.accent, 2.4)
    body_top = rule_y + 0.30
    return body_top


def _bullets(slide, theme: Theme, items, box, *, indent=0.34, size=None,
             line=1.42, gap=0.30, level2_scale=0.88, bullet_char="•",
             anchor=MSO_ANCHOR.TOP):
    """items: [{'text':..., 'level':0}]"""
    left, top, width, height = box
    texts = [i["text"] for i in items if i.get("text")]
    base = size or pick_size(texts, width, height, line=line, gap=gap, indent=indent)
    tf = textbox(slide, (left, top, width, height), anchor=anchor)
    for idx, item in enumerate(items):
        level = min(int(item.get("level", 0)), 2)
        s = base if level == 0 else max(12, int(base * level2_scale))
        add_para(
            tf, theme, item["text"], first=(idx == 0),
            size=s,
            color=theme.ink if level == 0 else theme.muted,
            bold=(level == 0 and item.get("bold", False)),
            line=line,
            space_before=0 if idx == 0 else gap * s,
            bullet=True,
            bullet_char=bullet_char if level == 0 else "–",
            indent=indent + level * 0.34,
            hanging=0.30,
        )
    return tf


def _two_columns(slide, theme: Theme, data: dict, top: float, *, divider=True,
                 col_gap=0.66):
    cols = [data.get("left") or {}, data.get("right") or {}]
    total_w = CONTENT_W
    col_w = (total_w - col_gap) / 2
    x = [MARGIN_X, MARGIN_X + col_w + col_gap]
    height = 6.30 - top
    for i, col in enumerate(cols):
        y = top
        if col.get("title"):
            tf = textbox(slide, (x[i], y, col_w, 0.34))
            add_para(tf, theme, col["title"], first=True, size=15, bold=True,
                     color=theme.accent, line=1.0)
            y += 0.46
        items = col.get("bullets") or []
        if items:
            _bullets(slide, theme, items, (x[i], y, col_w, height - (y - top)),
                     indent=0.30, gap=0.24, bullet_char="•")
    if divider:
        vline(slide, MARGIN_X + col_w + col_gap / 2, top + 0.05, height - 0.35,
              theme.rule, 0.75)


# ---------------------------------------------------------------- 版式

def cover(slide, theme: Theme, data: dict):
    """封面：文字按 12 栏排（editorial 让出右侧第 8–12 栏给色块），行位置走网格行。"""
    title = data.get("title", "")
    subtitle = data.get("subtitle", "")
    kicker = data.get("kicker", "")
    meta = data.get("meta") or {}
    meta_lines = [meta.get(k) for k in ("author", "org", "date") if meta.get(k)]
    dy, dxl, dxr = _geometry(data)

    # 文字栏：editorial 给右侧色块留出第 8–12 栏；其余主题用满 12 栏
    text_span = PANEL_COL if theme.key == "editorial" else D.COLS
    text_x = D.col_x(0) + dxl
    text_w = D.col_w(text_span) - dxl - dxr
    side_w = D.col_w(6) - dxl - dxr

    # 主题装饰（都是网格对齐的几何元素，不抢标题）
    if theme.key == "editorial":
        panel_x = D.col_x(PANEL_COL)
        rect(slide, (panel_x, 0, D.SLIDE_W - panel_x, SLIDE_H), fill=theme.accent_soft)
        vline(slide, panel_x, 0, SLIDE_H, theme.accent, 2.0)
    elif theme.key == "academic":
        vline(slide, 0, 0, SLIDE_H, theme.accent, 3.0)
    else:
        hline(slide, D.col_x(0), D.row_y(0.45), D.content_w(), theme.rule, 0.75)

    title_y = D.row_y(1.30) + dy
    if dy:
        title_y = min(max(title_y, 0.90), 3.00)   # 封面标题带的安全区间

    if kicker:
        tf = textbox(slide, (text_x, title_y - 0.40, text_w,
                             D.lines_height(D.T_CAPTION, 1, D.LEAD_BODY)))
        add_para(tf, theme, kicker, first=True, size=D.T_CAPTION,
                 color=theme.accent, bold=True, spacing=2.2)

    length = weighted_len(title)
    size = _title_pt(theme, 4) if length <= 24 else _title_pt(theme, 3)
    per_line = max(1.0, text_w * 72 / (size * 1.02))
    lines = max(1, math.ceil(length / per_line))
    title_h = D.lines_height(size, lines, D.LEAD_TITLE)
    tf = textbox(slide, (text_x, title_y, text_w, title_h + GAP))
    add_para(tf, theme, title, first=True, size=size, bold=True, color=theme.ink,
             name=theme.font_title, line=D.LEAD_TITLE)

    # 标题区按「至少两行」预留高度：短标题也有足够的呼吸，长标题自然向下长
    block_h = D.lines_height(size, max(2, lines), D.LEAD_TITLE)
    rule_y = title_y + block_h + GAP
    if dy and subtitle and rule_y > 4.90:      # 向下微调时别压住副标题/署名
        rule_y = 4.90
    hline(slide, text_x, rule_y, RULE_L, theme.accent, 2.6)
    if subtitle:
        tf = textbox(slide, (text_x, rule_y + GAP, side_w,
                             D.lines_height(D.T_BODY, 2, D.LEAD_BODY)))
        add_para(tf, theme, subtitle, first=True, size=D.T_BODY, color=theme.muted,
                 line=D.LEAD_BODY)

    if meta_lines:
        tf = textbox(slide, (text_x, D.row_y(5.1), side_w,
                             D.lines_height(D.T_SMALL, 3, 1.4)))
        for i, line in enumerate(meta_lines):
            add_para(tf, theme, line, first=(i == 0), size=D.T_SMALL,
                     color=theme.muted if i else theme.ink,
                     bold=(i == 0), line=1.4, space_before=0 if i == 0 else 2)


def section(slide, theme: Theme, data: dict):
    """章节页：背景大编号右对齐到版心右缘，左侧 3pt 竖条 + PART 标签 + 标题。"""
    number = str(data.get("number") or data.get("kicker") or "").strip()
    title = data.get("title", "")
    subtitle = data.get("subtitle", "")
    dy, dxl, dxr = _geometry(data)

    # 背景大编号：右对齐到版心右缘（文本框给足宽度，绝不折行）
    if number:
        ghost_span = 8 if len(number) <= 2 else 10
        ghost_color = theme.accent_soft if theme.key == "editorial" else theme.rule
        tf = textbox(slide, (D.col_x(D.COLS - ghost_span), D.row_y(-0.55),
                             D.col_w(ghost_span), D.lines_height(D.T_GHOST, 1, 0.95)))
        add_para(tf, theme, number, first=True, size=D.T_GHOST,
                 bold=(theme.key != "academic"), color=ghost_color,
                 name=theme.font_display, line=0.95, align=RIGHT)

    title_y = D.row_y(1.60) + dy
    if dy:
        title_y = min(max(title_y, 1.15), 4.00)   # 章节标题带的安全区间
    # 左侧竖条：与 PART 标签同高，长度固定 2.25in
    bar_top = D.row_y(1.15) if not dy else title_y - 0.45
    vline(slide, D.col_x(0), bar_top, 2.25, theme.accent, 3.0)
    text_x = D.col_x(0) + GAP + dxl
    if number:
        part_y = D.row_y(1.14) if not dy else title_y - 0.46
        tf = textbox(slide, (text_x, part_y, D.col_w(6) - dxl - dxr,
                             D.lines_height(D.T_CAPTION, 1, 1.2)))
        add_para(tf, theme, f"PART {number}" if number.isdigit() else number,
                 first=True, size=D.T_CAPTION, color=theme.accent, bold=True, spacing=2.2)

    size = _title_pt(theme, 3)
    title_w = D.col_w(8) - dxl - dxr
    per_line = max(1.0, title_w * 72 / (size * 1.02))
    title_h = D.lines_height(size, max(1, math.ceil(weighted_len(title) / per_line)), 1.2)
    tf = textbox(slide, (text_x, title_y, title_w, title_h + GAP))
    add_para(tf, theme, title, first=True, size=size, bold=True, color=theme.ink,
             name=theme.font_title, line=1.2)

    if subtitle:
        sub_y = max(D.row_y(3.50) + dy, title_y + title_h + GAP * 1.5)
        if dy:
            sub_y = min(sub_y, 5.20)
        tf = textbox(slide, (text_x, sub_y, title_w,
                             D.lines_height(D.T_BODY, 2, D.LEAD_BODY)))
        add_para(tf, theme, subtitle, first=True, size=D.T_BODY, color=theme.muted,
                 line=D.LEAD_BODY)


def bullets(slide, theme: Theme, data: dict):
    top = _title(slide, theme, data)
    items = data.get("bullets") or []
    char = "▪" if theme.key == "academic" else "•"
    dy, dxl, dxr = _geometry(data)
    _bullets(slide, theme, items,
             (MARGIN_X + dxl, top, CONTENT_W - dxl - dxr, 6.30 - top),
             bullet_char=char, anchor=MSO_ANCHOR.MIDDLE)


def two_col(slide, theme: Theme, data: dict):
    top = _title(slide, theme, data)
    _two_columns(slide, theme, data, top + 0.02)


def table_layout(slide, theme: Theme, data: dict):
    top = _title(slide, theme, data)
    spec = data.get("table") or {}
    header = spec.get("header") or []
    rows = spec.get("rows") or []
    if not header:
        return
    n = len(rows)
    avail = 6.25 - top
    row_h = min(0.62, max(0.34, (avail - 0.55) / max(1, n)))
    size = 14 if len(header) <= 4 else (13 if len(header) <= 6 else 12)
    table(slide, theme, (MARGIN_X, top, CONTENT_W, avail), header, rows,
          col_widths=spec.get("widths"), size=size, header_size=size, row_h=row_h)


def image(slide, theme: Theme, data: dict):
    top = _title(slide, theme, data)
    spec = data.get("image") or {}
    items = data.get("bullets") or []
    cap = spec.get("caption", "")
    height = 6.28 - top - (0.34 if cap else 0)
    dy, dxl, dxr = _geometry(data)
    left = MARGIN_X + dxl

    if items:
        img_box = (left, top, 7.35, height)
        text_x = left + 7.35 + 0.55
        text_w = SLIDE_W - text_x - (MARGIN_X + dxr)
        _bullets(slide, theme, items, (text_x, top + 0.05, text_w, height),
                 indent=0.30, gap=0.26)
    else:
        img_box = (left, top, CONTENT_W - dxl - dxr, height)

    picture(slide, spec["path"], img_box, mode="contain", border=theme.rule)
    if cap:
        tf = textbox(slide, (left, img_box[1] + img_box[3] + 0.10,
                             CONTENT_W - dxl - dxr, 0.28))
        add_para(tf, theme, cap, first=True, size=10.5, color=theme.muted, line=1.0,
                 align=CENTER if not items else PP_ALIGN.LEFT)


def quote(slide, theme: Theme, data: dict):
    """引用页：装饰引号 + 引文 + 出处，全部落在 12 栏网格与字号阶梯上。"""
    spec = data.get("quote") or {}
    text = spec.get("text") or data.get("title", "")
    cite = spec.get("cite", "")

    tf = textbox(slide, (D.col_x(0), D.row_y(0.40), D.col_w(2),
                         D.lines_height(D.T_DECOR, 1, 0.9)))
    add_para(tf, theme, "“", first=True, size=D.T_DECOR, bold=True,
             color=theme.accent_soft, name=theme.font_display, line=0.9)

    # 引文：在 T_H4–T_H2 里挑「一行放得下」的最大字号（用真实字宽量，避免孤行）
    body_x, body_w = D.col_x(0), D.col_w(12)
    quote_font = theme.font_title if theme.key == "editorial" else theme.font_body
    size = _title_pt(theme, 1)
    for step in (3, 2):
        cand = _title_pt(theme, step)
        if D.fits_one_line(text, quote_font, cand, body_w):
            size = cand
            break
    body_y = D.row_y(1.60)
    tf = textbox(slide, (body_x, body_y, body_w, D.row_y(4.45) - body_y))
    add_para(tf, theme, text, first=True, size=size, color=theme.ink,
             name=theme.font_title if theme.key == "editorial" else None,
             line=D.LEAD_BODY, bold=(theme.key == "editorial"))
    if cite:
        hline(slide, body_x, D.row_y(4.70), HAIRLINE, theme.accent, 2.0)
        tf = textbox(slide, (body_x, D.row_y(4.86), D.col_w(8),
                             D.lines_height(D.T_SMALL, 1, 1.2)))
        add_para(tf, theme, cite, first=True, size=D.T_SMALL, color=theme.muted, line=1.2)


def stats(slide, theme: Theme, data: dict):
    """数据卡：每个数字占 12/n 栏，栏宽由 col_w 决定，字号走 T_DISPLAY→T_H2。"""
    top = _title(slide, theme, data)
    items = data.get("stats") or []
    if not items:
        return
    n = len(items)
    span = max(2, D.COLS // n)
    y = top + GAP
    for i, item in enumerate(items):
        x = D.col_x(i * span)
        w = D.col_w(span)
        hline(slide, x, y, HAIRLINE, theme.accent, 2.2)

        value = str(item.get("value", ""))
        vlen = weighted_len(value)
        vsize = D.T_DISPLAY if vlen <= 4 else (D.T_H1 if vlen <= 7 else D.T_H2)
        vh = D.lines_height(vsize, 1, 1.05)
        tf = textbox(slide, (x, y + GAP, w, vh + GAP))
        add_para(tf, theme, value, first=True, size=vsize, bold=True,
                 color=theme.accent, name=theme.font_display, line=1.05)

        label_y = y + GAP + vh + GAP
        label_h = D.lines_height(D.T_H4, 1, 1.3)
        tf = textbox(slide, (x, label_y, w, label_h))
        add_para(tf, theme, str(item.get("label", "")), first=True, size=D.T_H4,
                 color=theme.ink, line=1.3)

        if item.get("note"):
            tf = textbox(slide, (x, label_y + label_h + GAP / 2, w,
                                 D.lines_height(D.T_SMALL, 3, 1.45)))
            add_para(tf, theme, str(item["note"]), first=True, size=D.T_SMALL,
                     color=theme.muted, line=1.45)


def _timeline_desc(desc) -> str:
    """时间线说明：Markdown 里 `月份 | 阶段 | 说明` 三段写法只取到前两段，
    剩下以 "|" 开头的部分用冒号接上，避免把分隔符画到页面上。"""
    parts = [p.strip() for p in str(desc or "").split("|") if p.strip()]
    if not parts:
        return ""
    return parts[0] + ("：" + "，".join(parts[1:]) if len(parts) > 1 else "")


def timeline(slide, theme: Theme, data: dict):
    """时间线：每个阶段占 12/n 栏，圆点、序号、名称、说明都贴着网格行。"""
    top = _title(slide, theme, data)
    steps = data.get("steps") or []
    if not steps:
        return
    n = len(steps)
    span = max(2, D.COLS // n)
    line_y = max(D.row_y(2.05), top + GAP * 3)     # 标题换行时整条线一起下移
    hline(slide, D.col_x(0), line_y, D.content_w(), theme.rule, 1.0)
    for i, step in enumerate(steps):
        x = D.col_x(i * span)
        w = D.col_w(span)
        rect(slide, (x, line_y - DOT / 2, DOT, DOT), fill=theme.accent,
             shape=MSO_SHAPE.OVAL)

        num_y = line_y + GAP
        num_h = D.lines_height(D.T_CAPTION, 1, 1.2)
        tf = textbox(slide, (x, num_y, w, num_h))
        add_para(tf, theme, f"0{i + 1}", first=True, size=D.T_CAPTION, bold=True,
                 color=theme.accent, spacing=1.2)

        name_y = num_y + num_h + GAP / 2
        name_h = D.lines_height(D.T_H4, 1, 1.3)
        tf = textbox(slide, (x, name_y, w, name_h))
        add_para(tf, theme, str(step.get("name", "")), first=True, size=D.T_H4,
                 bold=True, color=theme.ink, line=1.3)

        if step.get("desc"):
            tf = textbox(slide, (x, name_y + name_h + GAP / 2, w - GAP / 2,
                                 D.lines_height(D.T_SMALL, 3, 1.45)))
            add_para(tf, theme, _timeline_desc(step["desc"]), first=True,
                     size=D.T_SMALL, color=theme.muted, line=1.45)


def refs(slide, theme: Theme, data: dict):
    top = _title(slide, theme, data)
    items = data.get("refs") or []
    if not items:
        items = [b["text"] for b in (data.get("bullets") or [])]
    two = len(items) > 5
    col_w = (CONTENT_W - 0.7) / 2 if two else CONTENT_W
    size = 11.5 if two else 12.5
    for idx, text in enumerate(items):
        col = 0 if not two else idx // ((len(items) + 1) // 2)
        row = idx if not two else idx % ((len(items) + 1) // 2)
        x = MARGIN_X + col * (col_w + 0.7)
        y = top + row * 0.72
        tf = textbox(slide, (x, y, col_w, 0.7))
        add_para(tf, theme, f"[{idx + 1}]  {text}", first=True, size=size,
                 color=theme.ink, line=1.35, indent=0.34, hanging=0.34)


def end(slide, theme: Theme, data: dict):
    title = data.get("title", "谢谢")
    subtitle = data.get("subtitle", "")
    contact = data.get("contact", "")

    hline(slide, SLIDE_W / 2 - 0.42, 2.30, 0.84, theme.accent, 2.6)
    tf = textbox(slide, (1.2, 2.62, SLIDE_W - 2.4, 1.5))
    add_para(tf, theme, title, first=True, size=54, bold=True, color=theme.ink,
             name=theme.font_title, line=1.1, align=CENTER)
    if subtitle:
        tf = textbox(slide, (1.2, 4.15, SLIDE_W - 2.4, 0.9))
        add_para(tf, theme, subtitle, first=True, size=18, color=theme.muted,
                 line=1.4, align=CENTER)
    if contact:
        tf = textbox(slide, (1.2, 5.35, SLIDE_W - 2.4, 0.6))
        add_para(tf, theme, contact, first=True, size=12, color=theme.muted,
                 line=1.2, align=CENTER, spacing=0.6)


LAYOUTS = {
    "cover": cover,
    "section": section,
    "bullets": bullets,
    "two-col": two_col,
    "table": table_layout,
    "image": image,
    "quote": quote,
    "stats": stats,
    "timeline": timeline,
    "refs": refs,
    "end": end,
}
ALIASES = {
    "title": "cover", "封面": "cover", "章节": "section", "两栏": "two-col",
    "两列": "two-col", "表格": "table", "图片": "image", "引用": "quote",
    "数据": "stats", "数据卡": "stats", "流程": "timeline", "时间线": "timeline",
    "参考文献": "refs", "致谢": "end", "结尾": "end", "要点": "bullets",
}


def render(prs, theme: Theme, data: dict, index: int, total: int, footer: str = ""):
    slide = _new_slide(prs, theme)
    name = (data.get("layout") or "bullets").strip().lower()
    name = ALIASES.get(name, name)
    fn = LAYOUTS.get(name, bullets)
    for attempt in (name, "bullets"):
        try:
            LAYOUTS.get(attempt, bullets)(slide, theme, data)
            break
        except TypeError:
            raise
        except Exception as exc:  # 单页出错不影响整份文件
            if attempt == "bullets":
                raise
            print(f"  ! 版式 {name} 失败（{exc}），退回 bullets")
    if name not in ("cover", "end"):
        slide_number(slide, theme, index, total, data.get("footer", footer))
    notes(slide, data.get("note", ""))
    return slide
