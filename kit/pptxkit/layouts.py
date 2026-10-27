"""11 种版式。所有版式共用同一套网格，只换主题配色。"""
from __future__ import annotations

import math

from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from .primitives import (add_para, bg, hline, notes, picture, pick_size, rect,
                         slide_number, table, textbox, vline, weighted_len)
from .themes import (CONTENT_H, CONTENT_W, MARGIN_X, SLIDE_H, SLIDE_W,
                     TITLE_TOP, Theme, rgb)

CENTER = PP_ALIGN.CENTER
RIGHT = PP_ALIGN.RIGHT


# ---------------------------------------------------------------- 公共零件

def _new_slide(prs, theme: Theme):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg(slide, theme)
    return slide


def _kicker(slide, theme: Theme, text: str, top: float = 0.92, left: float = MARGIN_X):
    if not text:
        return
    tf = textbox(slide, (left, top, CONTENT_W, 0.24))
    add_para(tf, theme, text.upper() if text.isascii() else text, first=True,
             size=theme.kicker_size, color=theme.accent, bold=True, spacing=1.6)


def _title(slide, theme: Theme, data: dict) -> float:
    """画标题 + 短横线，返回正文起始 y。"""
    title = data.get("title", "").strip()
    if not title:
        return 1.55
    _kicker(slide, theme, data.get("kicker", ""))
    top = 1.20 if data.get("kicker") else TITLE_TOP + 0.18
    length = weighted_len(title)
    size = theme.title_size if length <= 26 else (27 if length <= 36 else 24)
    tf = textbox(slide, (MARGIN_X, top, CONTENT_W, 1.1))
    add_para(tf, theme, title, first=True, size=size, color=theme.ink, bold=True,
             name=theme.font_title, line=1.18)
    lines = 1 if length <= 24 else 2
    rule_y = top + (0.56 if lines == 1 else 0.98)
    hline(slide, MARGIN_X, rule_y, 0.86, theme.accent, 2.4)
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
    title = data.get("title", "")
    subtitle = data.get("subtitle", "")
    kicker = data.get("kicker", "")
    meta = data.get("meta") or {}
    meta_lines = [meta.get(k) for k in ("author", "org", "date") if meta.get(k)]

    if theme.key == "editorial":
        # 右上角线框圆 + 顶部细线，杂志感
        rect(slide, (9.55, -1.85, 4.5, 4.5), fill=None, line=theme.accent, line_w=1.0,
             shape=MSO_SHAPE.OVAL)
        hline(slide, MARGIN_X, 1.30, CONTENT_W, theme.rule, 0.75)
        top = 1.95
    elif theme.key == "academic":
        vline(slide, 0, 0, SLIDE_H, theme.accent, 3.0)
        top = 2.20
    else:
        hline(slide, MARGIN_X, 1.85, CONTENT_W, theme.rule, 0.75)
        top = 2.30

    if kicker:
        tf = textbox(slide, (MARGIN_X, top - 0.52, CONTENT_W, 0.3))
        add_para(tf, theme, kicker, first=True, size=theme.kicker_size + 1,
                 color=theme.accent, bold=True, spacing=2.2)

    length = weighted_len(title)
    size = 44 if length <= 24 else (38 if length <= 32 else 33)
    per_line = max(1.0, (CONTENT_W - 0.35) * 72 / (size * 1.02))
    lines = max(1, math.ceil(length / per_line))
    title_h = lines * size * 1.16 / 72
    tf = textbox(slide, (MARGIN_X, top, CONTENT_W - 0.3, title_h + 0.1))
    add_para(tf, theme, title, first=True, size=size, bold=True, color=theme.ink,
             name=theme.font_title, line=1.16)

    rule_y = top + title_h + 0.26
    hline(slide, MARGIN_X, rule_y, 1.05, theme.accent, 2.6)
    if subtitle:
        tf = textbox(slide, (MARGIN_X, rule_y + 0.30, 9.6, 0.9))
        add_para(tf, theme, subtitle, first=True, size=18, color=theme.muted, line=1.4)

    if meta_lines:
        tf = textbox(slide, (MARGIN_X, 5.95, 9.0, 0.9))
        for i, line in enumerate(meta_lines):
            add_para(tf, theme, line, first=(i == 0), size=12.5,
                     color=theme.muted if i else theme.ink,
                     bold=(i == 0), line=1.4, space_before=0 if i == 0 else 2)


def section(slide, theme: Theme, data: dict):
    number = str(data.get("number") or data.get("kicker") or "").strip()
    title = data.get("title", "")
    subtitle = data.get("subtitle", "")

    # 背景大数字，做层次（浅一档的颜色，不抢标题）
    if number:
        color = theme.accent_soft if theme.key == "editorial" else theme.rule
        if theme.key == "editorial":
            x, y, size = 8.30, 0.25, 235
        elif theme.key == "academic":
            x, y, size = 8.80, 0.45, 210
        else:
            x, y, size = 9.05, 0.70, 180
        tf = textbox(slide, (x, y, 4.4, 3.4))
        add_para(tf, theme, number, first=True, size=size,
                 bold=(theme.key != "academic"), color=color,
                 name=theme.font_display, line=0.95)

    vline(slide, MARGIN_X, 2.02, 2.32, theme.accent, 3.0)
    if number:
        tf = textbox(slide, (MARGIN_X + 0.32, 1.99, 6.4, 0.42))
        add_para(tf, theme, f"PART {number}" if number.isdigit() else number,
                 first=True, size=11, color=theme.accent, bold=True, spacing=2.2)
    tf = textbox(slide, (MARGIN_X + 0.32, 2.46, 7.5, 1.5))
    add_para(tf, theme, title, first=True, size=38, bold=True, color=theme.ink,
             name=theme.font_title, line=1.2)
    if subtitle:
        tf = textbox(slide, (MARGIN_X + 0.32, 4.35, 8.2, 0.8))
        add_para(tf, theme, subtitle, first=True, size=15, color=theme.muted, line=1.5)


def bullets(slide, theme: Theme, data: dict):
    top = _title(slide, theme, data)
    items = data.get("bullets") or []
    char = "▪" if theme.key == "academic" else "•"
    _bullets(slide, theme, items, (MARGIN_X, top, CONTENT_W, 6.30 - top),
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

    if items:
        img_box = (MARGIN_X, top, 7.35, height)
        text_x = MARGIN_X + 7.35 + 0.55
        text_w = SLIDE_W - text_x - MARGIN_X
        _bullets(slide, theme, items, (text_x, top + 0.05, text_w, height),
                 indent=0.30, gap=0.26)
    else:
        img_box = (MARGIN_X, top, CONTENT_W, height)

    picture(slide, spec["path"], img_box, mode="contain", border=theme.rule)
    if cap:
        tf = textbox(slide, (MARGIN_X, img_box[1] + img_box[3] + 0.10, CONTENT_W, 0.28))
        add_para(tf, theme, cap, first=True, size=10.5, color=theme.muted, line=1.0,
                 align=CENTER if not items else PP_ALIGN.LEFT)


def quote(slide, theme: Theme, data: dict):
    spec = data.get("quote") or {}
    text = spec.get("text") or data.get("title", "")
    cite = spec.get("cite", "")

    tf = textbox(slide, (MARGIN_X - 0.06, 1.28, 2.0, 1.5))
    add_para(tf, theme, "“", first=True, size=110, bold=True, color=theme.accent_soft,
             name=theme.font_display, line=0.9)

    size = 30 if weighted_len(text) <= 60 else 26
    tf = textbox(slide, (MARGIN_X + 0.12, 2.42, CONTENT_W - 0.6, 2.9))
    add_para(tf, theme, text, first=True, size=size, color=theme.ink,
             name=theme.font_title if theme.key == "editorial" else None,
             line=1.55, bold=(theme.key == "editorial"))
    if cite:
        hline(slide, MARGIN_X + 0.12, 5.52, 0.72, theme.accent, 2.0)
        tf = textbox(slide, (MARGIN_X + 0.12, 5.68, CONTENT_W - 1.0, 0.5))
        add_para(tf, theme, cite, first=True, size=13, color=theme.muted, line=1.2)


def stats(slide, theme: Theme, data: dict):
    top = _title(slide, theme, data)
    items = data.get("stats") or []
    if not items:
        return
    n = len(items)
    gap = 0.55
    w = (CONTENT_W - gap * (n - 1)) / n
    y = top + 0.30
    for i, item in enumerate(items):
        x = MARGIN_X + i * (w + gap)
        hline(slide, x, y, 0.66, theme.accent, 2.6)
        value = str(item.get("value", ""))
        vsize = 54 if weighted_len(value) <= 4 else (42 if weighted_len(value) <= 7 else 34)
        tf = textbox(slide, (x, y + 0.30, w, vsize / 72 * 1.4))
        add_para(tf, theme, value, first=True, size=vsize, bold=True,
                 color=theme.accent, name=theme.font_display, line=1.05)
        tf = textbox(slide, (x, y + 0.30 + vsize / 72 * 1.35 + 0.10, w, 0.9))
        add_para(tf, theme, str(item.get("label", "")), first=True, size=14,
                 color=theme.ink, line=1.45)
        if item.get("note"):
            tf = textbox(slide, (x, y + 0.30 + vsize / 72 * 1.35 + 0.62, w, 0.9))
            add_para(tf, theme, str(item["note"]), first=True, size=11,
                     color=theme.muted, line=1.5)


def timeline(slide, theme: Theme, data: dict):
    top = _title(slide, theme, data)
    steps = data.get("steps") or []
    if not steps:
        return
    n = len(steps)
    gap = 0.34
    w = (CONTENT_W - gap * (n - 1)) / n
    line_y = top + 1.05
    hline(slide, MARGIN_X, line_y, CONTENT_W, theme.rule, 1.0)
    for i, step in enumerate(steps):
        x = MARGIN_X + i * (w + gap)
        rect(slide, (x, line_y - 0.075, 0.15, 0.15), fill=theme.accent,
             shape=MSO_SHAPE.OVAL)
        tf = textbox(slide, (x, line_y + 0.24, w - 0.10, 0.36))
        add_para(tf, theme, f"0{i + 1}", first=True, size=11, bold=True,
                 color=theme.accent, spacing=1.2)
        tf = textbox(slide, (x, line_y + 0.58, w - 0.10, 0.9))
        add_para(tf, theme, str(step.get("name", "")), first=True, size=16,
                 bold=True, color=theme.ink, line=1.3)
        if step.get("desc"):
            tf = textbox(slide, (x, line_y + 1.32, w - 0.10, 1.6))
            add_para(tf, theme, str(step["desc"]), first=True, size=11.5,
                     color=theme.muted, line=1.45)


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
