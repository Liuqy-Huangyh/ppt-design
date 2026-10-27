"""设计系统演示：用 Noto Sans SC / Noto Serif SC / Inter + OKLCH 配色 + 12 栏网格，
渲染 6 页样张，和旧版式做对比。

    python scripts/design-showcase.py
    → out/design-showcase.pptx（再 render 成 PNG 目检）
"""
from __future__ import annotations

import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "kit"))

from pptxkit import design as D  # noqa: E402

OUT = ROOT / "out" / "design-showcase.pptx"
ACCENT = "#1B3A6B"
GOLD = "#B08423"
PAL = D.palette(ACCENT, mode="paper")
INK, MUTED, RULE, SOFT, BG = PAL.ink, PAL.muted, PAL.rule, PAL.accent_soft, PAL.bg


def hexrgb(h: str) -> RGBColor:
    v = h.lstrip("#")
    return RGBColor(int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16))


def set_font(run, name: str, size: float, color: str = INK, bold: bool = False,
             italic: bool = False, spacing: float | None = None) -> None:
    f = run.font
    f.name = name
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = hexrgb(color)
    rPr = f._element
    latin = rPr.find(qn("a:latin"))
    if latin is None:
        latin = etree.SubElement(rPr, qn("a:latin"))
    latin.set("typeface", name)
    anchor = latin
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.Element(qn(tag))
            anchor.addnext(el)
        el.set("typeface", name)
        anchor = el
    if spacing:
        rPr.set("spc", str(int(spacing * 100)))


def textbox(slide, x, y, w, h, *, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.auto_size = None
    return tf


def para(tf, text: str, *, first=False, font=D.SANS, size=D.T_BODY, color=INK,
         bold=False, italic=False, line=D.LEAD_BODY, space_before=0,
         align=PP_ALIGN.LEFT, spacing=None):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.line_spacing = line
    if space_before:
        p.space_before = Pt(space_before)
    run = p.add_run()
    run.text = text
    set_font(run, font, size, color, bold, italic, spacing)
    return p


def rect(slide, x, y, w, h, fill: str, *, shape=MSO_SHAPE.RECTANGLE):
    shp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = hexrgb(fill)
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def hline(slide, x, y, w, color=RULE, weight=0.9):
    rect(slide, x, y, w, max(weight / 72, 0.006), color)


def slide(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    f = s.background.fill
    f.solid()
    f.fore_color.rgb = hexrgb(BG)
    return s


def page_num(s, idx, total):
    tf = textbox(s, 11.4, 6.93, 0.95, 0.25)
    para(tf, f"{idx:02d} / {total:02d}", first=True, font=D.NUM, size=D.T_CAPTION,
         color=MUTED, line=1.0, align=PP_ALIGN.RIGHT)


# ---------------------------------------------------------------- 页面

def cover(prs):
    s = slide(prs)
    # 右侧出血色块
    rect(s, D.SLIDE_W - 4.55, 0, 4.55, D.SLIDE_H, SOFT)
    rect(s, D.SLIDE_W - 4.55, 0, 0.045, D.SLIDE_H, PAL.accent)

    tf = textbox(s, D.MARGIN_X, 1.55, 7.4, 0.3)
    para(tf, "DESIGN SYSTEM · 2026", first=True, font=D.NUM, size=10.5,
         color=PAL.accent, bold=True, line=1.0, spacing=2.4)

    tf = textbox(s, D.MARGIN_X, 2.02, 7.3, 2.6)
    para(tf, "让版面先于内容", first=True, font=D.SERIF, size=D.T_DISPLAY,
         color=INK, bold=False, line=1.1)
    para(tf, "排版系统演示", font=D.SERIF, size=D.T_DISPLAY, color=INK, line=1.1)

    hline(s, D.MARGIN_X, 4.62, 1.15, GOLD, 2.4)

    tf = textbox(s, D.MARGIN_X, 4.92, 6.9, 0.9)
    para(tf, "一套字体、五种字号、十二栏网格，替代凭感觉挪文本框", first=True,
         font=D.SANS, size=D.T_BODY, color=MUTED, line=1.6)

    tf = textbox(s, D.MARGIN_X, 6.28, 6.9, 0.8)
    para(tf, "Noto Serif SC × Noto Sans SC × Inter", first=True, font=D.NUM,
         size=D.T_CAPTION + 0.5, color=INK, bold=True, line=1.6, spacing=0.6)
    para(tf, "OKLCH 配色 · 1.25 字号阶梯 · 1in 版心边距", font=D.SANS,
         size=D.T_CAPTION, color=MUTED, line=1.6)

    tf = textbox(s, D.SLIDE_W - 3.85, 1.55, 3.4, 0.4)
    para(tf, "PP T  ·  D E S I G N", first=True, font=D.NUM, size=10.5,
         color="#FFFFFF", bold=True, line=1.0, align=PP_ALIGN.RIGHT, spacing=2.6)
    tf = textbox(s, D.SLIDE_W - 3.85, 2.05, 3.4, 0.9)
    para(tf, "2026", first=True, font=D.NUM, size=44, color="#FFFFFF", bold=True,
         line=1.0, align=PP_ALIGN.RIGHT)


def kpi(prs, idx, total):
    s = slide(prs)
    tf = textbox(s, D.MARGIN_X, 0.95, 9.0, 0.3)
    para(tf, "关键指标", first=True, font=D.SANS, size=10.5, color=PAL.accent,
         bold=True, line=1.0, spacing=2.2)
    tf = textbox(s, D.MARGIN_X, 1.28, 9.5, 0.8)
    para(tf, "三条数据，一眼看完", first=True, font=D.SANS, size=D.T_H3,
         color=INK, bold=True, line=1.2)
    hline(s, D.MARGIN_X, 2.12, D.content_w(), RULE)

    items = [
        ("38", "FPS", "边缘端实时帧率", "Jetson Orin · 640×640", "+12%"),
        ("4.2", "×", "参数量差距", "相比同精度云端模型", "−"),
        ("62", "%", "标注成本占比", "数据环节 / 总成本", "+8pt"),
    ]
    span = 4
    for i, (v, unit, label, note, delta) in enumerate(items):
        x = D.col_x(i * span)
        w = D.col_w(span)
        hline(s, x, 2.42, 0.55, PAL.accent, 2.2)
        tf = textbox(s, x, 2.66, w, 0.95)
        p = para(tf, v, first=True, font=D.NUM, size=D.T_DISPLAY + 8, color=PAL.accent,
                 bold=True, line=1.0)
        r = p.add_run()
        r.text = " " + unit
        set_font(r, D.NUM, D.T_H4, MUTED, bold=False)
        tf = textbox(s, x, 3.82, w, 0.4)
        para(tf, label, first=True, font=D.SANS, size=D.T_H4 - 2, color=INK, line=1.3)
        tf = textbox(s, x, 4.30, w, 0.4)
        para(tf, note, first=True, font=D.SANS, size=D.T_SMALL, color=MUTED, line=1.4)
        tf = textbox(s, x, 4.74, w, 0.35)
        para(tf, delta, first=True, font=D.NUM, size=D.T_CAPTION + 1,
             color=PAL.accent, bold=True, line=1.2)
    page_num(s, idx, total)


def editorial(prs, idx, total):
    s = slide(prs)
    # 左栏：标签 + 引言
    tf = textbox(s, D.MARGIN_X, 1.05, 3.0, 0.3)
    para(tf, "01 / 观察", first=True, font=D.NUM, size=10.5, color=PAL.accent,
         bold=True, line=1.0, spacing=2.0)

    tf = textbox(s, D.MARGIN_X, 1.52, D.col_w(4) - 0.2, 3.4)
    para(tf, "信息密度", first=True, font=D.SERIF, size=D.T_H3, color=INK, line=1.2)
    para(tf, "不是越低越好", font=D.SERIF, size=D.T_H3, color=MUTED, line=1.2,
         space_before=2)

    # 右栏：带悬挂序号的要点
    x = D.col_x(5)
    w = D.col_w(7)
    tf = textbox(s, x, 1.05, w, 0.8)
    para(tf, "改造前后，同一页的差别", first=True, font=D.SANS, size=D.T_H4,
         color=INK, bold=True, line=1.3)
    hline(s, x, 1.72, w, RULE)

    rows = [
        ("少", "每页只留一个论点，其余拆到下一页"),
        ("齐", "标题、正文、图注共用一个左边界"),
        ("序", "字号只有五档，读者两页就能记住层级"),
        ("白", "行距 1.55，段落之间至少留半个字高"),
    ]
    y = 2.16
    for i, (mark, body) in enumerate(rows, start=1):
        tf = textbox(s, x, y, 0.45, 0.4)
        para(tf, f"{i:02d}", first=True, font=D.NUM, size=D.T_SMALL + 1,
             color=PAL.accent, bold=True, line=1.2)
        tf = textbox(s, x + 0.55, y - 0.02, w - 0.6, 0.9)
        para(tf, body, first=True, font=D.SANS, size=D.T_BODY, color=INK, line=1.55)
        y += 0.96
    page_num(s, idx, total)


def timeline(prs, idx, total):
    s = slide(prs)
    tf = textbox(s, D.MARGIN_X, 1.0, 9.0, 0.3)
    para(tf, "节奏", first=True, font=D.SANS, size=10.5, color=PAL.accent,
         bold=True, line=1.0, spacing=2.2)
    tf = textbox(s, D.MARGIN_X, 1.32, 9.0, 0.8)
    para(tf, "四个阶段，一条主线", first=True, font=D.SANS, size=D.T_H3,
         color=INK, bold=True, line=1.2)

    steps = [("问题定义", "梳理三条技术路线"), ("方法设计", "跨层注意力 + 蒸馏"),
             ("实验验证", "COCO 与自建数据集"), ("落地部署", "38 FPS / 9 W")]
    y = 3.05
    hline(s, D.MARGIN_X, y, D.content_w(), RULE)
    for i, (name, desc) in enumerate(steps):
        x = D.col_x(i * 3)
        w = D.col_w(3)
        rect(s, x, y - 0.055, 0.11, 0.11, PAL.accent, shape=MSO_SHAPE.OVAL)
        tf = textbox(s, x, y + 0.22, w, 0.3)
        para(tf, f"0{i + 1}", first=True, font=D.NUM, size=D.T_CAPTION,
             color=PAL.accent, bold=True, line=1.0, spacing=1.4)
        tf = textbox(s, x, y + 0.52, w, 0.45)
        para(tf, name, first=True, font=D.SANS, size=D.T_H4 - 2, color=INK,
             bold=True, line=1.3)
        tf = textbox(s, x, y + 1.02, w - 0.15, 0.8)
        para(tf, desc, first=True, font=D.SANS, size=D.T_SMALL, color=MUTED, line=1.5)
    page_num(s, idx, total)


def quote(prs, idx, total):
    s = slide(prs)
    rect(s, 0, 0, 0.22, D.SLIDE_H, PAL.accent_soft)
    tf = textbox(s, D.MARGIN_X + 0.15, 1.35, 1.6, 1.2)
    para(tf, "“", first=True, font=D.SERIF, size=120, color=PAL.accent_soft,
         bold=True, line=0.8)

    tf = textbox(s, D.MARGIN_X + 0.1, 2.55, 10.6, 2.6)
    para(tf, "留白不是空，是让读者有时间把上一句记住。", first=True,
         font=D.SERIF, size=D.T_H1 - 6, color=INK, line=1.6)

    hline(s, D.MARGIN_X + 0.1, 5.35, 0.75, GOLD, 2.2)
    tf = textbox(s, D.MARGIN_X + 0.1, 5.55, 8.0, 0.4)
    para(tf, "版式笔记 · 2026", first=True, font=D.SANS, size=D.T_SMALL,
         color=MUTED, line=1.2)
    page_num(s, idx, total)


def ending(prs):
    s = slide(prs)
    hline(s, D.SLIDE_W / 2 - 0.35, 2.55, 0.7, GOLD, 2.4)
    tf = textbox(s, 1.0, 2.85, D.content_w(), 1.4)
    para(tf, "谢谢", first=True, font=D.SERIF, size=60, color=INK,
         line=1.1, align=PP_ALIGN.CENTER)
    tf = textbox(s, 1.0, 4.35, D.content_w(), 0.5)
    para(tf, "字体：Signal · 网格：12 栏 · 色彩：OKLCH", first=True, font=D.NUM,
         size=D.T_CAPTION, color=MUTED, line=1.2, align=PP_ALIGN.CENTER, spacing=0.8)


def main() -> None:
    prs = Presentation()
    prs.slide_width = Inches(D.SLIDE_W)
    prs.slide_height = Inches(D.SLIDE_H)
    cover(prs)
    total = 6
    kpi(prs, 2, total)
    editorial(prs, 3, total)
    timeline(prs, 4, total)
    quote(prs, 5, total)
    ending(prs)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(OUT)
    print(f"[OK] {OUT}")


if __name__ == "__main__":
    main()
