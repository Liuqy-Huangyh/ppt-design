"""PPTX 体检 + 反解成 Markdown。"""
from __future__ import annotations

import re
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu

from .primitives import weighted_len

FOOTER_RE = re.compile(r"^\s*\d{1,2}\s*/\s*\d{1,2}\s*$")
BULLET_RE = re.compile(r"^[•▪·\-–—]\s*")


def _iter_texts(slide):
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text.strip()
        if text:
            yield shape, text


def _is_footer(shape, prs) -> bool:
    top = shape.top / 914400 if shape.top is not None else 0
    return top > prs.slide_height / 914400 - 0.55


def _font_sizes(shape) -> list[float]:
    sizes = []
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            if r.font.size:
                sizes.append(r.font.size.pt)
    return sizes


def _fit_ratio(shape) -> float:
    """粗略估计：这段文字需要的高度 / 文本框高度。>1.1 视为可能溢出。"""
    tf = shape.text_frame
    width_in = max(0.6, (shape.width or Emu(0)) / 914400)
    height_in = max(0.2, (shape.height or Emu(0)) / 914400)
    total_pt = 0.0
    for p in tf.paragraphs:
        text = "".join(r.text for r in p.runs)
        sizes = [r.font.size.pt for r in p.runs if r.font.size]
        size = sizes[0] if sizes else 18.0
        per_line = max(1.0, width_in * 72 / (size * 1.0))
        lines = max(1, int(weighted_len(text) / per_line) + 1) if text else 1
        total_pt += lines * size * 1.32
    return total_pt / (height_in * 72)


def _installed_fonts() -> set[str]:
    names: set[str] = set()
    try:
        import winreg
    except ImportError:
        return names
    key_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            with winreg.OpenKey(hive, key_path) as key:
                i = 0
                while True:
                    try:
                        name, _value, _type = winreg.EnumValue(key, i)
                    except OSError:
                        break
                    names.add(name.split(" (")[0])
                    i += 1
        except OSError:
            continue
    return names


def _fonts_used(slide) -> set[str]:
    used: set[str] = set()
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                if r.font.name:
                    used.add(r.font.name)
    return used


def qa(path: str | Path) -> dict:
    path = Path(path)
    prs = Presentation(str(path))
    W = prs.slide_width / 914400
    H = prs.slide_height / 914400
    fonts = _fonts_used_all = set()
    report = {
        "file": str(path), "size_in": (round(W, 2), round(H, 2)),
        "slides": [], "fonts": set(), "warnings": [],
    }

    for idx, slide in enumerate(prs.slides, start=1):
        chars = 0
        size_acc = 0.0
        size_n = 0
        size_min = 999.0
        overflow = []
        images = tables = 0
        notes = slide.has_notes_slide and bool(slide.notes_slide.notes_text_frame.text.strip())
        for shape in slide.shapes:
            if shape.shape_type == 13 or getattr(shape, "image", None):
                images += 1
            if getattr(shape, "has_table", False) and shape.has_table:
                tables += 1
            if not shape.has_text_frame:
                continue
            text = shape.text_frame.text.strip()
            if not text or _is_footer(shape, prs) or FOOTER_RE.match(text):
                continue
            chars += len(text)
            fonts |= _fonts_used_slide(shape)
            for size in _font_sizes(shape):
                size_acc += size
                size_n += 1
                size_min = min(size_min, size)
            if len(text) > 24:
                ratio = _fit_ratio(shape)
                if ratio > 1.12:
                    overflow.append(f"「{text[:14]}…」需 {ratio:.1f} 倍高度")
        report["slides"].append({
            "index": idx, "chars": chars, "images": images, "tables": tables,
            "notes": notes, "overflow": overflow,
            "avg_size": round(size_acc / size_n, 1) if size_n else 0,
            "min_size": size_min if size_n else 0,
        })

    report["fonts"] = fonts
    installed = _installed_fonts()
    missing = sorted(f for f in fonts if installed and not _has_font(f, installed))
    report["missing_fonts"] = missing

    w = report["warnings"]
    for s in report["slides"]:
        if s["overflow"]:
            w.append(f"第 {s['index']} 页文字可能溢出：{'；'.join(s['overflow'][:2])}")
        if s["chars"] > 260 and s.get("min_size", 0) >= 13:
            w.append(f"第 {s['index']} 页 {s['chars']} 字，信息密度偏高，建议拆页")
        if not s["notes"]:
            w.append(f"第 {s['index']} 页没有讲者备注")
    if missing:
        w.append(f"缺少字体：{', '.join(missing)}（会回退到系统默认，版式可能位移）")
    return report



# 注册表里中文字体常以英文名登记，做一层别名匹配
FONT_ALIASES = {
    "微软雅黑": "microsoft yahei",
    "微软雅黑 light": "microsoft yahei light",
    "宋体": "simsun",
    "黑体": "simhei",
    "等线": "dengxian",
    "楷体": "kaiti",
    "仿宋": "fangsong",
}


def _has_font(name: str, installed: set[str]) -> bool:
    low = name.lower()
    if any(low in item.lower() for item in installed):
        return True
    alt = FONT_ALIASES.get(low)
    return bool(alt and any(alt in item.lower() for item in installed))

def _fonts_used_slide(shape) -> set[str]:
    used = set()
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            if r.font.name:
                used.add(r.font.name)
    return used


def format_qa(report: dict) -> str:
    W, H = report["size_in"]
    lines = [f"文件：{report['file']}", f"画布：{W} × {H} in ｜ 共 {len(report['slides'])} 页", ""]
    lines.append("页码  字数  图  表  备注")
    for s in report["slides"]:
        lines.append(f"{s['index']:>4}  {s['chars']:>4}  {s['images']:>2}  {s['tables']:>2}  {'有' if s['notes'] else '无'}")
    lines.append("")
    lines.append("字体：" + ("、".join(sorted(report["fonts"])) or "（未显式设置）"))
    if report["missing_fonts"]:
        lines.append("缺失字体：" + "、".join(report["missing_fonts"]))
    lines.append("")
    if report["warnings"]:
        lines.append(f"检查结果：{len(report['warnings'])} 条提示")
        for item in report["warnings"][:24]:
            lines.append("  · " + item)
        if len(report["warnings"]) > 24:
            lines.append(f"  … 另有 {len(report['warnings']) - 24} 条")
    else:
        lines.append("检查结果：干净，没有发现明显问题。")
    return "\n".join(lines)


# ---------------------------------------------------------------- 反解

def to_markdown(path: str | Path, out: str | Path | None = None) -> Path:
    """把现成的 .pptx 反解成可再次编辑的 Markdown（图片导出到 media/）。"""
    path = Path(path)
    prs = Presentation(str(path))
    out = Path(out) if out else path.with_suffix(".md")
    media = out.parent / "media"
    media.mkdir(parents=True, exist_ok=True)

    blocks: list[str] = [f"---", f"title: {path.stem}", "---", ""]
    img_seq = 0

    for idx, slide in enumerate(prs.slides, start=1):
        title = ""
        bullets: list[str] = []
        tables: list[list[list[str]]] = []
        pics: list[str] = []
        for shape in slide.shapes:
            if getattr(shape, "has_table", False) and shape.has_table:
                rows = []
                for r in shape.table.rows:
                    rows.append([c.text.replace("\n", " ").strip() for c in r.cells])
                tables.append(rows)
                continue
            if getattr(shape, "shape_type", None) == 13:
                img_seq += 1
                try:
                    blob = shape.image.blob
                    ext = shape.image.ext or "png"
                    name = media / f"image{img_seq:02d}.{ext}"
                    name.write_bytes(blob)
                    pics.append(str(name.relative_to(out.parent)).replace("\\", "/"))
                except Exception:
                    pics.append("")
                continue
            if not shape.has_text_frame:
                continue
            text = shape.text_frame.text.strip()
            if not text or _is_footer(shape, prs) or FOOTER_RE.match(text):
                continue
            if not title and (shape.top or 0) / 914400 < 2.6 and len(text) < 60:
                title = text.replace("\n", " ")
                continue
            for p in shape.text_frame.paragraphs:
                line = "".join(r.text for r in p.runs).strip()
                if not line:
                    continue
                level = 0
                if p.level:
                    level = min(1, p.level)
                bullets.append("  " * level + "- " + BULLET_RE.sub("", line))

        blocks.append(f"# {title or f'第 {idx} 页'}")
        blocks.extend(bullets)
        for table in tables:
            if not table:
                continue
            blocks.append("")
            blocks.append("| " + " | ".join(table[0]) + " |")
            blocks.append("|" + "---|" * len(table[0]))
            for row in table[1:]:
                blocks.append("| " + " | ".join(row) + " |")
        if pics:
            blocks.append("")
            for i, p in enumerate(pics):
                blocks.append(f"![图{i + 1}]({p})")
        note = ""
        if slide.has_notes_slide:
            note = slide.notes_slide.notes_text_frame.text.strip()
        if note:
            blocks.append("")
            blocks.append("::: notes")
            blocks.extend(note.splitlines())
            blocks.append(":::")
        blocks.append("")
        blocks.append("---")
        blocks.append("")

    out.write_text("\n".join(blocks), encoding="utf-8")
    return out
