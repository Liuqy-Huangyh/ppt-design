"""风格档案：从案例 PPT 里读出配色 / 字体 / 信息密度，存取档案、按反馈调参。

对外接口（docs/CONTRACT.md 第 6、7 节）
------------------------------------
    analyze(pptx_path, *, name="", source="") -> dict
    save(profile, store_dir) -> dict          # 落盘 profiles/<id>.json + cases/<id>/source.pptx
    load(store_dir, profile_id) -> dict
    list_all(store_dir) -> list[dict]
    apply(profile, base_theme=None) -> Theme
    learn(profile, rating, notes="") -> dict

设计要点
--------
* python-pptx 不暴露主题（ppt/theme/theme1.xml），所以配色直接解压读 XML；
  但真正决定观感的往往是幻灯片上的显式颜色，因此再扫一遍每页的填充/文字颜色，
  观察到的颜色优先，主题值兜底，最后落到安全默认值。
* analyze() 对任何输入都不能抛异常：最差返回一份默认档案。
"""
from __future__ import annotations

import copy
import json
import os
import re
import shutil
import statistics
import zipfile
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

from .themes import Theme, clone_with, get_theme, rgb

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
EMU_PER_INCH = 914400.0

DEFAULT_PALETTE = {
    "bg": "#FFFFFF",
    "ink": "#14213D",
    "muted": "#5C6B82",
    "accent": "#1B3A6B",
    "accent_soft": "#E9EEF7",
    "rule": "#D5DCE8",
}
PALETTE_KEYS = ("bg", "ink", "muted", "accent", "accent_soft", "rule")

DEFAULT_FONTS = {"title": "微软雅黑", "body": "微软雅黑"}
DEFAULT_METRICS = {
    "slides": 0,
    "chars_per_slide": 0.0,
    "bullets_per_slide": 0.0,
    "title_pt": 31.0,
    "body_pt": 19.0,
    "image_ratio": 0.0,
}
DEFAULT_STYLE = {"density": 1.0, "title_scale": 1.0, "bullet_char": "▪"}

DENSITY_RANGE = (0.6, 1.4)
TITLE_SCALE_RANGE = (0.85, 1.25)
HISTORY_LIMIT = 200          # learn.history 只保留最近 N 条，rounds 继续累加

_ID_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$")
_HEX_RE = re.compile(r"^[0-9A-Fa-f]{6}$")
_HEX8_RE = re.compile(r"^[0-9A-Fa-f]{8}$")


# ------------------------------------------------------------------ 小工具


def _clamp(value: float, lo: float, hi: float) -> float:
    return max(lo, min(hi, value))


def _now() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def _to_hex(value) -> str | None:
    """任意颜色表示 → '#RRGGBB'（大写）；认不出来返回 None。"""
    if value is None:
        return None
    if isinstance(value, str):
        s = value.strip().lstrip("#")
        if _HEX_RE.match(s):
            return "#" + s.upper()
        if _HEX8_RE.match(s):          # ARGB
            return "#" + s[2:].upper()
        return None
    try:
        r, g, b = int(value[0]), int(value[1]), int(value[2])
    except (TypeError, ValueError, IndexError, KeyError):
        return None
    if not all(0 <= c <= 255 for c in (r, g, b)):
        return None
    return "#%02X%02X%02X" % (r, g, b)


def _channels(hex_color: str) -> tuple[int, int, int]:
    s = hex_color.lstrip("#")
    return int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16)


def _lum(hex_color: str) -> float:
    """0（黑）– 1（白）的近似明度，用来判断深浅与对比。"""
    r, g, b = _channels(hex_color)
    return (r * 299 + g * 587 + b * 114) / 1000.0 / 255.0


def _contrast(a: str, b: str) -> float:
    return abs(_lum(a) - _lum(b))


def _dist(a: str, b: str) -> float:
    """两个颜色在 RGB 空间里的欧氏距离。"""
    return sum((x - y) ** 2 for x, y in zip(_channels(a), _channels(b))) ** 0.5


def _mix(a: str, b: str, weight_b: float) -> str:
    """a 与 b 混合，weight_b 是 b 的权重（0–1）。"""
    weight_b = _clamp(weight_b, 0.0, 1.0)
    ar, ag, ab = _channels(a)
    br, bg_, bb = _channels(b)
    out = tuple(
        int(round(x * (1 - weight_b) + y * weight_b))
        for x, y in ((ar, br), (ag, bg_), (ab, bb))
    )
    return "#%02X%02X%02X" % out


def _safe_median(values: list[float], default: float) -> float:
    clean = [float(v) for v in values if v is not None]
    if not clean:
        return default
    return round(statistics.median(clean), 1)


def _emu_in(value) -> float | None:
    if value is None:
        return None
    try:
        return float(value) / EMU_PER_INCH
    except (TypeError, ValueError):
        return None


# ------------------------------------------------------------------ 主题 XML


def _find_theme_root(pptx_path: Path):
    """从 zip 里读 ppt/theme/theme1.xml（找不到就用任意 themeN.xml）。"""
    try:
        with zipfile.ZipFile(pptx_path) as zf:
            names = [
                n for n in zf.namelist()
                if re.fullmatch(r"ppt/theme/theme\d+\.xml", n)
            ]
            names.sort(key=lambda n: (0 if n.endswith("theme1.xml") else 1, n))
            for name in names:
                try:
                    return ET.fromstring(zf.read(name))
                except Exception:
                    continue
    except (zipfile.BadZipFile, OSError, KeyError, RuntimeError):
        return None
    return None


def _scheme_color(el) -> str | None:
    for child in el:
        tag = child.tag.split("}")[-1]
        if tag == "srgbClr":
            return _to_hex(child.get("val"))
        if tag == "sysClr":
            return _to_hex(child.get("lastClr"))
    return None


def _theme_font(font_el, prefer_cjk: bool = True) -> str:
    """主题字体：中文稿优先东亚字形，西文稿优先 latin。"""
    if font_el is None:
        return ""
    ea = font_el.find(f"{A}ea")
    if ea is not None:
        typeface = (ea.get("typeface") or "").strip()
        if typeface and not typeface.startswith("+"):
            return typeface
    hans = ""
    for el in font_el.findall(f"{A}font"):
        if el.get("script") in ("Hans", "Hant"):
            typeface = (el.get("typeface") or "").strip()
            if typeface and not typeface.startswith("+"):
                hans = typeface
                break
    if prefer_cjk and hans:
        return hans
    latin = font_el.find(f"{A}latin")
    if latin is not None:
        typeface = (latin.get("typeface") or "").strip()
        if typeface and not typeface.startswith("+"):
            return typeface
    return hans


def _parse_theme(root, prefer_cjk: bool = True) -> tuple[dict, dict]:
    colors: dict[str, str] = {}
    fonts = {"major": "", "minor": ""}
    if root is None:
        return colors, fonts
    clr = root.find(f".//{A}clrScheme")
    if clr is not None:
        for child in clr:
            name = child.tag.split("}")[-1]
            hex_color = _scheme_color(child)
            if hex_color:
                colors[name] = hex_color
    scheme = root.find(f".//{A}fontScheme")
    if scheme is not None:
        fonts["major"] = _theme_font(scheme.find(f"{A}majorFont"), prefer_cjk)
        fonts["minor"] = _theme_font(scheme.find(f"{A}minorFont"), prefer_cjk)
    return colors, fonts


# ------------------------------------------------------------------ 幻灯片观察


def _slide_bg(slide) -> str | None:
    try:
        fill = slide.background.fill
        fill_type = getattr(fill, "type", None)
        if fill_type is not None and getattr(fill_type, "value", fill_type) == 1:
            hex_color = _to_hex(fill.fore_color.rgb)
            if hex_color:
                return hex_color
    except Exception:
        pass
    try:
        bg = slide._element.find(f"{P}bg")
        if bg is not None:
            srgb = bg.find(f".//{A}srgbClr")
            if srgb is not None:
                return _to_hex(srgb.get("val"))
    except Exception:
        pass
    return None


def _solid_fill(shape) -> str | None:
    try:
        sp_pr = shape._element.find(f"{P}spPr")
        if sp_pr is None:
            return None
        solid = sp_pr.find(f"{A}solidFill")
        if solid is None:
            return None
        srgb = solid.find(f"{A}srgbClr")
        if srgb is not None:
            return _to_hex(srgb.get("val"))
    except Exception:
        pass
    return None


def _line_color(shape) -> str | None:
    try:
        sp_pr = shape._element.find(f"{P}spPr")
        if sp_pr is None:
            return None
        ln = sp_pr.find(f"{A}ln")
        if ln is None:
            return None
        solid = ln.find(f"{A}solidFill")
        if solid is None:
            return None
        srgb = solid.find(f"{A}srgbClr")
        if srgb is not None:
            return _to_hex(srgb.get("val"))
    except Exception:
        pass
    return None


def _table_fills(shape) -> list[str]:
    """表格单元格底色（只看 tcPr 直接子节点，别把单元格里的文字颜色当成填充）。"""
    out: list[str] = []
    try:
        for tc_pr in shape._element.iter(f"{A}tcPr"):
            solid = tc_pr.find(f"{A}solidFill")
            if solid is None:
                continue
            srgb = solid.find(f"{A}srgbClr")
            if srgb is not None:
                hex_color = _to_hex(srgb.get("val"))
                if hex_color:
                    out.append(hex_color)
    except Exception:
        pass
    return out


def _iter_shapes(shapes, depth: int = 0):
    for shape in shapes:
        yield shape
        if depth >= 2:
            continue
        try:
            nested = getattr(shape, "shapes", None)
        except Exception:
            nested = None
        if nested is not None:
            yield from _iter_shapes(nested, depth + 1)


def _run_props(run) -> tuple[float | None, str, str | None]:
    """(字号 pt, 字体名, '#RRGGBB')；读不到的给 None / 空串。"""
    size = None
    font = ""
    color = None
    try:
        r_pr = run._r.find(f"{A}rPr")
    except Exception:
        r_pr = None
    if r_pr is not None:
        raw_size = r_pr.get("sz")
        if raw_size and raw_size.isdigit():
            size = round(int(raw_size) / 100.0, 1)
        solid = r_pr.find(f"{A}solidFill")
        if solid is not None:
            srgb = solid.find(f"{A}srgbClr")
            if srgb is not None:
                color = _to_hex(srgb.get("val"))
        for tag in ("ea", "latin", "cs"):
            el = r_pr.find(f"{A}{tag}")
            if el is None:
                continue
            typeface = (el.get("typeface") or "").strip()
            if typeface and not typeface.startswith("+"):
                font = typeface
                break
    if size is None:
        try:
            if run.font.size is not None:
                size = round(run.font.size.pt, 1)
        except Exception:
            pass
    if color is None:
        try:
            color = _to_hex(run.font.color.rgb)
        except Exception:
            pass
    if not font:
        try:
            font = (run.font.name or "").strip()
        except Exception:
            font = ""
        if font.startswith("+"):
            font = ""
    return size, font, color


def _observe(prs) -> dict:
    """扫全片：背景/填充/线条颜色、字体、字号、每页字数与要点数。"""
    slide_w = _emu_in(getattr(prs, "slide_width", None)) or 13.333
    slide_h = _emu_in(getattr(prs, "slide_height", None)) or 7.5
    area = max(1e-6, slide_w * slide_h)

    bg_counter: Counter = Counter()
    big_fill_counter: Counter = Counter()
    text_colors: Counter = Counter()
    fill_counter: Counter = Counter()
    line_counter: Counter = Counter()
    title_fonts: Counter = Counter()
    body_fonts: Counter = Counter()
    title_sizes: list[float] = []
    body_sizes: list[float] = []
    chars_total = 0.0
    cjk_total = 0
    bullet_total = 0
    picture_slides = 0
    slide_count = 0

    for slide in prs.slides:
        slide_count += 1
        bg = _slide_bg(slide)
        if bg:
            bg_counter[bg] += 1

        slide_picture = 0.0
        # (text, size, font, color, shape_top, paragraph_key)
        records: list[tuple[str, float | None, str, str | None, float, int]] = []
        paragraphs: dict[int, dict] = defaultdict(
            lambda: {"chars": 0, "sizes": [], "top": slide_h}
        )

        for shape in _iter_shapes(slide.shapes):
            top = _emu_in(getattr(shape, "top", None))
            top = 0.0 if top is None else top
            width = _emu_in(getattr(shape, "width", None)) or 0.0
            height = _emu_in(getattr(shape, "height", None)) or 0.0
            shape_area = max(0.0, width * height) / area

            shape_type = getattr(shape, "shape_type", None)
            if shape_type is not None and getattr(shape_type, "value", shape_type) == 13:
                slide_picture += shape_area

            fill = _solid_fill(shape)
            if fill:
                fill_counter[fill] += 1
                if shape_area >= 0.85:
                    big_fill_counter[fill] += 1

            line = _line_color(shape)
            if line:
                line_counter[line] += 1

            try:
                for cell_fill in _table_fills(shape):
                    fill_counter[cell_fill] += 1
            except Exception:
                pass

            try:
                has_tf = bool(shape.has_text_frame)
            except Exception:
                has_tf = False
            if not has_tf:
                continue
            try:
                paras = list(shape.text_frame.paragraphs)
            except Exception:
                continue
            for para in paras:
                key = id(para)
                info = paragraphs[key]
                info["top"] = min(info["top"], top)
                for run in para.runs:
                    text = run.text or ""
                    if not text.strip():
                        continue
                    size, font, color = _run_props(run)
                    records.append((text, size, font, color, top, key))
                    info["chars"] += len(text)
                    cjk_total += sum(1 for ch in text if "\u4e00" <= ch <= "\u9fff")
                    if size:
                        info["sizes"].append(size)

        # 本页最大字号 = 标题字号
        sizes = [rec[1] for rec in records if rec[1]]
        max_size = max(sizes) if sizes else None
        title_keys: set[int] = set()
        if max_size and max_size >= 16:
            title_sizes.append(max_size)
            for text, size, font, color, _top, key in records:
                if size != max_size:
                    continue
                title_keys.add(key)
                if font:
                    title_fonts[font] += len(text)

        slide_body_sizes: list[float] = []
        for text, size, font, color, top, key in records:
            if color:
                text_colors[color] += len(text)
            if key in title_keys:
                continue
            if top > slide_h - 0.95:          # 页脚 / 页码
                continue
            if font:
                body_fonts[font] += len(text)
            if size and 10 <= size:
                slide_body_sizes.append(size)

        for key, info in paragraphs.items():
            if key in title_keys:
                continue
            if info["chars"] < 4:
                continue
            if info["top"] > slide_h - 0.95:  # 页脚整段跳过
                continue
            bullet_total += 1
            chars_total += info["chars"]

        if slide_body_sizes:
            body_sizes.append(statistics.median(slide_body_sizes))
        if slide_picture > 0:
            picture_slides += 1

    return {
        "slides": slide_count,
        "bg": bg_counter,
        "big_fill": big_fill_counter,
        "text": text_colors,
        "fill": fill_counter,
        "line": line_counter,
        "title_fonts": title_fonts,
        "body_fonts": body_fonts,
        "title_sizes": title_sizes,
        "body_sizes": body_sizes,
        "chars_total": chars_total,
        "has_cjk": cjk_total >= 4,
        "bullets_total": bullet_total,
        "picture_slides": picture_slides,
        "slide_w": slide_w,
        "slide_h": slide_h,
    }


# ------------------------------------------------------------------ 归纳成档案


def _pick_bg(obs: dict, theme_colors: dict) -> str:
    slides = max(1, obs["slides"])
    if obs["big_fill"]:
        color, count = obs["big_fill"].most_common(1)[0]
        if count >= max(2, slides * 0.4):
            return color
    if obs["bg"]:
        return obs["bg"].most_common(1)[0][0]
    for key in ("lt1", "bg1"):
        if theme_colors.get(key):
            return theme_colors[key]
    return DEFAULT_PALETTE["bg"]


def _build_palette(obs: dict, theme_colors: dict) -> dict:
    palette = dict(DEFAULT_PALETTE)
    bg = _pick_bg(obs, theme_colors)
    palette["bg"] = bg

    text_rank = [
        color for color, _ in obs["text"].most_common()
        if color and _contrast(color, bg) > 0.12
    ]
    # 主文字：出现最多的深（浅）色
    ink = text_rank[0] if text_rank else (theme_colors.get("dk1") or DEFAULT_PALETTE["ink"])
    if _contrast(ink, bg) < 0.12:
        ink = DEFAULT_PALETTE["ink"]
    palette["ink"] = ink

    # 次要文字：与底色对比适中、且不是主文字的那个
    muted = None
    for color in text_rank[1:]:
        if color == ink:
            continue
        if 0.15 <= _contrast(color, bg) <= 0.72:
            muted = color
            break
    palette["muted"] = muted or _mix(ink, bg, 0.35)

    # 强调色：优先用色块，其次用文字里最"跳"的颜色
    accent = None
    for color, _count in obs["fill"].most_common():
        if color in (bg, ink, palette["muted"]):
            continue
        if _contrast(color, bg) >= 0.22:
            accent = color
            break
    if accent is None:
        for color in text_rank[1:]:
            if color not in (ink, palette["muted"]) and _contrast(color, bg) >= 0.22:
                accent = color
                break
    accent = accent or theme_colors.get("accent1") or DEFAULT_PALETTE["accent"]
    if _contrast(accent, bg) < 0.12:
        accent = DEFAULT_PALETTE["accent"]
    palette["accent"] = accent

    # 浅底强调色：最接近"强调色淡化到 90% 底色"的那个填充或文字色
    tint = _mix(accent, bg, 0.90)
    accent_soft = None
    best_score = None
    soft_pool: Counter = Counter()
    soft_pool.update(obs["fill"])
    soft_pool.update(obs["text"])
    for color, _count in soft_pool.most_common():
        if color in (bg, accent, ink, palette["muted"]):
            continue
        if _contrast(color, bg) > 0.45 or _contrast(color, bg) < 0.015:
            continue
        score = _dist(color, tint)
        if best_score is None or score < best_score:
            accent_soft, best_score = color, score
    palette["accent_soft"] = accent_soft or tint

    # 分隔线：接近"主文字淡化到 85% 底色"的中性色（描边/浅文字/浅色块都能当证据）
    neutral = _mix(ink, bg, 0.85)
    rule = None
    best_score = None
    candidates: Counter = Counter()
    candidates.update(obs["line"])
    candidates.update(obs["text"])
    candidates.update(obs["fill"])
    for color, _count in candidates.most_common():
        if color in (bg, ink, palette["muted"], accent, palette["accent_soft"]):
            continue
        if _contrast(color, bg) < 0.03:
            continue
        score = _dist(color, neutral)
        if score > 140:
            continue
        if best_score is None or score < best_score:
            rule, best_score = color, score
    palette["rule"] = rule or neutral

    for key in PALETTE_KEYS:
        if not _to_hex(palette.get(key)):
            palette[key] = DEFAULT_PALETTE[key]
        else:
            palette[key] = _to_hex(palette[key])
    return palette


def _build_fonts(obs: dict, theme_fonts: dict) -> dict:
    title = obs["title_fonts"].most_common(1)[0][0] if obs["title_fonts"] else ""
    body = obs["body_fonts"].most_common(1)[0][0] if obs["body_fonts"] else ""
    return {
        "title": title or theme_fonts.get("major") or DEFAULT_FONTS["title"],
        "body": body or theme_fonts.get("minor") or DEFAULT_FONTS["body"],
    }


def _build_metrics(obs: dict) -> dict:
    slides = obs["slides"]
    metrics = dict(DEFAULT_METRICS)
    metrics["slides"] = slides
    if not slides:
        return metrics
    metrics["chars_per_slide"] = round(obs["chars_total"] / slides, 1)
    metrics["bullets_per_slide"] = round(obs["bullets_total"] / slides, 1)
    metrics["title_pt"] = _safe_median(obs["title_sizes"], DEFAULT_METRICS["title_pt"])
    metrics["body_pt"] = _safe_median(obs["body_sizes"], DEFAULT_METRICS["body_pt"])
    metrics["image_ratio"] = round(min(1.0, obs["picture_slides"] / slides), 2)
    return metrics


# ------------------------------------------------------------------ 版式骨架
# 10 个固定比例（页坐标 0–1）+ 来源/置信度。pptx 用形状几何测量，图片用 PIL 墨迹投影估计；
# 生成时按 layout_hints() 只做「标题纵向位置 + 左右边距」的限幅微调。

LAYOUT_KEYS = (
    "title_top_ratio", "title_height_ratio",
    "margin_left_ratio", "margin_right_ratio",
    "margin_top_ratio", "margin_bottom_ratio",
    "image_area_ratio", "text_area_ratio",
    "alignment", "columns",
)

# 默认骨架 =「没有可学证据」时的中性值：title_top_ratio 与版式的标题基准一致，
# 因此默认档案算出的位移恒为 0，生成结果与未接入骨架时逐像素一致。
DEFAULT_LAYOUT = {
    "title_top_ratio": 0.16,
    "title_height_ratio": 0.08,
    "margin_left_ratio": 0.075,
    "margin_right_ratio": 0.075,
    "margin_top_ratio": 0.06,
    "margin_bottom_ratio": 0.06,
    "image_area_ratio": 0.0,
    "text_area_ratio": 0.0,
    "alignment": "left",
    "columns": 1,
    "source": "default",
    "confidence": 0.0,
}

# 版式基准（对应 layouts.py 现有坐标）：标题带顶 0.16 页高、左右边距 0.075 页宽。
_BASELINE_TITLE_TOP = 0.16
_BASELINE_MARGIN_X = 0.075
_TITLE_SHIFT_LIMIT = 0.15        # ±0.15 页高
_MARGIN_SHIFT_LIMIT = 0.10       # ±0.10 页宽
_FOOTER_TOP = 0.88               # 页脚/页码区域，排除在内容 bbox 之外


def _shape_rect(shape, page_w: float, page_h: float) -> dict | None:
    """形状相对页面的矩形（0–1）；读不到尺寸时返回 None。"""
    left = _emu_in(getattr(shape, "left", None))
    top = _emu_in(getattr(shape, "top", None))
    width = _emu_in(getattr(shape, "width", None))
    height = _emu_in(getattr(shape, "height", None))
    if left is None or top is None or width is None or height is None:
        return None
    if width <= 0 or height <= 0 or page_w <= 0 or page_h <= 0:
        return None
    l = _clamp(left / page_w, 0.0, 1.0)
    t = _clamp(top / page_h, 0.0, 1.0)
    r = _clamp((left + width) / page_w, 0.0, 1.0)
    b = _clamp((top + height) / page_h, 0.0, 1.0)
    if r <= l or b <= t:
        return None
    return {"left": l, "top": t, "right": r, "bottom": b,
            "area": (r - l) * (b - t)}


def _shape_kind(shape) -> str | None:
    """'image' / 'text' / None。图片优先（带文字的画布不算图片）。"""
    try:
        shape_type = getattr(shape, "shape_type", None)
        if shape_type is not None and getattr(shape_type, "value", shape_type) == 13:
            return "image"
    except Exception:
        pass
    try:
        if getattr(shape, "has_text_frame", False) and shape.text_frame.text.strip():
            return "text"
    except Exception:
        pass
    return None


def _shape_font_size(shape) -> float:
    sizes = []
    try:
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                if run.font.size:
                    sizes.append(run.font.size.pt)
    except Exception:
        pass
    return max(sizes) if sizes else 0.0


def _paragraph_alignment(shape) -> str | None:
    try:
        for paragraph in shape.text_frame.paragraphs:
            value = paragraph.alignment
            if value is None:
                continue
            text = str(value)
            if "CENTER" in text:
                return "center"
            if "LEFT" in text or "JUSTIFY" in text:
                return "left"
    except Exception:
        pass
    return None


def _text_columns(rects: list[dict]) -> int:
    """按左右文本块判断 1/2 栏：两簇不重叠、左右都有内容才算两栏。"""
    if len(rects) < 2:
        return 1
    left = [r for r in rects if r["right"] <= 0.60]
    right = [r for r in rects if r["left"] >= 0.40]
    if not left or not right:
        return 1
    left_edge = max(r["right"] for r in left)
    right_edge = min(r["left"] for r in right)
    return 2 if left_edge <= right_edge + 0.05 else 1


def _slide_layout_metrics(slide, page_w: float, page_h: float) -> dict | None:
    """单页骨架：标题 bbox、内容 bbox、图文面积、栏数、对齐。"""
    text_rects: list[dict] = []
    image_rects: list[dict] = []
    title_rect = None
    title_size = -1.0
    title_align = None

    for shape in _iter_shapes(slide.shapes):
        kind = _shape_kind(shape)
        rect = _shape_rect(shape, page_w, page_h)
        if kind is None or rect is None:
            continue
        if rect["top"] > _FOOTER_TOP:      # 页脚 / 页码
            continue
        if kind == "image":
            image_rects.append(rect)
            continue
        text_rects.append(rect)
        size = _shape_font_size(shape)
        if rect["top"] < 0.60 and size > title_size:
            title_rect, title_size, title_align = rect, size, _paragraph_alignment(shape)

    content = text_rects + image_rects
    if not content:
        return None

    if title_rect is None and text_rects:
        title_rect = min(text_rects, key=lambda r: r["top"])

    body_rects = [r for r in text_rects if r is not title_rect]
    if not title_rect:
        title_rect = content[0]

    image_area = _clamp(sum(r["area"] for r in image_rects), 0.0, 1.0)
    text_area = _clamp(sum(r["area"] for r in text_rects), 0.0, 1.0)

    align = title_align
    if align is None:
        center = (title_rect["left"] + title_rect["right"]) / 2
        span = title_rect["right"] - title_rect["left"]
        align = "center" if abs(center - 0.5) <= 0.04 and span <= 0.8 else "left"

    return {
        "title_top_ratio": title_rect["top"],
        "title_height_ratio": title_rect["bottom"] - title_rect["top"],
        "margin_left_ratio": min(r["left"] for r in content),
        "margin_right_ratio": 1.0 - max(r["right"] for r in content),
        "margin_top_ratio": min(r["top"] for r in content),
        "margin_bottom_ratio": 1.0 - max(r["bottom"] for r in content),
        "image_area_ratio": image_area,
        "text_area_ratio": text_area,
        "alignment": align,
        "columns": _text_columns(body_rects),
    }


def _median(values: list[float], default: float) -> float:
    clean = [float(v) for v in values if v is not None]
    if not clean:
        return default
    return statistics.median(clean)


def _deck_layout(prs) -> dict:
    """全 deck 取中位数（对齐/栏数取众数），得到版式骨架。"""
    page_w = _emu_in(getattr(prs, "slide_width", None)) or 13.333
    page_h = _emu_in(getattr(prs, "slide_height", None)) or 7.5
    per_slide: list[dict] = []
    align_votes: Counter = Counter()
    column_votes: Counter = Counter()
    for slide in prs.slides:
        metrics = _slide_layout_metrics(slide, page_w, page_h)
        if not metrics:
            continue
        per_slide.append(metrics)
        align_votes[metrics["alignment"]] += 1
        column_votes[metrics["columns"]] += 1

    if not per_slide:
        return dict(DEFAULT_LAYOUT)

    layout = {}
    for key in ("title_top_ratio", "title_height_ratio",
                "margin_left_ratio", "margin_right_ratio",
                "margin_top_ratio", "margin_bottom_ratio",
                "image_area_ratio", "text_area_ratio"):
        layout[key] = round(_clamp(_median([m[key] for m in per_slide],
                                           DEFAULT_LAYOUT[key]), 0.0, 1.0), 3)
    layout["alignment"] = align_votes.most_common(1)[0][0] if align_votes else "left"
    layout["columns"] = column_votes.most_common(1)[0][0] if column_votes else 1
    layout["source"] = "pptx"
    layout["confidence"] = 1.0
    return layout


def _layout_block(layout: dict, source: str, confidence: float) -> dict:
    """规范化 layout 块：固定 10 字段 + source/confidence。"""
    out = dict(DEFAULT_LAYOUT)
    if isinstance(layout, dict):
        for key in LAYOUT_KEYS:
            if key in layout and layout[key] is not None:
                out[key] = layout[key]
    out["source"] = str(source or out.get("source") or "pptx")
    out["confidence"] = round(_clamp(float(confidence), 0.0, 1.0), 2)
    return out


def layout_hints(profile: dict | None) -> dict:
    """把档案里的 layout 骨架换算成生成时可用的微调量（英寸）。

    返回：
        title_shift   标题纵向位移（已限幅 ±0.15 页高）
        margin_left   左边距位移（±0.10 页宽）
        margin_right  右边距位移
        alignment / columns / source / confidence / layout：原样带出，便于调用方记录
    无档案或默认档案时全部为 0，版式与未接入骨架时逐像素一致。
    """
    profile = profile if isinstance(profile, dict) else {}
    layout = profile.get("layout") if isinstance(profile.get("layout"), dict) else {}
    layout = _layout_block(layout, layout.get("source", "default"),
                           layout.get("confidence", 0.0))

    size = profile.get("slide_size")
    try:
        page_w = float(size[0]) if size else 13.333
        page_h = float(size[1]) if size else 7.5
    except (TypeError, ValueError, IndexError):
        page_w, page_h = 13.333, 7.5

    title_shift = _clamp((float(layout["title_top_ratio"]) - _BASELINE_TITLE_TOP) * page_h,
                         -_TITLE_SHIFT_LIMIT * page_h, _TITLE_SHIFT_LIMIT * page_h)
    margin_left = _clamp((float(layout["margin_left_ratio"]) - _BASELINE_MARGIN_X) * page_w,
                         -_MARGIN_SHIFT_LIMIT * page_w, _MARGIN_SHIFT_LIMIT * page_w)
    margin_right = _clamp((float(layout["margin_right_ratio"]) - _BASELINE_MARGIN_X) * page_w,
                          -_MARGIN_SHIFT_LIMIT * page_w, _MARGIN_SHIFT_LIMIT * page_w)

    def _round(value: float) -> float:
        value = round(value, 3)
        return 0.0 if abs(value) < 0.005 else value

    return {
        "title_shift": _round(title_shift),
        "margin_left": _round(margin_left),
        "margin_right": _round(margin_right),
        "alignment": layout["alignment"],
        "columns": layout["columns"],
        "source": layout["source"],
        "confidence": float(layout["confidence"]),
        "layout": layout,
    }



def _image_layout(image_path: Path) -> dict:
    """一张幻灯片图片的骨架估计（纯 PIL + numpy，不做 OCR）。

    做法：
    * 以四边中位色当背景，灰度差 > 28 记为「墨迹」；
    * 行/列墨迹投影 → 内容 bbox → 四边留白比例；
    * 顶部连续墨迹带 → 标题位置/高度；
    * 16px 图块的颜色方差 + 饱和度 → 图片面积（照片/图表命中，纯色块不算）；
    * 标题带以下列投影里最长的空档 → 1/2 栏；标题带重心 → 左对齐/居中。
    """
    from PIL import Image
    import numpy as np

    with Image.open(image_path) as img:
        img = img.convert("RGB")
        img.thumbnail((1280, 1280))
        arr = np.asarray(img, dtype=np.float32)
    h, w = arr.shape[:2]
    if h < 16 or w < 16:
        raise ValueError("图片太小")

    weights = np.array([0.299, 0.587, 0.114], dtype=np.float32)
    gray = arr @ weights

    # 背景：四边 6% 区域的中位色
    edge = max(2, int(min(h, w) * 0.06))
    border = np.concatenate([
        arr[:edge].reshape(-1, 3), arr[-edge:].reshape(-1, 3),
        arr[:, :edge].reshape(-1, 3), arr[:, -edge:].reshape(-1, 3),
    ])
    bg_gray = float(np.median(border, axis=0) @ weights)
    ink = np.abs(gray - bg_gray) > 28.0
    if int(ink.sum()) < max(16, int(0.002 * h * w)):
        return _layout_block(dict(DEFAULT_LAYOUT), "image", 0.1)

    # 图片块：16px 图块的颜色方差 + 饱和度（照片/图表命中，纯色块会被低估）
    tile = 16
    th, tw = h // tile, w // tile
    image_pixels = np.zeros_like(ink)
    if th and tw:
        a = arr[:th * tile, :tw * tile].reshape(th, tile, tw, tile, 3)
        tile_max = a.max(axis=-1)
        tile_min = a.min(axis=-1)
        sat = (tile_max - tile_min) / np.maximum(1.0, tile_max)
        # 只有「整块大多数像素都有彩色」才算图片：深色文字所在图块也有高方差/高饱和，
        # 但彩色像素占比低（白底 + 笔画），用占比把它排除掉。
        colorful_frac = (sat > 0.25).mean(axis=(1, 3))
        var_t = a.var(axis=(1, 3)).mean(axis=-1)
        mean_t = a.mean(axis=(1, 3)).mean(axis=-1)
        image_tiles = (var_t > 260.0) & (colorful_frac > 0.45) & (mean_t > 25.0)
        image_pixels[:th * tile, :tw * tile] = np.repeat(
            np.repeat(image_tiles, tile, axis=0), tile, axis=1)
    image_area = float(image_pixels.mean())
    # 标题/栏数/对齐只看「非图片」墨迹，避免整块照片被当成标题
    text_ink = ink & ~image_pixels
    if int(text_ink.sum()) < max(16, int(0.0005 * h * w)):
        text_ink = ink

    # 内容 bbox（文字 + 图片）：四边留白
    row_ink = ink.mean(axis=1)
    col_ink = ink.mean(axis=0)
    row_thr = max(0.003, min(0.03, float(row_ink.max()) * 0.03))
    col_thr = max(0.003, min(0.03, float(col_ink.max()) * 0.03))
    rows = np.where(row_ink > row_thr)[0]
    col_candidates = np.where(col_ink > col_thr)[0]
    # 贴边、几乎整页高的窄竖线是装饰（封面左侧色条），不算内容边距
    def _decoration(x: int) -> bool:
        coverage = float(ink[:, x].mean())
        return coverage > 0.85 and (x < w * 0.03 or x >= w * 0.97)

    col_idx = np.array([int(x) for x in col_candidates if not _decoration(int(x))],
                       dtype=int)
    if col_idx.size == 0:
        col_idx = col_candidates
    if rows.size == 0 or col_idx.size == 0:
        return _layout_block(dict(DEFAULT_LAYOUT), "image", 0.15)
    top_row, bottom_row = int(rows[0]), int(rows[-1])
    left_col, right_col = int(col_idx[0]), int(col_idx[-1])

    # 顶部标题带：在文字墨迹上按行切段（允许 ≤3% 页高的小空隙），取顶部最密集的一段
    text_rows = text_ink.mean(axis=1)
    threshold = max(0.01, float(text_rows.max()) * 0.25)
    gap_limit = max(1, int(h * 0.03))
    runs: list[tuple[int, int, float]] = []
    start = None
    last = None
    gap = 0
    for y in range(top_row, bottom_row + 1):
        if text_rows[y] >= threshold:
            if start is None:
                start = y
            last = y
            gap = 0
        elif start is not None:
            gap += 1
            if gap > gap_limit:
                runs.append((start, int(last), float(text_rows[start:last + 1].sum())))
                start = None
    if start is not None:
        runs.append((start, int(last), float(text_rows[start:last + 1].sum())))

    top_runs = [r for r in runs if r[0] < 0.5 * h] or runs
    if top_runs:
        densest = max(r[2] for r in top_runs)
        band_start, band_end, _ = next(
            r for r in top_runs if r[2] >= densest * 0.5)
    else:
        band_start, band_end = top_row, top_row
    title_top = band_start / h
    title_h = max(1, band_end - band_start + 1) / h

    bbox_area = ((right_col - left_col + 1) / w) * ((bottom_row - top_row + 1) / h)
    text_area = _clamp(bbox_area - image_area, 0.0, 1.0)

    # 栏数：标题带以下的文字列投影里最长空档（≥6% 内容宽且两侧都有墨迹）
    columns = 1
    body = text_ink[band_end + 1:bottom_row + 1, left_col:right_col + 1]
    if body.size and body.shape[0] >= 4:
        empty = body.mean(axis=0) <= 0.01
        best = 0
        run = 0
        best_end = 0
        for x, e in enumerate(empty):
            if e:
                run += 1
                if run > best:
                    best, best_end = run, x
            else:
                run = 0
        if best >= 0.06 * (right_col - left_col + 1) and 0 < best_end < len(empty) - 1:
            columns = 2

    # 对齐：标题带墨迹的左右边界是否居中且不占满
    band_cols = np.where(text_ink[band_start:band_end + 1, :].any(axis=0))[0]
    alignment = "left"
    if band_cols.size >= 2:
        b_left, b_right = int(band_cols[0]), int(band_cols[-1])
        center = (b_left + b_right) / 2.0 / w
        span = (b_right - b_left) / w
        if abs(center - 0.5) <= 0.04 and span <= 0.7:
            alignment = "center"

    confidence = 0.35
    if int(ink.sum()) > 0.01 * h * w:
        confidence += 0.2
    if band_end > top_row:
        confidence += 0.2
    if image_area > 0.01:
        confidence += 0.1
    if columns == 2:
        confidence += 0.05

    return _layout_block({
        "title_top_ratio": round(title_top, 3),
        "title_height_ratio": round(title_h, 3),
        "margin_left_ratio": round(left_col / w, 3),
        "margin_right_ratio": round(1 - (right_col + 1) / w, 3),
        "margin_top_ratio": round(top_row / h, 3),
        "margin_bottom_ratio": round(1 - (bottom_row + 1) / h, 3),
        "image_area_ratio": round(image_area, 3),
        "text_area_ratio": round(text_area, 3),
        "alignment": alignment,
        "columns": columns,
    }, "image", min(0.9, confidence))


def analyze_image(image_path: str | Path, *, name: str = "", source: str = "") -> dict:
    """从一张幻灯片图片估计版式骨架（png/jpg/webp，纯 PIL，无 OCR/模型）。

    返回与 analyze() 同结构的档案；只有 layout 有意义（配色/字体/密度保持默认，
    不做文字识别）。坏图片返回默认档案，不抛异常。
    """
    path = Path(image_path)
    profile = _default_profile()
    profile["name"] = (name or path.stem or "未命名图片").strip()
    profile["source"] = (source or path.name or "image.png").strip()
    if path.is_file():
        profile["_source_path"] = str(path.resolve())
    try:
        profile["layout"] = _image_layout(path)
    except Exception:
        profile["layout"] = dict(DEFAULT_LAYOUT)
    profile["style"] = dict(DEFAULT_STYLE)
    profile["learn"] = {"rounds": 0, "history": []}
    return profile


def _default_profile() -> dict:
    return {
        "id": "",
        "name": "",
        "source": "",
        "created_at": _now(),
        "slide_size": [13.33, 7.5],
        "palette": dict(DEFAULT_PALETTE),
        "fonts": dict(DEFAULT_FONTS),
        "metrics": dict(DEFAULT_METRICS),
        "layout": dict(DEFAULT_LAYOUT),
        "style": dict(DEFAULT_STYLE),
        "learn": {"rounds": 0, "history": []},
    }


def analyze(pptx_path: str | Path, *, name: str = "", source: str = "") -> dict:
    """分析一个 .pptx，返回风格档案。任何情况下都不抛异常（最差返回默认档案）。"""
    path = Path(pptx_path)
    profile = _default_profile()
    profile["name"] = (name or path.stem or "未命名案例").strip()
    profile["source"] = (source or path.name or "source.pptx").strip()
    if path.is_file():
        profile["_source_path"] = str(path.resolve())

    try:
        from pptx import Presentation

        with zipfile.ZipFile(path) as zf:      # 先确认是合法 zip，避免 python-pptx 报怪错
            zf.namelist()
        prs = Presentation(str(path))
        obs = _observe(prs)
        theme_colors, theme_fonts = _parse_theme(
            _find_theme_root(path), prefer_cjk=obs["has_cjk"])

        profile["slide_size"] = [
            round(obs["slide_w"], 2),
            round(obs["slide_h"], 2),
        ]
        if not obs["slides"]:
            # 合法但一页都没有：没有可分析的证据，保持默认档案
            return profile
        profile["palette"] = _build_palette(obs, theme_colors)
        profile["fonts"] = _build_fonts(obs, theme_fonts)
        profile["metrics"] = _build_metrics(obs)
        profile["layout"] = _deck_layout(prs)
    except Exception:
        # 损坏 / 加密 / 空文件：退回默认档案，但保留名字与时间
        pass

    profile["style"] = dict(DEFAULT_STYLE)
    profile["learn"] = {"rounds": 0, "history": []}
    return profile


# ------------------------------------------------------------------ 档案存取


def _validate_id(profile_id: str) -> str:
    pid = str(profile_id or "").strip()
    if not _ID_RE.match(pid) or ".." in pid:
        raise FileNotFoundError(f"非法的档案 id：{profile_id!r}")
    return pid


def _new_id() -> str:
    """档案 id：pf-YYYYMMDD-<8 位随机 hex>（评审 F9：4 位 hex 会碰撞覆盖旧档案）。"""
    return "pf-%s-%s" % (datetime.now().strftime("%Y%m%d"), os.urandom(4).hex())


def _unique_id(store: Path) -> str:
    """生成一个 store 中尚未占用的 id。"""
    for _ in range(100):
        pid = _new_id()
        if not _profile_path(store, pid).exists():
            return pid
    raise RuntimeError("无法分配唯一的档案 id")


def _public(profile: dict) -> dict:
    """去掉内部私有字段（下划线开头）后的档案副本。"""
    out = {}
    for key, value in (profile or {}).items():
        if isinstance(key, str) and key.startswith("_"):
            continue
        out[key] = copy.deepcopy(value)
    return out


def _profile_path(store: Path, profile_id: str) -> Path:
    return store / "profiles" / f"{profile_id}.json"


def save(profile: dict, store_dir: str | Path) -> dict:
    """落盘：profiles/<id>.json + cases/<id>/source.pptx。

    profile 里的 `_source_path`（analyze 写入）或 `source` 指向的真实文件会被拷贝成
    cases/<id>/source.pptx；用户文件名只作展示，绝不参与路径拼接。
    """
    if not isinstance(profile, dict):
        raise TypeError("profile 必须是 dict")
    store = Path(store_dir)
    (store / "profiles").mkdir(parents=True, exist_ok=True)

    pf = _public(profile)
    pid = str(pf.get("id") or "").strip()
    source_is_new = bool(profile.get("_source_path"))
    if not pid or not _ID_RE.match(pid) or ".." in pid:
        pid = _unique_id(store)
    elif source_is_new and _profile_path(store, pid).exists():
        # 新案例与已有档案撞 id（同日多次上传）：重新生成，绝不覆盖旧档案。
        pid = _unique_id(store)
    pf["id"] = pid
    pf.setdefault("created_at", _now())
    pf.setdefault("name", "")
    pf.setdefault("source", "")
    pf.setdefault("slide_size", [13.33, 7.5])
    pf["palette"] = {**DEFAULT_PALETTE, **(pf.get("palette") or {})}
    pf["fonts"] = {**DEFAULT_FONTS, **(pf.get("fonts") or {})}
    pf["metrics"] = {**DEFAULT_METRICS, **(pf.get("metrics") or {})}
    pf["layout"] = {**DEFAULT_LAYOUT, **(pf.get("layout") or {})}
    pf["style"] = {**DEFAULT_STYLE, **(pf.get("style") or {})}
    learn_obj = pf.get("learn") if isinstance(pf.get("learn"), dict) else {}
    pf["learn"] = {
        "rounds": int(learn_obj.get("rounds") or 0),
        "history": list(learn_obj.get("history") or [])[-HISTORY_LIMIT:],
    }

    source_file = profile.get("_source_path")
    if not source_file:
        candidate = str(profile.get("source") or "")
        if candidate and Path(candidate).is_file():
            source_file = candidate
    if source_file and Path(source_file).is_file():
        case_dir = store / "cases" / pid
        case_dir.mkdir(parents=True, exist_ok=True)
        suffix = Path(source_file).suffix.lower()
        if suffix not in (".pptx", ".png", ".jpg", ".jpeg", ".webp", ".bmp"):
            suffix = ".pptx"
        target_name = f"source{suffix}"
        try:
            shutil.copyfile(Path(source_file), case_dir / target_name)
            pf["source"] = target_name
        except OSError:
            pass

    path = _profile_path(store, pid)
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(json.dumps(pf, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, path)
    return pf


def load(store_dir: str | Path, profile_id: str) -> dict:
    store = Path(store_dir)
    pid = _validate_id(profile_id)
    path = _profile_path(store, pid)
    if not path.is_file():
        raise FileNotFoundError(f"找不到风格档案：{pid}")
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError(f"档案文件格式不对：{path}")
    data.setdefault("id", pid)
    data["palette"] = {**DEFAULT_PALETTE, **(data.get("palette") or {})}
    data["fonts"] = {**DEFAULT_FONTS, **(data.get("fonts") or {})}
    data["metrics"] = {**DEFAULT_METRICS, **(data.get("metrics") or {})}
    data["layout"] = {**DEFAULT_LAYOUT, **(data.get("layout") or {})}
    data["style"] = {**DEFAULT_STYLE, **(data.get("style") or {})}
    if not isinstance(data.get("learn"), dict):
        data["learn"] = {"rounds": 0, "history": []}
    data["learn"].setdefault("rounds", 0)
    data["learn"].setdefault("history", [])
    return data


def list_all(store_dir: str | Path) -> list[dict]:
    store = Path(store_dir)
    root = store / "profiles"
    if not root.is_dir():
        return []
    out: list[dict] = []
    for path in root.glob("*.json"):
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            if not isinstance(data, dict):
                continue
            data.setdefault("id", path.stem)
            data["palette"] = {**DEFAULT_PALETTE, **(data.get("palette") or {})}
            data["fonts"] = {**DEFAULT_FONTS, **(data.get("fonts") or {})}
            data["metrics"] = {**DEFAULT_METRICS, **(data.get("metrics") or {})}
            data["layout"] = {**DEFAULT_LAYOUT, **(data.get("layout") or {})}
            data["style"] = {**DEFAULT_STYLE, **(data.get("style") or {})}
            out.append(data)
        except (OSError, ValueError):
            continue
    out.sort(key=lambda p: (str(p.get("created_at", "")), str(p.get("id", ""))), reverse=True)
    return out


# ------------------------------------------------------------------ 档案 → 主题


def _hex_or(value, fallback: str) -> str:
    return _to_hex(value) or fallback


def apply(profile: dict, base_theme: str | None = None) -> Theme:
    """用档案的配色/字体/标题比例覆盖一个内置主题，返回新的 Theme 实例。"""
    profile = profile or {}
    palette = {**DEFAULT_PALETTE, **(profile.get("palette") or {})}
    fonts = {**DEFAULT_FONTS, **(profile.get("fonts") or {})}
    style = {**DEFAULT_STYLE, **(profile.get("style") or {})}

    if isinstance(base_theme, Theme):
        base = base_theme
    else:
        base_name = (base_theme or profile.get("base_theme")
                     or profile.get("theme") or None)
        base = get_theme(base_name)

    try:
        title_scale = float(style.get("title_scale") or 1.0)
    except (TypeError, ValueError):
        title_scale = 1.0
    title_scale = _clamp(title_scale, *TITLE_SCALE_RANGE)
    title_size = int(round(base.title_size * title_scale))
    title_size = max(16, min(72, title_size))

    font_title = str(fonts.get("title") or base.font_title)
    font_body = str(fonts.get("body") or base.font_body)
    font_display = font_title if font_title != base.font_title else base.font_display

    return clone_with(
        base,
        bg=rgb(_hex_or(palette.get("bg"), "FFFFFF")),
        ink=rgb(_hex_or(palette.get("ink"), "14213D")),
        muted=rgb(_hex_or(palette.get("muted"), "5C6B82")),
        accent=rgb(_hex_or(palette.get("accent"), "1B3A6B")),
        accent_soft=rgb(_hex_or(palette.get("accent_soft"), "E9EEF7")),
        rule=rgb(_hex_or(palette.get("rule"), "D5DCE8")),
        font_title=font_title,
        font_body=font_body,
        font_display=font_display,
        title_size=title_size,
    )


# ------------------------------------------------------------------ 反馈调参


def learn(profile: dict, rating: int, notes: str = "") -> dict:
    """按契约第 7 节就地调整 density / title_scale，并追加一条 history。"""
    if not isinstance(profile, dict):
        raise TypeError("profile 必须是 dict")
    try:
        rating = int(rating)
    except (TypeError, ValueError):
        raise ValueError("rating 必须是 1–5 的整数")
    rating = max(1, min(5, rating))

    style = profile.setdefault("style", dict(DEFAULT_STYLE))
    if not isinstance(style, dict):
        style = dict(DEFAULT_STYLE)
        profile["style"] = style
    try:
        density = float(style.get("density") or 1.0)
    except (TypeError, ValueError):
        density = 1.0
    try:
        title_scale = float(style.get("title_scale") or 1.0)
    except (TypeError, ValueError):
        title_scale = 1.0

    if rating >= 4:
        density *= 1.03
    elif rating <= 2:
        density *= 0.90
        title_scale *= 0.97
    # rating == 3：不变

    density = _clamp(round(density, 4), *DENSITY_RANGE)
    title_scale = _clamp(round(title_scale, 4), *TITLE_SCALE_RANGE)
    style["density"] = density
    style["title_scale"] = title_scale
    style.setdefault("bullet_char", DEFAULT_STYLE["bullet_char"])

    learn_obj = profile.get("learn")
    if not isinstance(learn_obj, dict):
        learn_obj = {"rounds": 0, "history": []}
        profile["learn"] = learn_obj
    try:
        learn_obj["rounds"] = int(learn_obj.get("rounds") or 0) + 1
    except (TypeError, ValueError):
        learn_obj["rounds"] = 1
    history = learn_obj.get("history")
    if not isinstance(history, list):
        history = []
        learn_obj["history"] = history
    history.append({
        "at": _now(),
        "rating": rating,
        "density": density,
        "title_scale": title_scale,
        "notes": str(notes or ""),
    })
    if len(history) > HISTORY_LIMIT:
        # 只保留最近 HISTORY_LIMIT 条；rounds 是累计计数，不清零。
        del history[:-HISTORY_LIMIT]
    return profile
