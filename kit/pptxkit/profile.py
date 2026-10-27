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
        try:
            shutil.copyfile(Path(source_file), case_dir / "source.pptx")
            pf["source"] = "source.pptx"
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
