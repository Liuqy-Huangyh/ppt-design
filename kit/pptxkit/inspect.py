"""PPTX 体检 + 反解成 Markdown + 渲染后视觉评审。"""
from __future__ import annotations

import argparse
import re
import sys
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


# ---------------------------------------------------------------- 渲染评审

REVIEW_MIN_FONT_PT = 14.0
REVIEW_MIN_BODY_LEN = 40          # 长文本才按“正文”审最小字号
REVIEW_FOOTER_MARGIN = 0.55       # 页脚区：top > H - 0.55in
REVIEW_CAPTION_ZONE = 1.35        # 底部说明/版权区：top > H - 1.35in
REVIEW_OVERFLOW_RATIO = 1.15
REVIEW_OVERFLOW_MIN_LEN = 24
REVIEW_OVERLAP_RATIO = 0.35       # 交集 / 较小面积
REVIEW_OVERLAP_MIN_AREA = 0.35    # 参与判定的最小形状面积（平方英寸）
REVIEW_BOUNDS_TOLERANCE = 0.06    # 越界容差（英寸）
REVIEW_BLANK_COVERAGE = 0.40      # 形状覆盖 < 40% 视为 >60% 空白
REVIEW_BLANK_MAX_CHARS = 24
REVIEW_IMAGE_STRETCH = 0.12       # 宽高比偏差超过 12% 视为拉伸
REVIEW_FOCUS_MIN_INK = 0.002      # 渲染墨迹 < 0.2% 视为没有视觉焦点
REVIEW_RENDER_BLANK_BBOX = 0.10
REVIEW_MAX_ISSUES_PER_PAGE = 8
_PAGE_NUM_RE = re.compile(r"(\d+)")


def _inches(value) -> float:
    return (value or 0) / 914400


def _bbox(shape) -> tuple[float, float, float, float]:
    return (_inches(shape.left), _inches(shape.top),
            _inches(shape.width), _inches(shape.height))


def _is_footer_shape(shape, prs) -> bool:
    top = _inches(shape.top)
    return top > _inches(prs.slide_height) - REVIEW_FOOTER_MARGIN


def _in_caption_zone(shape, prs) -> bool:
    top = _inches(shape.top)
    return top > _inches(prs.slide_height) - REVIEW_CAPTION_ZONE


def _text_sizes(shape) -> list[float]:
    sizes = []
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            if run.font.size:
                sizes.append(float(run.font.size.pt))
    return sizes


def _coverage(slide, width_in: float, height_in: float,
              step: float = 0.1) -> float:
    """形状包围盒并集占整页的比例（网格近似，用于留白判定）。"""
    cols = max(1, int(round(width_in / step)))
    rows = max(1, int(round(height_in / step)))
    cells: set[tuple[int, int]] = set()
    for shape in slide.shapes:
        left, top, w, h = _bbox(shape)
        x0 = max(0, int(left / step))
        y0 = max(0, int(top / step))
        x1 = min(cols, int((left + w) / step + 0.5))
        y1 = min(rows, int((top + h) / step + 0.5))
        for x in range(x0, x1):
            for y in range(y0, y1):
                cells.add((x, y))
    return len(cells) / (cols * rows)


def _overlap_issues(content: list[dict]) -> list[dict]:
    issues = []
    for i in range(len(content)):
        for j in range(i + 1, len(content)):
            a, b = content[i], content[j]
            if a["area"] < REVIEW_OVERLAP_MIN_AREA or b["area"] < REVIEW_OVERLAP_MIN_AREA:
                continue
            ax, ay, aw, ah = a["bbox"]
            bx, by, bw, bh = b["bbox"]
            ix = max(0.0, min(ax + aw, bx + bw) - max(ax, bx))
            iy = max(0.0, min(ay + ah, by + bh) - max(ay, by))
            inter = ix * iy
            if inter <= 0:
                continue
            ratio = inter / min(a["area"], b["area"])
            if ratio > REVIEW_OVERLAP_RATIO:
                issues.append({
                    "gate": "defect", "kind": "元素重叠",
                    "detail": f"「{a['label']}」与「{b['label']}」重叠 {ratio:.0%}",
                })
    return issues


def _image_issues(slide) -> list[dict]:
    issues = []
    for shape in slide.shapes:
        if not getattr(shape, "image", None):
            continue
        try:
            image = shape.image
            pixel_w, pixel_h = image.size
        except Exception:
            continue
        if not pixel_w or not pixel_h:
            continue
        _, _, w, h = _bbox(shape)
        if not w or not h:
            continue
        # 有意的裁切（crop）不算拉伸
        try:
            if any(float(getattr(shape, f"crop_{side}", 0) or 0) for side in
                   ("left", "right", "top", "bottom")):
                continue
        except Exception:
            pass
        displayed = w / h
        original = pixel_w / pixel_h
        stretch = abs(displayed - original) / original
        if stretch > REVIEW_IMAGE_STRETCH:
            issues.append({
                "gate": "defect", "kind": "图片拉伸",
                "detail": f"显示比例 {displayed:.2f} 与原始 {original:.2f} 相差 {stretch:.0%}",
            })
    return issues


def _structural_issues(prs, slide, width_in: float, height_in: float) -> tuple[list[dict], dict]:
    issues: list[dict] = []
    content: list[dict] = []
    chars = 0
    images = 0
    max_pt = 0.0
    for shape in slide.shapes:
        has_text = shape.has_text_frame and bool(shape.text_frame.text.strip())
        is_picture = bool(getattr(shape, "image", None))
        is_table = bool(getattr(shape, "has_table", False) and shape.has_table)
        if is_picture:
            images += 1
        if getattr(shape, "shape_type", None) == 13:
            images = max(images, 1)
        if not has_text and not is_picture and not is_table:
            continue
        left, top, w, h = _bbox(shape)
        text = ""
        if has_text:
            text = shape.text_frame.text.strip()
            if not _is_footer_shape(shape, prs) and not FOOTER_RE.match(text):
                chars += len(text)
        label = (text[:12].replace("\n", " ") if text else
                 ("图片" if is_picture else "表格"))

        # 1) 文字越界
        if has_text and (left < -REVIEW_BOUNDS_TOLERANCE
                         or top < -REVIEW_BOUNDS_TOLERANCE
                         or left + w > width_in + REVIEW_BOUNDS_TOLERANCE
                         or top + h > height_in + REVIEW_BOUNDS_TOLERANCE):
            issues.append({
                "gate": "defect", "kind": "文字越界",
                "detail": f"「{label}」超出页面边界",
            })

        # 2) 文字溢出（基于形状边界 + 估行）
        if has_text and len(text) >= REVIEW_OVERFLOW_MIN_LEN:
            ratio = _fit_ratio(shape)
            if ratio > REVIEW_OVERFLOW_RATIO:
                issues.append({
                    "gate": "defect", "kind": "文字溢出",
                    "detail": f"「{label}…」需要 {ratio:.1f} 倍高度",
                })

        # 3) 最小字号（只看正文式长文本，页脚/底部说明不算）
        if has_text and not _is_footer_shape(shape, prs) and not _in_caption_zone(shape, prs):
            sizes = _text_sizes(shape)
            if sizes:
                max_pt = max(max_pt, max(sizes))
                body_like = (len(text) >= REVIEW_MIN_BODY_LEN
                             or len(shape.text_frame.paragraphs) >= 2
                             or (h >= 0.55 and len(text) >= 20))
                if body_like and min(sizes) < REVIEW_MIN_FONT_PT:
                    issues.append({
                        "gate": "defect", "kind": "字号过小",
                        "detail": f"「{label}」最小 {min(sizes):.1f}pt（建议 ≥{REVIEW_MIN_FONT_PT:.0f}pt）",
                    })

        # 4) 标题/正文层级
        if has_text and len(text) >= 6:
            content.append({"bbox": (left, top, w, h), "area": w * h, "label": label})

    # 5) 元素重叠
    issues.extend(_overlap_issues(content))
    # 6) 图片拉伸
    issues.extend(_image_issues(slide))

    # 7) 页内留白异常（>60% 空白且无图）
    coverage = _coverage(slide, width_in, height_in)
    if images == 0 and chars < REVIEW_BLANK_MAX_CHARS and coverage < REVIEW_BLANK_COVERAGE:
        issues.append({
            "gate": "visual", "kind": "留白过多",
            "detail": f"内容仅覆盖 {coverage:.0%}，>60% 空白且无图",
        })
    # 8) 标题与正文层级成立（只在有较长正文时把关）
    if chars >= 40 and max_pt and max_pt < 18:
        issues.append({
            "gate": "visual", "kind": "标题层级弱",
            "detail": f"全页最大字号仅 {max_pt:.1f}pt，标题与正文拉不开",
        })

    stats = {"chars": chars, "images": images, "coverage": round(coverage, 3),
             "max_pt": round(max_pt, 1)}
    return issues, stats


def _load_png_metrics(path: Path) -> dict | None:
    """用 Pillow 读渲染图：墨迹占比 + 内容包围盒占比。"""
    try:
        from PIL import Image
    except ImportError:
        return None
    try:
        image = Image.open(path)
        if image.mode in ("RGBA", "LA", "P"):
            image = image.convert("RGBA")
            canvas = Image.new("RGBA", image.size, (255, 255, 255, 255))
            image = Image.alpha_composite(canvas, image).convert("RGB")
        else:
            image = image.convert("RGB")
    except Exception:
        return None
    width, height = image.size
    if not width or not height:
        return None
    small = image.resize((160, 90))
    colors = small.getcolors(maxcolors=160 * 90)
    if not colors:
        return None
    background = max(colors, key=lambda item: item[0])[1]

    def diff(color) -> int:
        return (abs(color[0] - background[0]) + abs(color[1] - background[1])
                + abs(color[2] - background[2]))

    pixel = image.load()
    step = 4
    ink = 0
    total = 0
    x0, y0, x1, y1 = width, height, 0, 0
    for y in range(0, height, step):
        for x in range(0, width, step):
            total += 1
            if diff(pixel[x, y]) > 24:
                ink += 1
                if x < x0:
                    x0 = x
                if y < y0:
                    y0 = y
                if x > x1:
                    x1 = x
                if y > y1:
                    y1 = y
    ink_ratio = ink / total if total else 0.0
    bbox_ratio = ((x1 - x0) * (y1 - y0)) / (width * height) if ink else 0.0
    return {"width": width, "height": height,
            "ink": round(ink_ratio, 4), "bbox": round(bbox_ratio, 4)}


def _page_pngs(png_dir) -> dict[int, Path]:
    """把渲染目录里的 PNG 按页码排序：slide-01.png / 幻灯片1.PNG 都认。"""
    if not png_dir:
        return {}
    folder = Path(png_dir)
    if not folder.is_dir():
        return {}
    files = []
    for path in folder.glob("*.png"):
        match = _PAGE_NUM_RE.search(path.stem)
        number = int(match.group(1)) if match else 0
        files.append((number, path.name.lower(), path))
    files.sort()
    return {index: path for index, (_num, _name, path) in enumerate(files, start=1)}


def _render_issues(metrics: dict, stats: dict) -> list[dict]:
    issues = []
    if metrics["ink"] < REVIEW_FOCUS_MIN_INK:
        issues.append({
            "gate": "visual", "kind": "无视觉焦点",
            "detail": f"渲染后墨迹仅占 {metrics['ink']:.2%}",
        })
    if (stats["images"] == 0 and stats["chars"] < REVIEW_BLANK_MAX_CHARS
            and metrics["bbox"] < REVIEW_RENDER_BLANK_BBOX):
        issues.append({
            "gate": "visual", "kind": "留白过多",
            "detail": f"渲染内容仅占版面 {metrics['bbox']:.0%}，且无图",
        })
    return issues


def _dedupe(issues: list[dict]) -> list[dict]:
    """同一页同类问题只保留一条（细节取首次出现），避免渲染与结构重复报。"""
    seen = set()
    out = []
    for issue in issues:
        key = issue.get("kind")
        if key in seen:
            continue
        seen.add(key)
        out.append(issue)
    return out[:REVIEW_MAX_ISSUES_PER_PAGE]


def review(path: str | Path, *, png_dir=None) -> dict:
    """两闸门逐页评审：视觉（焦点/铺满/层级）+ 缺陷（越界溢出/重叠/字号/拉伸）。

    png_dir 提供渲染后的逐页 PNG 时叠加渲染分析；没有 PNG 就只做结构检查，
    并在 notes 里明确写「未渲染」。
    """
    path = Path(path)
    prs = Presentation(str(path))
    width_in = _inches(prs.slide_width)
    height_in = _inches(prs.slide_height)
    pngs = _page_pngs(png_dir)
    notes = []
    if pngs:
        notes.append(f"已分析 {len(pngs)} 张渲染 PNG")
    else:
        notes.append("未渲染（缺少 Office/WPS）：仅执行结构检查")

    pages = []
    for index, slide in enumerate(prs.slides, start=1):
        issues, stats = _structural_issues(prs, slide, width_in, height_in)
        render_metrics = None
        png = pngs.get(index)
        if png is not None:
            render_metrics = _load_png_metrics(png)
            if render_metrics is None:
                notes.append("渲染 PNG 无法解析：该页只看结构")
            else:
                issues.extend(_render_issues(render_metrics, stats))
        issues = _dedupe(issues)
        pages.append({
            "page": index,
            "status": "PASS" if not issues else "NEEDS_REVISION",
            "issues": issues,
            "render": render_metrics,
            "stats": stats,
        })
    pass_count = sum(1 for page in pages if page["status"] == "PASS")
    return {
        "file": str(path),
        "size_in": (round(width_in, 2), round(height_in, 2)),
        "rendered": bool(pngs),
        "pages": pages,
        "notes": notes,
        "pass": pass_count,
        "needs": len(pages) - pass_count,
        "total": len(pages),
    }


def format_review(report: dict) -> str:
    width_in, height_in = report["size_in"]
    lines = [
        f"文件：{report['file']}",
        f"画布：{width_in} × {height_in} in ｜ 共 {report['total']} 页",
    ]
    if report["rendered"]:
        lines.append("渲染：已分析逐页 PNG（视觉闸门生效）")
    else:
        lines.append("渲染：未渲染（缺少 Office/WPS）：仅执行结构检查")
    lines.append("")
    for page in report["pages"]:
        reasons = "；".join(
            f"{issue['kind']}（{issue['detail']}）" if issue.get("detail") else issue["kind"]
            for issue in page["issues"])
        suffix = f"：{reasons}" if reasons else ""
        lines.append(f"[{page['status']}] 第 {page['page']} 页{suffix}")
    lines.append("")
    lines.append(f"汇总：PASS {report['pass']} / NEEDS_REVISION {report['needs']}"
                 f"（共 {report['total']} 页）")
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="python -m pptxkit.inspect",
                                     description="PPTX 体检：结构 + 渲染后评审")
    parser.add_argument("--visual", action="store_true", help="逐页视觉/缺陷评审")
    parser.add_argument("file", nargs="?", help="待评审的 .pptx")
    parser.add_argument("--png-dir", default=None,
                        help="render-pptx.ps1 导出的逐页 PNG 目录（可选）")
    args = parser.parse_args(argv)
    if not args.visual or not args.file:
        parser.print_help()
        return 2
    try:
        report = review(args.file, png_dir=args.png_dir)
    except Exception as exc:
        print(f"评审失败：{exc}", file=sys.stderr)
        return 2
    print(format_review(report))
    return 0 if report["needs"] == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())

