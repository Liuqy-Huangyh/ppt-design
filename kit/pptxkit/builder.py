"""把 Markdown / JSON 描述渲染成 .pptx。"""
from __future__ import annotations

import json
import re
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches

from . import markdown as md
from .layouts import render
from .themes import SLIDE_H, SLIDE_W, Theme, get_theme

_DRIVE_RE = re.compile(r"^[A-Za-z]:")


def _new_prs() -> Presentation:
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    return prs


def _resolve(path: str, base: Path, asset_root: Path | None = None) -> str:
    """把 Markdown 里的图片路径解析成绝对路径，且只允许落在可信目录内。

    安全口径（评审 F1）：拒绝绝对路径、盘符、UNC 与 `..` 逃逸；用户提供的
    outline 不能被用来读取服务端任意文件。允许的范围：
    * `base` 目录内（正常相对路径）；
    * 可选的 `asset_root`（CLI 场景下工作区 `assets/` 白名单，供模板用
      `../assets/...` 引用共享图片）。
    越界时抛 ValueError，由调用方决定如何呈现（服务端会先校验并返回
    400 invalid_outline）。
    """
    raw = str(path or "").strip()
    if not raw or "\x00" in raw:
        raise ValueError("图片路径为空或包含非法字符")
    normalised = raw.replace("\\", "/")
    if (normalised.startswith("/") or normalised.startswith("//")
            or _DRIVE_RE.match(normalised)):
        raise ValueError(f"不允许绝对路径图片：{raw}")
    root = Path(base).resolve()
    candidate = (root / normalised).resolve()
    if candidate == root or root in candidate.parents:
        return str(candidate)
    if asset_root is not None:
        allowed = Path(asset_root).resolve()
        if candidate == allowed or allowed in candidate.parents:
            return str(candidate)
    raise ValueError(f"图片路径越出 base_dir：{raw}")



def build(doc: dict, out: str | Path, *, theme: str | None = None,
          theme_obj: Theme | None = None,
          skeleton: dict | None = None,
          base_dir: str | Path | None = None,
          asset_root: str | Path | None = None) -> Path:
    """doc = {"meta": {...}, "slides": [...]}

    theme_obj 优先于 theme 名：风格档案（profile.apply()）返回的 Theme 实例
    可以直接传进来，配色/字体/标题比例都从案例里带过来。
    skeleton 是 profile.layout_hints() 的结果：只微调 cover/section/bullets/image
    的标题纵向位置与左右边距（已限幅）；不传时输出与旧版逐像素一致。
    asset_root 是额外的图片白名单目录（CLI 用它允许模板引用工作区 assets/）。
    """
    meta = doc.get("meta") or {}
    slides = doc.get("slides") or []
    th = theme_obj if isinstance(theme_obj, Theme) else get_theme(theme or meta.get("theme"))
    base = Path(base_dir or ".")

    # 有 front matter 标题时，自动补一张封面
    all_slides: list[dict] = []
    if meta.get("title"):
        author = meta.get("author", "")
        all_slides.append({
            "layout": "cover",
            "title": meta["title"],
            "subtitle": meta.get("subtitle", ""),
            "kicker": meta.get("kicker") or meta.get("type") or "",
            "meta": {
                "author": author,
                "org": meta.get("org", ""),
                "date": meta.get("date", ""),
            },
        })
    all_slides.extend(slides)
    if not all_slides:
        raise SystemExit("没有可渲染的页面：请在 Markdown 里写 # 标题")

    # 图片路径相对 Markdown 文件解析
    for s in all_slides:
        if s.get("image") and s["image"].get("path"):
            s["image"]["path"] = _resolve(s["image"]["path"], base, asset_root)
        # 版式骨架（可选）：layouts 只在 cover/section/bullets/image 里读它
        if skeleton:
            s["_skeleton"] = skeleton

    footer = meta.get("footer") or " · ".join(
        [x for x in (meta.get("author", ""), meta.get("org", "")) if x]
    )

    prs = _new_prs()
    total = len(all_slides)
    for i, data in enumerate(all_slides, start=1):
        render(prs, th, data, i, total, footer)

    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    try:
        prs.save(str(out))
    except PermissionError as exc:
        raise SystemExit(
            f"无法写入 {out}：文件可能正被 PowerPoint / WPS 打开。关闭后重试，或用 -o 指定另一个文件名。"
        ) from exc
    return out


def build_file(src: str | Path, out: str | Path | None = None, *,
               theme: str | None = None,
               theme_obj: Theme | None = None,
               skeleton: dict | None = None) -> Path:
    src = Path(src)
    if not src.exists():
        raise SystemExit(f"找不到文件：{src}")
    if src.suffix.lower() == ".json":
        doc = json.loads(src.read_text(encoding="utf-8"))
        slides = doc.get("slides") or []
        md._infer_layouts(slides, doc.get("meta") or {})
        data_dir = src.parent
    else:
        text = src.read_text(encoding="utf-8")
        doc = md.parse(text)
        data_dir = src.parent
    out = Path(out) if out else src.with_name(src.stem + ".pptx")
    # 工作区模板会用 ../assets/ 引用共享图片：只把同级 assets/ 加进白名单，
    # 不让 `..` 变成通用的目录逃逸（评审 F1）。
    asset_root = data_dir.parent / "assets"
    if not asset_root.is_dir():
        asset_root = None
    return build(doc, out, theme=theme, theme_obj=theme_obj, skeleton=skeleton,
                 base_dir=data_dir, asset_root=asset_root)
