"""Markdown → 单文件 HTML 演示（零依赖、可离线、可直接丢到任何静态托管上）。

键盘：← → 空格 翻页 / O 总览 / S 讲者备注 / F 全屏 / P 打印成 PDF / ? 帮助
"""
from __future__ import annotations

import html
import re
from pathlib import Path

from . import markdown as md
from .themes import Theme, get_theme, rgb

_INLINE = re.compile(r"(\*\*.+?\*\*|`.+?`|\*[^*]+\*)")

# 写进 <style> 的字体名只接受白名单字符；颜色必须是 #RRGGBB。
# 档案来自用户上传的 pptx，字体名是攻击者可控数据（评审 F10）。
SAFE_FONT_RE = re.compile(r"^[A-Za-z0-9 _\-\u4e00-\u9fff]{1,64}$")
HEX_COLOR_RE = re.compile(r"^#[0-9A-Fa-f]{6}$")
_DRIVE_RE = re.compile(r"^[A-Za-z]:")
ALLOWED_IMAGE_MODES = ("relative", "placeholder")
FALLBACK_FONT = "Microsoft YaHei"


def _safe_font(name, fallback: str = FALLBACK_FONT) -> str:
    text = str(name or "").strip()
    if SAFE_FONT_RE.fullmatch(text):
        return text
    return fallback


def _safe_color(value, fallback: str) -> str:
    try:
        text = _hex(value)
    except Exception:
        return fallback
    return text if HEX_COLOR_RE.fullmatch(text) else fallback


def _safe_image_src(path, mode: str = "relative") -> str | None:
    """图片引用只保留浏览器可用的相对引用；绝对路径/UNC/盘符/scheme 降级为占位。

    网页版不会去读服务端文件系统，所以 `../assets/...` 这类相对引用可以保留
    （CLI 生成的 html 放在 out/ 里仍能显示工作区图片）；只要不回显本机绝对路径。
    服务端生成 deck 时传 placeholder，避免 deck 目录里出现失效链接。
    """
    if mode not in ALLOWED_IMAGE_MODES:
        mode = "relative"
    if mode == "placeholder":
        return None
    raw = str(path or "").strip()
    if not raw or "\x00" in raw:
        return None
    normalised = raw.replace("\\", "/")
    if (normalised.startswith("/") or normalised.startswith("//")
            or _DRIVE_RE.match(normalised)):
        return None
    parts = [part for part in normalised.split("/") if part not in ("", ".")]
    if not parts:
        return None
    if ":" in parts[0]:          # 拒绝 data:/javascript: 等 scheme
        return None
    return normalised


def inline(text: str) -> str:
    out = []
    for chunk in _INLINE.split(text):
        if not chunk:
            continue
        esc = html.escape(chunk)
        if chunk.startswith("**") and chunk.endswith("**") and len(chunk) > 4:
            out.append(f"<strong>{html.escape(chunk[2:-2])}</strong>")
        elif chunk.startswith("`") and chunk.endswith("`") and len(chunk) > 2:
            out.append(f"<code>{html.escape(chunk[1:-1])}</code>")
        elif chunk.startswith("*") and chunk.endswith("*") and len(chunk) > 2:
            out.append(f"<em>{html.escape(chunk[1:-1])}</em>")
        else:
            out.append(esc)
    return "".join(out)


def _hex(color) -> str:
    return "#%02X%02X%02X" % (color[0], color[1], color[2])


def _bullets_html(items, cls="bullets") -> str:
    lis = []
    for it in items:
        level = min(int(it.get("level", 0)), 2)
        marker = "" if it.get("bullet", True) else ' data-plain="1"'
        lis.append(f'<li class="lvl{level}"{marker}>{inline(it["text"])}</li>')
    return f'<ul class="{cls}">' + "".join(lis) + "</ul>"


def _slide_html(data: dict, idx: int, total: int, footer: str,
                image_mode: str = "relative") -> str:
    layout = data.get("layout") or "bullets"
    parts: list[str] = []
    kicker = data.get("kicker") or ""
    if layout == "cover":
        if kicker:
            parts.append(f'<div class="kicker">{inline(kicker)}</div>')
        parts.append(f'<h1 class="cover-title">{inline(data.get("title",""))}</h1>')
        parts.append('<div class="rule"></div>')
        if data.get("subtitle"):
            parts.append(f'<p class="cover-sub">{inline(data["subtitle"])}</p>')
        meta = data.get("meta") or {}
        lines = [meta.get(k) for k in ("author", "org", "date") if meta.get(k)]
        if lines:
            parts.append('<div class="cover-meta">' + "".join(
                f"<span>{inline(x)}</span>" for x in lines) + "</div>")
    elif layout == "section":
        num = str(data.get("number") or data.get("kicker") or "")
        parts.append(f'<div class="section-num">{inline(num)}</div>')
        if num:
            parts.append(f'<div class="kicker">PART {inline(num)}</div>')
        parts.append(f'<h1 class="section-title">{inline(data.get("title",""))}</h1>')
        if data.get("subtitle"):
            parts.append(f'<p class="section-sub">{inline(data["subtitle"])}</p>')
    elif layout == "table":
        tb = data.get("table") or {}
        head = "".join(f"<th>{inline(h)}</th>" for h in tb.get("header", []))
        rows = "".join(
            "<tr>" + "".join(f"<td>{inline(c)}</td>" for c in r) + "</tr>"
            for r in tb.get("rows", [])
        )
        parts.append(_title_block(data))
        parts.append(f"<table><thead><tr>{head}</tr></thead><tbody>{rows}</tbody></table>")
    elif layout == "image":
        img = data.get("image") or {}
        parts.append(_title_block(data))
        cap = img.get("caption", "")
        src = _safe_image_src(img.get("path", ""), image_mode)
        if src:
            body = f'<img src="{html.escape(src, quote=True)}" alt="{html.escape(cap)}">'
        else:
            body = '<div class="img-missing">（图片未随本网页提供）</div>'
        if cap:
            body += f'<figcaption>{inline(cap)}</figcaption>'
        side = ""
        if data.get("bullets"):
            side = '<div class="img-side">' + _bullets_html(data["bullets"]) + "</div>"
        parts.append(f'<figure class="img-wrap">{body}</figure>{side}')
    elif layout == "quote":
        q = data.get("quote") or {}
        parts.append('<div class="quote-mark">“</div>')
        parts.append(f'<blockquote>{inline(q.get("text") or data.get("title",""))}</blockquote>')
        if q.get("cite"):
            parts.append(f'<div class="cite">{inline(q["cite"])}</div>')
    elif layout == "stats":
        parts.append(_title_block(data))
        cards = "".join(
            f'<div class="stat"><div class="stat-v">{inline(str(s.get("value","")))}</div>'
            f'<div class="stat-l">{inline(str(s.get("label","")))}</div>'
            + (f'<div class="stat-n">{inline(str(s["note"]))}</div>' if s.get("note") else "")
            + "</div>"
            for s in (data.get("stats") or [])
        )
        parts.append(f'<div class="stats">{cards}</div>')
    elif layout == "timeline":
        parts.append(_title_block(data))
        steps = "".join(
            f'<div class="step"><div class="dot"></div><div class="step-i">{i:02d}</div>'
            f'<div class="step-n">{inline(str(s.get("name","")))}</div>'
            f'<div class="step-d">{inline(str(s.get("desc","")))}</div></div>'
            for i, s in enumerate(data.get("steps") or [], start=1)
        )
        parts.append(f'<div class="timeline">{steps}</div>')
    elif layout == "refs":
        parts.append(_title_block(data))
        refs = data.get("refs") or [b["text"] for b in data.get("bullets") or []]
        parts.append('<ol class="refs">' + "".join(f"<li>{inline(r)}</li>" for r in refs) + "</ol>")
    elif layout == "two-col":
        parts.append(_title_block(data))
        cols = ""
        for key in ("left", "right"):
            col = data.get(key)
            if not col:
                continue
            cols += '<div class="col">'
            if col.get("title"):
                cols += f'<h3>{inline(col["title"])}</h3>'
            cols += _bullets_html(col.get("bullets") or []) + "</div>"
        parts.append(f'<div class="cols">{cols}</div>')
    elif layout == "end":
        parts.append('<div class="rule center"></div>')
        parts.append(f'<h1 class="end-title">{inline(data.get("title","谢谢"))}</h1>')
        if data.get("subtitle"):
            parts.append(f'<p class="end-sub">{inline(data["subtitle"])}</p>')
        if data.get("contact"):
            parts.append(f'<p class="end-contact">{inline(data["contact"])}</p>')
    else:  # bullets
        parts.append(_title_block(data))
        if data.get("subtitle") and not data.get("title"):
            parts.append(f'<p class="lead">{inline(data["subtitle"])}</p>')
        parts.append(_bullets_html(data.get("bullets") or []))

    n = "" if layout in ("cover", "end") else f'<div class="pagenum">{idx:02d} / {total:02d}</div>'
    foot = f'<div class="foot">{html.escape(footer)}</div>' if footer and layout not in ("cover", "end") else ""
    note = data.get("note") or ""
    note_html = f'<aside class="notes">{inline(note)}</aside>' if note else '<aside class="notes"></aside>'
    extra = " has-side" if layout == "image" and data.get("bullets") else ""
    return (f'<section class="slide s-{layout}{extra}" data-layout="{layout}">{foot}'
            + "".join(parts) + n + note_html + "</section>")


def _title_block(data: dict) -> str:
    out = []
    if data.get("kicker"):
        out.append(f'<div class="kicker">{inline(data["kicker"])}</div>')
    if data.get("title"):
        out.append(f'<h1>{inline(data["title"])}</h1>')
        out.append('<div class="rule"></div>')
    if data.get("subtitle"):
        out.append(f'<p class="lead">{inline(data["subtitle"])}</p>')
    return "".join(out)


def build(md_path: str | Path, out_path: str | Path | None = None,
          theme: str | None = None, title: str | None = None,
          theme_obj: Theme | None = None, image_mode: str = "relative") -> Path:
    """Markdown → 单文件 HTML。

    theme_obj 优先于 theme 名：风格档案（profile.apply()）返回的 Theme 可以直接
    传进来，网页版与 pptx 用同一套配色/字体（契约 §5.1）。档案数据来自用户上传，
    写进 <style> 前会对颜色和字体名做白名单校验（评审 F10）。
    image_mode=placeholder 时所有图片都渲染成占位（服务端 deck 目录不携带图片）。
    """
    src = Path(md_path)
    text = src.read_text(encoding="utf-8")
    doc = md.parse(text)
    meta = doc.get("meta") or {}
    th = theme_obj if isinstance(theme_obj, Theme) else get_theme(theme or meta.get("theme"))
    slides = doc.get("slides") or []
    if meta.get("title"):
        slides = [{
            "layout": "cover", "title": meta["title"],
            "subtitle": meta.get("subtitle", ""),
            "kicker": meta.get("kicker") or meta.get("type") or "",
            "meta": {"author": meta.get("author"), "org": meta.get("org"),
                     "date": meta.get("date")},
        }] + slides
    total = len(slides)
    footer = meta.get("footer") or " · ".join(
        [x for x in (meta.get("author", ""), meta.get("org", "")) if x])

    body = "\n".join(_slide_html(s, i, total, footer, image_mode)
                     for i, s in enumerate(slides, start=1))
    page_title = title or meta.get("title") or src.stem
    out = Path(out_path) if out_path else src.with_suffix(".html")

    font_title = _safe_font(th.font_title)
    font_body = _safe_font(th.font_body)
    css_vars = f"""
    --bg:{_safe_color(th.bg, "#FFFFFF")}; --ink:{_safe_color(th.ink, "#14213D")};
    --muted:{_safe_color(th.muted, "#5C6B82")};
    --accent:{_safe_color(th.accent, "#1B3A6B")}; --soft:{_safe_color(th.accent_soft, "#E9EEF7")};
    --rule:{_safe_color(th.rule, "#D5DCE8")};
    --font-title:"{font_title}", "{FALLBACK_FONT}", sans-serif;
    --font-body:"{font_body}", "{FALLBACK_FONT}", sans-serif;
    """.strip()

    html_doc = TEMPLATE.format(
        title=html.escape(page_title), theme_vars=css_vars, body=body,
        total=total, theme_key=html.escape(str(th.key or ""), quote=True),
    )
    out.write_text(html_doc, encoding="utf-8")
    return out


TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN" data-theme="{theme_key}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{title}</title>
<style>
:root {{ {theme_vars} }}
* {{ box-sizing:border-box; margin:0; padding:0; }}
html,body {{ height:100%; background:#0d0d0f; overflow:hidden; }}
body {{ font-family:var(--font-body); color:var(--ink); -webkit-font-smoothing:antialiased; }}
#app {{ position:fixed; inset:0; display:grid; place-items:center; }}
.stage {{ width:1280px; height:720px; position:relative; transform-origin:center center; }}
.slide {{
  position:absolute; inset:0; display:none; background:var(--bg);
  padding:64px 76px 76px; overflow:hidden;
}}
.slide.active {{ display:block; }}
h1 {{ font-family:var(--font-title); font-size:42px; line-height:1.2; font-weight:700; letter-spacing:.2px; }}
.kicker {{ font-size:12px; letter-spacing:2.4px; color:var(--accent); font-weight:700; margin-bottom:10px; text-transform:uppercase; }}
.rule {{ width:64px; height:3px; background:var(--accent); margin:18px 0 30px; }}
.rule.center {{ margin:0 auto 30px; }}
.lead {{ font-size:20px; color:var(--muted); line-height:1.5; margin-bottom:8px; }}
ul.bullets {{ list-style:none; }}
ul.bullets li {{ font-size:20px; line-height:1.6; padding-left:26px; position:relative; margin-bottom:16px; }}
ul.bullets li::before {{ content:"•"; position:absolute; left:4px; color:var(--accent); font-weight:700; }}
ul.bullets li.lvl1 {{ font-size:17px; color:var(--muted); margin-left:26px; margin-bottom:10px; }}
ul.bullets li.lvl1::before {{ content:"–"; color:var(--muted); }}
ul.bullets li.lvl2 {{ font-size:16px; color:var(--muted); margin-left:52px; }}
ul.bullets li[data-plain] {{ padding-left:0; }}
ul.bullets li[data-plain]::before {{ content:""; }}
.foot {{ position:absolute; left:76px; bottom:34px; font-size:11px; color:var(--muted); }}
.pagenum {{ position:absolute; right:76px; bottom:34px; font-size:11px; color:var(--muted); letter-spacing:.6px; }}
.s-cover {{ padding-top:150px; }}
.s-cover .cover-title {{ font-size:56px; line-height:1.18; max-width:1000px; }}
.s-cover .cover-sub {{ font-size:21px; color:var(--muted); margin-top:6px; max-width:900px; }}
.cover-meta {{ position:absolute; left:76px; bottom:96px; display:flex; gap:18px; align-items:baseline; }}
.cover-meta span {{ font-size:13px; color:var(--muted); }}
.cover-meta span:first-child {{ color:var(--ink); font-weight:600; font-size:15px; }}
.s-section .section-num {{ position:absolute; right:40px; top:-30px; font-size:210px; font-weight:800;
  color:var(--soft); font-family:var(--font-title); letter-spacing:-6px; }}
.s-section .kicker, .s-section .section-title, .s-section .section-sub {{ position:relative; }}
.s-section {{ padding-top:200px; }}
.s-section .section-title {{ font-size:46px; }}
.s-section .section-sub {{ font-size:17px; color:var(--muted); margin-top:26px; max-width:760px; }}
blockquote {{ font-size:32px; line-height:1.55; font-family:var(--font-title); max-width:1000px; margin-top:18px; }}
.s-quote {{ padding-top:150px; }}
.quote-mark {{ font-size:150px; line-height:.7; color:var(--soft); font-family:var(--font-title); }}
.cite {{ font-size:14px; color:var(--muted); margin-top:34px; padding-left:16px; border-left:3px solid var(--accent); }}
.s-stats {{ padding-top:120px; }}
.stats {{ display:flex; gap:42px; margin-top:44px; }}
.stat {{ flex:1; border-top:3px solid var(--accent); padding-top:18px; }}
.stat-v {{ font-size:52px; font-weight:700; color:var(--accent); line-height:1.05; }}
.stat-l {{ font-size:14px; color:var(--muted); margin-top:10px; line-height:1.5; }}
.stat-n {{ font-size:12px; color:var(--muted); margin-top:8px; opacity:.85; }}
.s-timeline {{ padding-top:130px; }}
.timeline {{ display:flex; gap:22px; margin-top:52px; position:relative; }}
.timeline::before {{ content:""; position:absolute; left:0; right:0; top:7px; height:1px; background:var(--rule); }}
.step {{ flex:1; position:relative; padding-top:34px; }}
.dot {{ position:absolute; top:0; left:0; width:15px; height:15px; border-radius:50%; background:var(--accent); }}
.step-i {{ font-size:11px; letter-spacing:1.4px; color:var(--accent); font-weight:700; }}
.step-n {{ font-size:17px; font-weight:700; margin-top:8px; }}
.step-d {{ font-size:13px; color:var(--muted); margin-top:10px; line-height:1.55; }}
table {{ width:100%; border-collapse:collapse; margin-top:34px; font-size:15px; }}
th {{ background:var(--accent); color:#fff; text-align:left; padding:12px 14px; font-weight:600; }}
td {{ padding:11px 14px; border-bottom:1px solid var(--rule); }}
tbody tr:nth-child(even) {{ background:var(--soft); }}
.s-image figure {{ position:absolute; left:76px; right:76px; top:230px; bottom:96px; display:flex;
  flex-direction:column; align-items:center; justify-content:center; }}
.s-image img {{ max-width:100%; max-height:88%; object-fit:contain; }}
.s-image .img-missing {{ padding:26px 34px; border:1px dashed var(--rule); border-radius:10px;
  color:var(--muted); font-size:15px; letter-spacing:.5px; }}
.s-image.has-side figure {{ right:auto; width:720px; }}
.s-image .img-side {{ position:absolute; right:76px; top:230px; width:340px; }}
.s-image figcaption {{ font-size:12px; color:var(--muted); margin-top:12px; }}
.cols {{ display:flex; gap:60px; margin-top:8px; }}
.col {{ flex:1; }}
.col h3 {{ font-size:15px; color:var(--accent); margin-bottom:18px; letter-spacing:.6px; }}
.s-refs ol.refs {{ columns:2; column-gap:46px; margin-top:28px; list-style:none; counter-reset:r; }}
.s-refs ol.refs li {{ font-size:13px; line-height:1.55; margin-bottom:14px; break-inside:avoid;
  color:var(--ink); padding-left:30px; position:relative; counter-increment:r; }}
.s-refs ol.refs li::before {{ content:"[" counter(r) "]"; position:absolute; left:0; color:var(--muted); }}
.s-end {{ flex-direction:column; align-items:center; justify-content:center; }}
.s-end.active {{ display:flex; }}
.end-title {{ font-size:60px; }}
.end-sub {{ font-size:19px; color:var(--muted); margin-top:18px; }}
.end-contact {{ font-size:13px; color:var(--muted); margin-top:40px; letter-spacing:.4px; }}
/* 进度条 / 控件 */
#bar {{ position:fixed; left:0; top:0; height:2px; background:var(--accent); width:0; transition:width .25s; z-index:40; }}
#hint {{ position:fixed; right:18px; top:14px; font:11px/1.6 ui-monospace,monospace; color:#8d8d96; z-index:40; }}
#notes {{ position:fixed; left:0; right:0; bottom:0; max-height:34vh; overflow:auto; background:#111114;
  color:#d8d8e0; font:14px/1.7 var(--font-body); padding:18px 24px; display:none; z-index:30;
  border-top:1px solid #2a2a30; }}
#notes.on {{ display:block; }}
#notes b {{ color:#ffd166; font-size:11px; letter-spacing:1.4px; display:block; margin-bottom:6px; }}
#help {{ position:fixed; inset:0; background:rgba(8,8,10,.94); color:#eee; display:none; z-index:50;
  padding:8vh 10vw; }}
#help.on {{ display:block; }}
#help h2 {{ font-size:20px; margin-bottom:18px; }}
#help kbd {{ background:#26262c; border-radius:4px; padding:2px 7px; font:12px ui-monospace,monospace; margin-right:8px; }}
#help p {{ font:14px/2.2 var(--font-body); color:#cfcfd6; }}
#overview {{ position:fixed; inset:0; background:#0b0b0d; display:none; grid-template-columns:repeat(4,1fr);
  gap:18px; padding:34px; overflow:auto; z-index:35; }}
#overview.on {{ display:grid; }}
#overview .thumb {{ position:relative; aspect-ratio:16/9; background:#000; overflow:hidden; cursor:pointer;
  outline:1px solid #26262c; }}
#overview .thumb .inner {{ width:1280px; height:720px; transform-origin:top left; position:absolute; }}
#overview .thumb .inner .slide {{ display:block; position:relative; width:1280px; height:720px; }}
#overview .thumb span {{ position:absolute; left:6px; bottom:4px; color:#9a9aa3; font:10px ui-monospace,monospace; }}
@media print {{
  @page {{ size:1280px 720px; margin:0; }}
  html, body {{ background:#fff; overflow:visible; height:auto; }}
  #bar, #hint, #notes, #help, #overview {{ display:none !important; }}
  #app {{ position:static; display:block; }}
  .stage {{ width:1280px; height:auto; transform:none !important; }}
  .slide {{ position:relative !important; display:block !important;
            width:1280px; height:720px; page-break-after:always; }}
}}
</style>
</head>
<body>
<div id="bar"></div>
<div id="hint">← → 翻页 · O 总览 · S 备注 · ? 帮助</div>
<div id="app"><div class="stage" id="stage">
{body}
</div></div>
<div id="overview"></div>
<div id="notes"><b>讲者备注</b><div id="note-body"></div></div>
<div id="help">
  <h2>快捷键</h2>
  <p><kbd>→</kbd><kbd>空格</kbd> 下一页　<kbd>←</kbd> 上一页　<kbd>Home</kbd>/<kbd>End</kbd> 首/末页</p>
  <p><kbd>O</kbd> 缩略图总览　<kbd>S</kbd> 讲者备注　<kbd>F</kbd> 全屏　<kbd>P</kbd> 打印/导出 PDF</p>
  <p><kbd>?</kbd> 关闭帮助　｜　地址栏 #/3 可直接定位到第 3 页</p>
</div>
<script>
const stage = document.getElementById('stage');
const slides = [...stage.querySelectorAll('.slide')];
let cur = 0;
function scale() {{
  const s = Math.min(innerWidth / 1280, innerHeight / 720);
  stage.style.transform = 'scale(' + s + ')';
}}
addEventListener('resize', scale); scale();
function show(i) {{
  cur = Math.max(0, Math.min(slides.length - 1, i));
  slides.forEach((s, k) => s.classList.toggle('active', k === cur));
  document.getElementById('bar').style.width = ((cur + 1) / slides.length * 100) + '%';
  const note = slides[cur].querySelector('.notes');
  document.getElementById('note-body').textContent = note ? note.textContent : '';
  history.replaceState(null, '', '#/' + (cur + 1));
}}
function overview(on) {{
  const el = document.getElementById('overview');
  if (on && !el.dataset.built) {{
    slides.forEach((s, i) => {{
      const t = document.createElement('div');
      t.className = 'thumb';
      const inner = document.createElement('div');
      inner.className = 'inner';
      const clone = s.cloneNode(true);
      clone.classList.add('active');
      inner.appendChild(clone);
      t.appendChild(inner);
      const label = document.createElement('span');
      label.textContent = String(i + 1).padStart(2, '0');
      t.appendChild(label);
      t.onclick = () => {{ overview(false); show(i); }};
      el.appendChild(t);
    }});
    el.dataset.built = '1';
  }}
  el.classList.toggle('on', on);
  if (on) {{
    const w = el.querySelector('.thumb').clientWidth;
    el.querySelectorAll('.inner').forEach(x => x.style.transform = 'scale(' + (w / 1280) + ')');
  }}
}}
addEventListener('keydown', e => {{
  const k = e.key;
  if (k === 'ArrowRight' || k === 'PageDown' || k === ' ') {{ e.preventDefault(); show(cur + 1); }}
  else if (k === 'ArrowLeft' || k === 'PageUp') {{ e.preventDefault(); show(cur - 1); }}
  else if (k === 'Home') show(0);
  else if (k === 'End') show(slides.length - 1);
  else if (k === 'o' || k === 'O') overview(!document.getElementById('overview').classList.contains('on'));
  else if (k === 's' || k === 'S') document.getElementById('notes').classList.toggle('on');
  else if (k === '?') document.getElementById('help').classList.toggle('on');
  else if (k === 'Escape') {{ document.getElementById('help').classList.remove('on');
    document.getElementById('overview').classList.remove('on'); }}
  else if (k === 'f' || k === 'F') {{ document.fullscreenElement ? document.exitFullscreen()
    : document.documentElement.requestFullscreen(); }}
  else if (k === 'p' || k === 'P') {{ window.print(); }}
}});
addEventListener('hashchange', () => show((parseInt(location.hash.replace('#/', ''), 10) || 1) - 1));
show((parseInt(location.hash.replace('#/', ''), 10) || 1) - 1);
</script>
</body>
</html>
"""
