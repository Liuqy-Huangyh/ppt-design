"""命令行入口：python -m pptxkit <命令>"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from . import builder, html_deck, inspect as inspector
from .themes import THEMES


def cmd_build(args) -> int:
    out = builder.build_file(args.src, args.out, theme=args.theme)
    n = len(__import__("pptx").Presentation(str(out)).slides)
    print(f"[OK] 已生成 {out}（{n} 页，主题 {args.theme or '自动'}）")
    return 0


def cmd_html(args) -> int:
    out = html_deck.build(args.src, args.out, theme=args.theme)
    print(f"[OK] 已生成 {out}（双击打开，← → 翻页，F 全屏）")
    return 0


def cmd_qa(args) -> int:
    report = inspector.qa(args.src)
    print(inspector.format_qa(report))
    return 0


def cmd_to_md(args) -> int:
    out = inspector.to_markdown(args.src, args.out)
    print(f"[OK] 已反解为 Markdown：{out}")
    return 0


def cmd_themes(_args) -> int:
    print("可用主题：")
    for key, th in THEMES.items():
        print(f"  {key:<10} {th.label}　标题字体={th.font_title}　正文字体={th.font_body}")
    return 0


def cmd_profile(args) -> int:
    """分析一个案例 PPT，打印（可选落盘）风格档案。"""
    from . import profile as profiler

    src = Path(args.src)
    if not src.exists():
        raise FileNotFoundError(f"找不到文件：{src}")
    pf = profiler.analyze(src, name=args.name or "", source=src.name)
    theme = profiler.apply(pf, args.base_theme)
    metrics = pf["metrics"]
    palette = pf["palette"]
    fonts = pf["fonts"]
    print(f"[OK] 已分析 {src}（{metrics['slides']} 页，"
          f"{pf['slide_size'][0]}×{pf['slide_size'][1]} in）")
    print("  配色  " + "  ".join(f"{key}={palette[key]}" for key in
                                 ("bg", "ink", "muted", "accent", "accent_soft", "rule")))
    print(f"  字体  标题={fonts['title']}　正文={fonts['body']}")
    print(f"  指标  每页字数={metrics['chars_per_slide']}　每页要点={metrics['bullets_per_slide']}　"
          f"标题={metrics['title_pt']}pt　正文={metrics['body_pt']}pt　"
          f"图片占比={metrics['image_ratio']:.0%}")
    print(f"  风格  密度={pf['style']['density']}　标题比例={pf['style']['title_scale']}")
    print(f"  主题预览  bg=#{theme.bg} ink=#{theme.ink} accent=#{theme.accent} "
          f"标题字号={theme.title_size}pt 标题字体={theme.font_title}")

    if args.out:
        out = Path(args.out)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(profiler._public(pf), ensure_ascii=False, indent=2),
                       encoding="utf-8")
        print(f"[OK] 档案已写入 {out}")
    if args.store:
        saved = profiler.save(pf, args.store)
        print(f"[OK] 档案已入库：{saved['id']}（{Path(args.store) / 'profiles' / (saved['id'] + '.json')}）")
    if args.json:
        print(json.dumps(profiler._public(pf), ensure_ascii=False, indent=2))
    return 0


def main(argv: list[str] | None = None) -> int:
    # Windows 控制台默认 GBK：遇到生僻字只替换不崩
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(errors="replace")
        except Exception:
            pass
    parser = argparse.ArgumentParser(
        prog="pptxkit", description="Markdown/JSON → 可编辑 PPTX 或单文件 HTML 演示")
    sub = parser.add_subparsers(dest="cmd", required=True)

    b = sub.add_parser("build", help="Markdown/JSON → PPTX")
    b.add_argument("src")
    b.add_argument("-o", "--out")
    b.add_argument("--theme", choices=list(THEMES), default=None)
    b.set_defaults(func=cmd_build)

    h = sub.add_parser("html", help="Markdown → 单文件 HTML 演示")
    h.add_argument("src")
    h.add_argument("-o", "--out")
    h.add_argument("--theme", choices=list(THEMES), default=None)
    h.set_defaults(func=cmd_html)

    q = sub.add_parser("qa", help="检查 PPTX：字数、溢出、字体、备注")
    q.add_argument("src")
    q.set_defaults(func=cmd_qa)

    t = sub.add_parser("to-md", help="PPTX → Markdown（可再编辑）")
    t.add_argument("src")
    t.add_argument("-o", "--out")
    t.set_defaults(func=cmd_to_md)

    th = sub.add_parser("themes", help="列出内置主题")
    th.set_defaults(func=cmd_themes)

    pf = sub.add_parser("profile", help="分析案例 PPTX → 打印/保存风格档案")
    pf.add_argument("src")
    pf.add_argument("-o", "--out", help="把档案写成 JSON 文件")
    pf.add_argument("--store", help="档案库目录（落盘 profiles/<id>.json 与 cases/<id>/source.pptx）")
    pf.add_argument("--name", default="", help="档案显示名，默认用文件名")
    pf.add_argument("--base-theme", choices=list(THEMES), default=None,
                    help="预览 apply() 之后的主题（默认 academic）")
    pf.add_argument("--json", action="store_true", help="把完整档案打印成 JSON")
    pf.set_defaults(func=cmd_profile)

    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except SystemExit:
        raise
    except FileNotFoundError as exc:
        print(f"[!] {exc}")
        return 2
    except Exception as exc:  # noqa: BLE001
        print(f"[!] 出错了：{exc}")
        if "--trace" in sys.argv:
            raise
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
