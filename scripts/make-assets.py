"""生成模板用的示意图（纯 Pillow 绘制，可重复执行）。

产出：
    assets/figs/framework.png  研究框架 / 流水线示意
    assets/figs/results.png    实验结果柱状图
"""
from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "assets" / "figs"
OUT.mkdir(parents=True, exist_ok=True)

INK = (20, 33, 61)
MUTED = (108, 122, 143)
ACCENT = (27, 58, 107)
ACCENT2 = (201, 162, 39)
SOFT = (233, 238, 247)
RULE = (213, 220, 232)


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    candidates = [
        r"C:\Windows\Fonts\msyhbd.ttc" if bold else r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\simhei.ttf",
    ]
    for path in candidates:
        if Path(path).exists():
            try:
                return ImageFont.truetype(path, size)
            except OSError:
                continue
    return ImageFont.load_default()


def arrow(draw: ImageDraw.ImageDraw, x1: int, y1: int, x2: int, y2: int, color, width=3):
    draw.line([(x1, y1), (x2, y2)], fill=color, width=width)
    head = 12
    draw.polygon([(x2, y2), (x2 - head, y2 - head // 2), (x2 - head, y2 + head // 2)], fill=color)


def make_framework() -> Path:
    W, H = 1600, 760
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)

    d.text((60, 46), "整体框架", font=font(34, True), fill=INK)
    d.text((60, 96), "输入 → 编码 → 融合 → 输出，四个阶段端到端可训练", font=font(20), fill=MUTED)
    d.line([(60, 140), (1540, 140)], fill=RULE, width=2)

    stages = [
        ("01", "输入", "多源图像\n统一预处理"),
        ("02", "编码", "轻量主干\n多尺度特征"),
        ("03", "融合", "跨层注意力\n特征对齐"),
        ("04", "输出", "检测头\n置信度校准"),
    ]
    bw, bh, gap = 300, 250, 80
    x0, y0 = 120, 280
    for i, (num, name, desc) in enumerate(stages):
        x = x0 + i * (bw + gap)
        d.rounded_rectangle([x, y0, x + bw, y0 + bh], radius=18, fill=SOFT, outline=RULE, width=2)
        d.text((x + 24, y0 + 20), num, font=font(22, True), fill=ACCENT2)
        d.text((x + 24, y0 + 58), name, font=font(30, True), fill=ACCENT)
        d.multiline_text((x + 24, y0 + 118), desc, font=font(20), fill=MUTED, spacing=10)
        if i < len(stages) - 1:
            arrow(d, x + bw + 12, y0 + bh // 2, x + bw + gap - 14, y0 + bh // 2, ACCENT, 3)

    d.text((120, 590), "训练阶段：知识蒸馏 + 一致性正则　　推理阶段：单卡 38 FPS @ 640×640",
           font=font(19), fill=MUTED)
    d.line([(120, 660), (1480, 660)], fill=RULE, width=1)
    d.text((120, 680), "注：示意图由 scripts/make-assets.py 生成，可直接替换成自己的架构图。",
           font=font(16), fill=MUTED)

    path = OUT / "framework.png"
    img.save(path, "PNG")
    return path


def make_results() -> Path:
    W, H = 1600, 860
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    d.text((70, 50), "精度—速度权衡", font=font(34, True), fill=INK)
    d.text((70, 100), "在 COCO val2017 上对比同类方法，气泡越大表示参数量越多", font=font(19), fill=MUTED)

    left, top, right, bottom = 160, 200, 1480, 720
    d.line([(left, bottom), (right, bottom)], fill=RULE, width=2)
    d.line([(left, bottom), (left, top)], fill=RULE, width=2)

    # 网格与 Y 轴刻度
    for i, val in enumerate([30, 35, 40, 45]):
        y = bottom - (val - 28) * 34
        d.line([(left, y), (right, y)], fill=(240, 242, 246), width=1)
        d.text((left - 56, y - 12), f"{val}", font=font(17), fill=MUTED)
    d.text((70, top - 40), "mAP (%)", font=font(17), fill=MUTED)

    # 柱：五个方法
    data = [
        ("YOLOv8‑s", 44.9, 21.5, True),
        ("YOLOv10‑s", 46.3, 24.5, False),
        ("RT‑DETR", 43.0, 32.0, False),
        ("本文（轻量）", 46.8, 12.4, True),
        ("本文（完整）", 48.6, 27.3, True),
    ]
    slot = (right - left) / len(data)
    for i, (name, mAP, params, ours) in enumerate(data):
        cx = left + slot * (i + 0.5)
        h = (mAP - 28) * 34
        bw = 96
        color = ACCENT if not ours else ACCENT2
        d.rounded_rectangle([cx - bw / 2, bottom - h, cx + bw / 2, bottom - 4], radius=6, fill=color)
        d.text((cx - 34, bottom - h - 34), f"{mAP}%", font=font(18, True), fill=INK)
        label = name
        d.text((cx - len(label) * 7, bottom + 16), label, font=font(17), fill=INK if ours else MUTED)

    # 图例
    d.rectangle([left, 780, left + 18, 798], fill=ACCENT)
    d.text((left + 28, 778), "已有方法", font=font(17), fill=MUTED)
    d.rectangle([left + 160, 780, left + 178, 798], fill=ACCENT2)
    d.text((left + 188, 778), "本文方法", font=font(17), fill=MUTED)
    d.text((right - 420, 778), "参数量：12.4M / 27.3M", font=font(17), fill=MUTED)

    path = OUT / "results.png"
    img.save(path, "PNG")
    return path


if __name__ == "__main__":
    for p in (make_framework(), make_results()):
        print(f"[OK] {p}")
