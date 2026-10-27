"""把导出的 PNG 拼成一张总览图，方便一眼检查整份 deck 的版式。

用法：
    python scripts/contact-sheet.py out/academic-defense-png
    python scripts/contact-sheet.py out/xxx-png -o out/xxx-overview.png -c 4 -w 420
"""
from __future__ import annotations

import argparse
import re
from pathlib import Path

from PIL import Image, ImageDraw


def build(folder: Path, out: Path, cols: int, width: int) -> Path:
    files = sorted(folder.glob("*.png"), key=lambda p: int(re.sub(r"\D", "", p.stem) or 0))
    if not files:
        raise SystemExit(f"{folder} 里没有 PNG")
    thumbs = []
    for f in files:
        im = Image.open(f).convert("RGB")
        h = int(width * im.height / im.width)
        thumbs.append(im.resize((width, h), Image.LANCZOS))
    rows = (len(thumbs) + cols - 1) // cols
    gap, pad = 10, 16
    th = thumbs[0].height
    W = cols * width + (cols - 1) * gap + pad * 2
    H = rows * th + (rows - 1) * gap + pad * 2
    sheet = Image.new("RGB", (W, H), (245, 245, 247))
    d = ImageDraw.Draw(sheet)
    for i, im in enumerate(thumbs):
        r, c = divmod(i, cols)
        x = pad + c * (width + gap)
        y = pad + r * (th + gap)
        sheet.paste(im, (x, y))
        d.rectangle([x - 1, y - 1, x + width, y + th], outline=(210, 210, 215))
        d.text((x + 4, y + 4), f"{i + 1:02d}", fill=(120, 120, 130))
    sheet.save(out)
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("folder")
    ap.add_argument("-o", "--out")
    ap.add_argument("-c", "--cols", type=int, default=4)
    ap.add_argument("-w", "--width", type=int, default=420)
    a = ap.parse_args()
    folder = Path(a.folder)
    out = Path(a.out) if a.out else folder.parent / f"{folder.name}-overview.png"
    print(build(folder, out, a.cols, a.width))
