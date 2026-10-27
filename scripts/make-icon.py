"""生成 assets/icon.ico：深蓝底 + 白色 P/幻灯片几何，多尺寸（256/128/64/48/32/16）。

运行：
    & "$env:USERPROFILE\\.dsh\\dsh-runtimes\\dsh-primary-runtime\\dependencies\\python\\python.exe" scripts\\make-icon.py
"""
from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "assets" / "icon.ico"
SIZES = (256, 128, 64, 48, 32, 16)

DEEP_BLUE = (12, 32, 68)
DEEP_BLUE_LIGHT = (23, 58, 107)
WHITE = (255, 255, 255)
ACCENT = (110, 168, 255)


def draw_icon(size: int) -> Image.Image:
    scale = size / 256.0
    image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)

    def rect(x0, y0, x1, y1, fill, radius=0):
        box = [round(v * scale) for v in (x0, y0, x1, y1)]
        if radius:
            draw.rounded_rectangle(box, radius=round(radius * scale), fill=fill)
        else:
            draw.rectangle(box, fill=fill)

    # 深蓝底 + 右侧浅一档的幻灯片面
    rect(0, 0, 256, 256, DEEP_BLUE, radius=52)
    rect(132, 44, 232, 176, DEEP_BLUE_LIGHT, radius=16)
    rect(146, 58, 218, 162, DEEP_BLUE, radius=10)

    # 幻灯片内容：标题条 + 两行要点
    rect(154, 72, 206, 82, ACCENT, radius=4)
    rect(154, 96, 200, 104, WHITE, radius=3)
    rect(154, 116, 188, 124, WHITE, radius=3)
    rect(154, 136, 196, 144, (255, 255, 255, 190), radius=3)

    # 前景白色 P：字母切口 + 幻灯片边框，简洁可缩到 16px
    rect(48, 58, 104, 70, WHITE, radius=5)       # P 顶部
    rect(48, 58, 62, 198, WHITE, radius=6)       # P 竖笔
    rect(48, 116, 104, 128, WHITE, radius=5)     # P 中横
    draw.ellipse([round(88 * scale), round(58 * scale),
                  round(150 * scale), round(122 * scale)], fill=WHITE)
    draw.ellipse([round(102 * scale), round(72 * scale),
                  round(136 * scale), round(108 * scale)], fill=DEEP_BLUE)
    rect(60, 156, 96, 162, ACCENT, radius=3)     # 一条强调下划线
    return image


def main() -> int:
    TARGET.parent.mkdir(parents=True, exist_ok=True)
    base = draw_icon(256)
    base.save(TARGET, format="ICO", sizes=[(size, size) for size in SIZES])
    with Image.open(TARGET) as check:
        sizes = sorted(check.ico.sizes())
    print(f"wrote {TARGET} ({TARGET.stat().st_size} bytes, sizes={sizes})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
