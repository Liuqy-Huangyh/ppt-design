"""下载并安装开源字体（Noto Sans SC / Noto Serif SC / Inter），按当前用户安装，不需要管理员权限。

用法：
    python scripts/install-fonts.py            # 下载 + 安装
    python scripts/install-fonts.py --check    # 只检查已安装情况

安装位置：%LOCALAPPDATA%\\Microsoft\\Windows\\Fonts\\，并写 HKCU 字体注册表。
"""
from __future__ import annotations

import argparse
import os
import shutil
import sys
import urllib.request
import winreg
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CACHE = ROOT / "assets" / "fonts"
USER_FONTS = Path(os.environ["LOCALAPPDATA"]) / "Microsoft" / "Windows" / "Fonts"
REG_KEY = r"Software\Microsoft\Windows NT\CurrentVersion\Fonts"

BASE = "https://cdn.jsdelivr.net/gh/notofonts/noto-cjk@main"

FILES = [
    # (本地文件名, 下载地址)
    ("NotoSansSC-Regular.otf", f"{BASE}/Sans/SubsetOTF/SC/NotoSansSC-Regular.otf"),
    ("NotoSansSC-Medium.otf", f"{BASE}/Sans/SubsetOTF/SC/NotoSansSC-Medium.otf"),
    ("NotoSansSC-Bold.otf", f"{BASE}/Sans/SubsetOTF/SC/NotoSansSC-Bold.otf"),
    ("NotoSerifSC-Regular.otf", f"{BASE}/Serif/SubsetOTF/SC/NotoSerifSC-Regular.otf"),
    ("NotoSerifSC-Bold.otf", f"{BASE}/Serif/SubsetOTF/SC/NotoSerifSC-Bold.otf"),
    ("Inter-Variable.ttf",
     "https://cdn.jsdelivr.net/gh/google/fonts@main/ofl/inter/Inter%5Bopsz,wght%5D.ttf"),
]


def download(name: str, url: str) -> Path:
    CACHE.mkdir(parents=True, exist_ok=True)
    dest = CACHE / name
    if dest.exists() and dest.stat().st_size > 100_000:
        print(f"  已缓存 {name}（{dest.stat().st_size / 1e6:.1f} MB）")
        return dest
    mirrors = ["cdn.jsdelivr.net", "fastly.jsdelivr.net", "gcore.jsdelivr.net"]
    candidates = [url] + [url.replace("cdn.jsdelivr.net", m) for m in mirrors[1:]]
    last = None
    for _attempt in range(3):
        for cand in candidates:
            try:
                print(f"  下载 {name} …", flush=True)
                with urllib.request.urlopen(cand, timeout=120) as resp, open(dest, "wb") as fh:
                    shutil.copyfileobj(resp, fh)
                if dest.stat().st_size < 100_000:
                    raise OSError("文件过小")
                print(f"  完成 {name}（{dest.stat().st_size / 1e6:.1f} MB）")
                return dest
            except Exception as exc:
                last = exc
                print(f"    重试（{type(exc).__name__}: {str(exc)[:40]}）", flush=True)
                continue
    raise SystemExit(f"下载失败：{name} -> {last}")


def instantiate_inter(var_path: Path) -> list[Path]:
    """把 Inter 可变字体实例化成静态字重，PowerPoint 对可变字体支持不稳。"""
    try:
        from fontTools.varLib.instancer import instantiateVariableFont
        from fontTools.ttLib import TTFont
    except ImportError:
        print("  ! 缺少 fonttools，跳过 Inter 实例化")
        return []
    out = []
    for weight, suffix in ((400, "Regular"), (500, "Medium"), (700, "Bold")):
        dest = CACHE / f"Inter-{suffix}.ttf"
        if dest.exists():
            out.append(dest)
            continue
        font = TTFont(var_path)
        instantiateVariableFont(font, {"wght": weight, "opsz": 20}, inplace=True)
        # 改名字，避免三个实例在系统里显示成一个字体
        for record in font["name"].names:
            if record.nameID in (1, 2, 4, 6, 16, 17):
                try:
                    text = record.toUnicode()
                except Exception:
                    continue
                if record.nameID == 1:
                    record.string = "Inter"
                elif record.nameID == 2:
                    record.string = suffix
                elif record.nameID == 4:
                    record.string = f"Inter {suffix}"
                elif record.nameID == 6:
                    record.string = f"Inter-{suffix}"
                elif record.nameID in (16, 17):
                    record.string = "Inter" if record.nameID == 16 else suffix
                else:
                    _ = text
        font.save(dest)
        out.append(dest)
        print(f"  实例化 {dest.name}")
    return out


def family_of(path: Path) -> tuple[str, str]:
    from fontTools.ttLib import TTFont
    font = TTFont(path, fontNumber=0)
    family, sub = "Unknown", "Regular"
    for record in font["name"].names:
        if record.nameID == 1 and record.platformID == 3:
            family = record.toUnicode()
        if record.nameID == 2 and record.platformID == 3:
            sub = record.toUnicode()
    return family, sub


def install(paths: list[Path]) -> None:
    USER_FONTS.mkdir(parents=True, exist_ok=True)
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_KEY, 0, winreg.KEY_SET_VALUE) as key:
        for src in paths:
            family, sub = family_of(src)
            target = USER_FONTS / src.name
            if not target.exists() or target.stat().st_size != src.stat().st_size:
                shutil.copy2(src, target)
            label = f"{family} {sub if sub.lower() not in ('regular', 'normal') else ''}".strip()
            value_name = f"{label} (OpenType)"
            winreg.SetValueEx(key, value_name, 0, winreg.REG_SZ, str(target))
            print(f"  [OK] {value_name} -> {target.name}")


def check() -> None:
    print(f"用户字体目录：{USER_FONTS}")
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_KEY) as key:
        count = 0
        while True:
            try:
                name, value, _ = winreg.EnumValue(key, count)
            except OSError:
                break
            if any(k in name for k in ("Noto Sans SC", "Noto Serif SC", "Inter")):
                print(f"  [OK] {name}  {value}")
            count += 1
    print(f"缓存：{CACHE}")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true")
    args = ap.parse_args()
    if args.check:
        check()
        return 0

    print("1/3 下载字体")
    files = [download(name, url) for name, url in FILES]

    print("2/3 实例化 Inter 静态字重")
    files += instantiate_inter(CACHE / "Inter-Variable.ttf")

    print("3/3 安装到当前用户")
    install(files)
    print("\n装好了。若 PowerPoint/WPS 已打开，请重启后再选字体。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
