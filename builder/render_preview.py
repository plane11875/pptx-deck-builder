#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
render_preview.py — 把 slides/*.svg 渲染成 output/preview/*.png。

注意:预览图是从 **SVG 源** 渲染的(和编译进 PPTX 的是同一份几何),
不是把 PPTX 截图 —— 渲染 PPTX 需要 LibreOffice,这里不引入这个依赖。

依赖:cairosvg(pip install cairosvg),可选,缺失时跳过。
用法:python3 builder/render_preview.py [--scale 2]
"""

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SLIDES = ROOT / "slides"
OUT_DIR = ROOT / "output" / "preview"


def main() -> int:
    ap = argparse.ArgumentParser(description="渲染幻灯片预览图")
    ap.add_argument("--scale", type=float, default=2.0, help="渲染倍率(默认 2)")
    args = ap.parse_args()

    try:
        import cairosvg
    except ImportError:
        print("跳过预览渲染:未安装 cairosvg(pip install cairosvg)", file=sys.stderr)
        return 0

    files = sorted(SLIDES.glob("slide_*.svg"))
    if not files:
        print("没有找到 slides/*.svg,先运行 python3 builder/make_slides.py", file=sys.stderr)
        return 1

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    for svg in files:
        png = OUT_DIR / f"{svg.stem}.png"
        cairosvg.svg2png(url=str(svg), write_to=str(png), scale=args.scale)
        print(f"rendered {png.relative_to(ROOT)} ({png.stat().st_size // 1024} KB)")
    print(f"共 {len(files)} 张预览图 → {OUT_DIR.relative_to(ROOT)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
