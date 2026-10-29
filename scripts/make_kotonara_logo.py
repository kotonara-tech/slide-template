# -*- coding: utf-8 -*-
"""黒背景の「ことテク」ロゴ画像から、白・明るい部分を輪郭トレースしてベクター化し、
指定色の透過 SVG と高解像度 PNG を作る。

使い方:
  python scripts/make_kotonara_logo.py <入力画像> <出力ディレクトリ> [--color 4D9622] [--name kotonara-logo] [--png-width 4000]

  assets/ のロゴは次の2回の実行で作っている（リポジトリルートで）:
    python scripts/make_kotonara_logo.py assets/kotonara-logo-source.jpg assets
    python scripts/make_kotonara_logo.py assets/kotonara-logo-source.jpg assets --color FFFFFF --name kotonara-logo-white

依存: numpy, Pillow, potracer（pip install potracer）, cairosvg
"""
import argparse
import os

import cairosvg
import numpy as np
import potrace
from PIL import Image, ImageFilter

UPSCALE = 4          # トレース前の拡大率（輪郭を滑らかにするため）
THRESHOLD = 0.45     # 明るさ（RGB の最大値）のしきい値。0〜1


def build_mask(path):
    im = Image.open(path).convert("RGB")
    w, h = im.size
    big = im.resize((w * UPSCALE, h * UPSCALE), Image.LANCZOS)
    lum = big.split()
    arr = np.maximum.reduce([np.asarray(c, dtype=np.float32) for c in lum]) / 255.0
    g = Image.fromarray((arr * 255).astype(np.uint8)).filter(ImageFilter.GaussianBlur(UPSCALE * 0.6))
    arr = np.asarray(g, dtype=np.float32) / 255.0
    return arr > THRESHOLD, (w * UPSCALE, h * UPSCALE)


def trace_svg(mask, size, color):
    # potracer は False（0）の画素を図形として扱うため反転して渡す
    bmp = potrace.Bitmap(~mask)
    plist = bmp.trace(turdsize=40, turnpolicy=potrace.POTRACE_TURNPOLICY_MINORITY,
                      alphamax=1.0, opticurve=True, opttolerance=0.2)
    parts = []
    for curve in plist:
        fs = curve.start_point
        d = [f"M{fs.x:.2f},{fs.y:.2f}"]
        for seg in curve.segments:
            if seg.is_corner:
                d.append(f"L{seg.c.x:.2f},{seg.c.y:.2f}L{seg.end_point.x:.2f},{seg.end_point.y:.2f}")
            else:
                d.append(f"C{seg.c1.x:.2f},{seg.c1.y:.2f} {seg.c2.x:.2f},{seg.c2.y:.2f} "
                         f"{seg.end_point.x:.2f},{seg.end_point.y:.2f}")
        d.append("Z")
        parts.append("".join(d))
    w, h = size
    return (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {w} {h}" '
            f'width="{w}" height="{h}">'
            f'<path fill="#{color}" fill-rule="evenodd" d="{" ".join(parts)}"/></svg>')


def trim_svg(svg_text, png_probe):
    """透明な余白を除くため、描画結果の外接矩形で viewBox を詰める。"""
    im = Image.open(png_probe)
    l, t, r, b = im.getchannel("A").getbbox()
    pad = 4
    return l - pad, t - pad, (r - l) + 2 * pad, (b - t) + 2 * pad


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("out_dir")
    ap.add_argument("--color", default="4D9622")
    ap.add_argument("--name", default="kotonara-logo", help="出力ファイル名（拡張子なし）")
    ap.add_argument("--png-width", type=int, default=4000)
    a = ap.parse_args()
    os.makedirs(a.out_dir, exist_ok=True)

    mask, size = build_mask(a.src)
    svg = trace_svg(mask, size, a.color)

    probe = os.path.join(a.out_dir, f"_{a.name}_probe.png")
    cairosvg.svg2png(bytestring=svg.encode(), write_to=probe)
    x, y, w, h = trim_svg(svg, probe)
    os.remove(probe)
    svg = svg.replace(f'viewBox="0 0 {size[0]} {size[1]}" width="{size[0]}" height="{size[1]}"',
                      f'viewBox="{x} {y} {w} {h}" width="{w}" height="{h}"')

    svg_path = os.path.join(a.out_dir, a.name + ".svg")
    png_path = os.path.join(a.out_dir, a.name + ".png")
    with open(svg_path, "w", encoding="utf-8") as f:
        f.write(svg)
    cairosvg.svg2png(bytestring=svg.encode(), write_to=png_path, output_width=a.png_width)
    im = Image.open(png_path)
    print(f"SVG: {svg_path}  ({len(svg)//1024} KB)")
    print(f"PNG: {png_path}  {im.size}")


if __name__ == "__main__":
    main()
