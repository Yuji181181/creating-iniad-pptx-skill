#!/usr/bin/env python3
"""白背景の公式ロゴ画像（JPG など）から、余白を切り詰めて背景を透過した PNG を作る（ロゴ差し替え時に使う）。

使い方: .venv/bin/python scripts/prepare_logo.py <新しいロゴ.jpg> assets/iniad-logo.png
"""
import sys
from PIL import Image


def main(src, dst):
    im = Image.open(src).convert("RGB")
    # 白に近い画素を背景とみなして bbox を取る
    mask = im.point(lambda v: 255 if v < 235 else 0).convert("L")
    bbox = mask.getbbox()
    pad = 4
    bbox = (max(bbox[0] - pad, 0), max(bbox[1] - pad, 0), min(bbox[2] + pad, im.width), min(bbox[3] + pad, im.height))
    im = im.crop(bbox)
    rgba = im.convert("RGBA")
    px = rgba.load()
    lo, hi = 14, 150   # 白からの距離がこれ未満 → 完全透明 / 以上 → 不透明（JPEG ノイズ対策）
    for y in range(rgba.height):
        for x in range(rgba.width):
            r, g, b, _ = px[x, y]
            d = 255 - min(r, g, b)
            a = max(0.0, min(1.0, (d - lo) / (hi - lo)))
            if a == 0:
                px[x, y] = (255, 255, 255, 0)
                continue
            # 白背景との合成を逆算して本来の色を復元（白いにじみを消す）
            un = [max(0, min(255, round(255 - (255 - c) / a))) for c in (r, g, b)]
            px[x, y] = (*un, round(a * 255))
    rgba.save(dst)
    print("saved", dst, rgba.size)


if __name__ == "__main__":
    main(*sys.argv[1:3])
