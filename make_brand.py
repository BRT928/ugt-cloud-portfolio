#!/usr/bin/env python3
"""Regenerate the logo + favicon assets in brand/ from brand/src/ugt-cloud-logo.png.

build.py copies brand/* into docs/assets on every build, so this only needs
re-running if the source logo changes.
"""
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFilter

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "brand" / "src" / "ugt-cloud-logo.png"
OUT = ROOT / "brand"
NAVY = np.array([21, 59, 182])
CLOUD = np.array([0, 128, 235])


def mark_square(logo: Image.Image) -> Image.Image:
    """Cut the navy triangle-in-circle mark out of the logo (dropping the cloud behind it)."""
    a = np.array(logo)
    rgb = a[..., :3].astype(float)
    navy = ((rgb - NAVY) ** 2).sum(-1) < ((rgb - CLOUD) ** 2).sum(-1)
    m = a.copy()
    m[~navy, 3] = 0
    m[:, int(a.shape[1] * 0.60):, 3] = 0          # drop the "UGT Cloud" wordmark
    img = Image.fromarray(m)
    img = img.crop(img.getbbox())
    s = max(img.size)
    sq = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    sq.paste(img, ((s - img.width) // 2, (s - img.height) // 2))
    return sq


def bold(src: Image.Image, px: int) -> Image.Image:
    if px <= 1:
        return src
    alpha = src.getchannel("A").filter(ImageFilter.MaxFilter(px | 1))
    solid = Image.new("RGBA", src.size, tuple(NAVY) + (255,))
    solid.putalpha(alpha)
    return solid


def icon(mark: Image.Image, size: int, pad: float, radius: float, boldpx: int = 0) -> Image.Image:
    S = size * 8
    base = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(base)
    if radius > 0:
        d.rounded_rectangle((0, 0, S - 1, S - 1), radius=int(S * radius), fill=(255, 255, 255, 255))
    else:
        d.rectangle((0, 0, S - 1, S - 1), fill=(255, 255, 255, 255))
    inner = int(S * (1 - 2 * pad))
    base.alpha_composite(bold(mark, boldpx).resize((inner, inner), Image.LANCZOS), ((S - inner) // 2,) * 2)
    return base.resize((size, size), Image.LANCZOS)


def main() -> None:
    logo = Image.open(SRC).convert("RGBA")
    logo = logo.crop(logo.getbbox())
    # Sidebar logo: shown at 172 CSS px wide → 2x / 3x renditions
    for w in (344, 516):
        h = round(logo.height * w / logo.width)
        r = logo.resize((w, h), Image.LANCZOS)
        r.save(OUT / f"logo-{w}.png", optimize=True)
        r.save(OUT / f"logo-{w}.webp", quality=92, method=6)
    mark = mark_square(logo)
    # Favicons: navy mark on a white rounded square; strokes thickened for tiny sizes
    ico = {n: icon(mark, n, 0.04, 0.2, {16: 13, 32: 7, 48: 5}[n]) for n in (16, 32, 48)}
    ico[48].save(OUT / "favicon.ico", sizes=[(16, 16), (32, 32), (48, 48)], append_images=[ico[16], ico[32]])
    ico[32].save(OUT / "favicon-32.png", optimize=True)
    icon(mark, 180, 0.10, 0, 0).convert("RGB").save(OUT / "apple-touch-icon.png", optimize=True)  # opaque; iOS rounds it
    print("brand assets →", OUT)


if __name__ == "__main__":
    main()
