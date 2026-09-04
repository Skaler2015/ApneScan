#!/usr/bin/env python3
"""Generate the PNG logo assets an MSIX / Microsoft Store package needs, from
the app's existing icon + wide logo.

Usage:
    python make_msix_assets.py <square_src.png> <wide_src.png> <out_dir>

Square tiles come from the 512x512 icon; the wide tile from the wide logo.
All tiles are transparent PNGs with the artwork centred (the manifest's
BackgroundColor shows behind them on the Start tile)."""
import os
import sys

from PIL import Image

SRC_SQUARE = sys.argv[1] if len(sys.argv) > 1 else "apnescan_icon.png"
SRC_WIDE = sys.argv[2] if len(sys.argv) > 2 else "apnescan_logo.png"
OUT = sys.argv[3] if len(sys.argv) > 3 else "Assets"

os.makedirs(OUT, exist_ok=True)
square = Image.open(SRC_SQUARE).convert("RGBA")
wide = Image.open(SRC_WIDE).convert("RGBA")


def make(name, w, h, source, frac=0.90):
    canvas = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    art = source.copy()
    art.thumbnail((max(1, int(w * frac)), max(1, int(h * frac))), Image.LANCZOS)
    canvas.alpha_composite(art, ((w - art.width) // 2, (h - art.height) // 2))
    canvas.save(os.path.join(OUT, name))
    print("  ->", name, "%dx%d" % (w, h))


# Square tiles + store logo (from the square icon)
for name, size in (
    ("Square44x44Logo.png", 44),
    ("Square71x71Logo.png", 71),
    ("Square150x150Logo.png", 150),
    ("Square310x310Logo.png", 310),
    ("StoreLogo.png", 50),
):
    make(name, size, size, square)

# Also the common target-size / unplated variants for the taskbar (nice to have)
for name, size in (
    ("Square44x44Logo.targetsize-24_altform-unplated.png", 24),
    ("Square44x44Logo.targetsize-256.png", 256),
):
    make(name, size, size, square)

# Wide tile (from the wide logo)
make("Wide310x150Logo.png", 310, 150, wide, frac=0.94)

print("MSIX assets written to:", OUT)
