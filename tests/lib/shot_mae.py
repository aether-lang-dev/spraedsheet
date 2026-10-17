#!/usr/bin/env python3
"""shot_mae.py REF.png SHOT.png X Y W H [MAXERR]

Mean absolute error over R, G and B between two screenshots (the driver's
/screenshot PNGs) inside the rectangle X, Y, W, H (in the screenshots'
pixels), and the share of pixels that differ by more than 48 in any
channel. Prints `mae <value> differ <fraction>`; exits 1 when MAXERR is
given and the error is over it, 2 when either image is smaller than the
rectangle. PNG decoding is sae's (tests/lib/png_mae.py, beside this
checkout as ../sae), pure Python with zlib only."""
import os, sys
here = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(here, "..", "..", "sae", "tests", "lib"))
from png_mae import read_png  # noqa: E402

def rgb(img, x, y):
    w, h, bpp, rows = img
    px = rows[y][x * bpp:(x + 1) * bpp]
    if bpp >= 3: return px[0], px[1], px[2]
    return (px[0],) * 3

def main():
    ref, shot = read_png(sys.argv[1]), read_png(sys.argv[2])
    x, y, w, h = (int(v) for v in sys.argv[3:7])
    maxerr = float(sys.argv[7]) if len(sys.argv) > 7 else None
    for img in (ref, shot):
        if img[0] < x + w or img[1] < y + h:
            print("an image is smaller than the rectangle"); sys.exit(2)
    total, far, n = 0, 0, 0
    for j in range(y, y + h):
        for i in range(x, x + w):
            a, b = rgb(ref, i, j), rgb(shot, i, j)
            d = [abs(p - q) for p, q in zip(a, b)]
            total += sum(d); n += 3
            if max(d) > 48: far += 1
    mae = total / n
    print("mae %.3f differ %.4f" % (mae, far / (w * h)))
    if maxerr is not None and mae > maxerr: sys.exit(1)

main()
