#!/usr/bin/env python3
"""outline.py PNG BX:BY:BW:RW X:Y:W:H -- how many pixels of the grid's cursor
outline (#2f6fd0) lie on the edge band of a cell's box in a screenshot.

The screenshot's scale and offset come from the driver's red banner, BW
points wide at (BX, BY) as /widgets gives it, or, where the banner is not a
solid bar in the picture, from the window, RW points wide. The box (points)
is widened by two points either side, and only its edge band counted.
"""
import struct
import sys
import zlib


def read_png(path):
    data = open(path, "rb").read()
    pos, idat = 8, b""
    while pos < len(data):
        ln, typ = struct.unpack(">I4s", data[pos:pos + 8])
        body = data[pos + 8:pos + 8 + ln]
        if typ == b"IHDR":
            w, h, depth, ctype = struct.unpack(">IIBB", body[:10])
            step = depth // 8
            bpp = (3 if ctype == 2 else 4) * step
        elif typ == b"IDAT":
            idat += body
        pos += 12 + ln
    raw = zlib.decompress(idat)
    stride = w * bpp
    rows, prev, i = [], bytearray(stride), 0
    for _ in range(h):
        f = raw[i]
        line = bytearray(raw[i + 1:i + 1 + stride])
        i += 1 + stride
        for x in range(stride):
            a = line[x - bpp] if x >= bpp else 0
            b = prev[x]
            c = prev[x - bpp] if x >= bpp else 0
            if f == 1: line[x] = (line[x] + a) & 255
            elif f == 2: line[x] = (line[x] + b) & 255
            elif f == 3: line[x] = (line[x] + (a + b) // 2) & 255
            elif f == 4:
                p = a + b - c
                pa, pb, pc = abs(p - a), abs(p - b), abs(p - c)
                line[x] = (line[x] + (a if pa <= pb and pa <= pc else b if pb <= pc else c)) & 255
        rows.append(bytes(line))
        prev = line
    return w, h, bpp, step, rows


def main():
    w, h, bpp, step, rows = read_png(sys.argv[1])
    bx, by, bw, rw = (float(v) for v in sys.argv[2].split(":"))
    px = lambda r, x: (r[x * bpp], r[x * bpp + step], r[x * bpp + 2 * step])
    red = [(x, y) for y in range(0, h // 4) for x in range(w) if px(rows[y], x)[0] > 150 and px(rows[y], x)[1] < 110 and px(rows[y], x)[2] < 110]
    if red:
        l = min(p[0] for p in red); r_ = max(p[0] for p in red); t = min(p[1] for p in red); b = max(p[1] for p in red)
    if red and len(red) > 0.5 * (r_ - l + 1) * (b - t + 1):
        scale = (r_ - l + 1) / bw; ox, oy = l - bx * scale, t - by * scale
    else:
        scale = w / rw; ox, oy = 0.0, 0.0
    x, y, cw, ch = (float(v) for v in sys.argv[3].split(":"))
    x0, y0 = int(ox + (x - 2) * scale), int(oy + (y - 2) * scale)
    x1, y1 = int(ox + (x + cw + 2) * scale), int(oy + (y + ch + 2) * scale)
    band = max(2, int(4 * scale))
    n = 0
    for yy in range(max(0, y0), min(h, y1)):
        r = rows[yy]
        for xx in range(max(0, x0), min(w, x1)):
            if x0 + band <= xx < x1 - band and y0 + band <= yy < y1 - band:
                continue
            cr, cg, cb = px(r, xx)
            if abs(cr - 0x2f) < 40 and abs(cg - 0x6f) < 40 and abs(cb - 0xd0) < 40:
                n += 1
    print(n)


main()
