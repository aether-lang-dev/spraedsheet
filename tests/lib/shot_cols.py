#!/usr/bin/env python3
"""shot_cols.py OLD.png NEW.png Y0 Y1 -- how two screenshots' strip of rows Y0 to
Y1 differ, column by column: the runs of columns that are the same, and
those inserted, deleted or changed. Printed a run a line (`equal`, `insert`,
`delete`, `replace`, then the columns in each picture).

For re-recording a baseline (tests/specs/baselines) after a change to what a
strip shows, such as a toolbar with buttons added: the new baseline differs
from one recorded of the old code only by `insert` runs, as wide as the
buttons added, where no `replace` run stands. PNG decoding is sae's
(tests/lib/png_mae.py, beside this checkout as ../sae)."""
import difflib
import os
import sys
here = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(here, "..", "..", "sae", "tests", "lib"))
from png_mae import read_png  # noqa: E402


def columns(img, y0, y1):
    w, h, bpp, rows = img
    return [bytes(c for y in range(y0, y1) for c in rows[y][x * bpp:(x + 1) * bpp]) for x in range(w)]


def main():
    a, b = read_png(sys.argv[1]), read_png(sys.argv[2])
    y0, y1 = int(sys.argv[3]), int(sys.argv[4])
    sm = difflib.SequenceMatcher(None, columns(a, y0, y1), columns(b, y0, y1), autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        print("%s old %d-%d new %d-%d" % (tag, i1, i2, j1, j2))


main()
