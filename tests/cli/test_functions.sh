#!/bin/sh
# sheet functions: the formula functions, sorted (by code point: names() sorts so, and the shell's sort is pinned to it, not to a locale that ignores the dots).
. "$(dirname "$0")/lib.sh"
out=$(sheet functions)
has "SUM" "$out" "SUM"
has "XLOOKUP" "$out" "XLOOKUP"
check "sorted" "$(printf '%s\n' "$out" | LC_ALL=C sort -c 2>&1 && echo sorted)" "sorted"
done_
