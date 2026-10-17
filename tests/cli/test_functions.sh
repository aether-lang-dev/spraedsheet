#!/bin/sh
# sheet functions: the formula functions, sorted.
. "$(dirname "$0")/lib.sh"
out=$(sheet functions)
has "SUM" "$out" "SUM"
has "XLOOKUP" "$out" "XLOOKUP"
check "sorted" "$(printf '%s\n' "$out" | sort -c 2>&1 && echo sorted)" "sorted"
done_
