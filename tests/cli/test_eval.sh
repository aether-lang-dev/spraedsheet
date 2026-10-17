#!/bin/sh
# sheet eval: a formula's value; with --file, over that file's cells.
. "$(dirname "$0")/lib.sh"
check "a sum" "$(sheet eval '=SUM(1,2,3)*2')" "12"
printf 'a,b\n1,2\n3,4\n' > "$T/n.csv"
check "over a file's cells" "$(sheet eval '=SUM(A2:B3)' --file "$T/n.csv")" "10"
sheet eval '=1+' > "$T/out" 2> "$T/err"; rc=$?
check "a formula that does not parse: exit 1" "$rc" "1"
has "and why, on stderr" "$(cat "$T/err")" "sheet:"
done_
