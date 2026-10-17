#!/bin/sh
# sheet run: command lines in turn, what each gave; --save the result; a
# refused one stops the run with exit 1.
. "$(dirname "$0")/lib.sh"
out=$(sheet run "cell.set cell=A1 input=4" 'cell.set cell=A2 input="=A1*10"' "cell.get cell=A2" --save "$T/r.csv")
has "each command's answer" "$out" "cell.get cell=A2: 40"
check "saved" "$(tr -d '\r' < "$T/r.csv")" "$(printf '4\n40')"
sheet run "cell.set cell=ZZZ9 input=1" > "$T/out" 2> "$T/err"; rc=$?
check "a refused command: exit 1" "$rc" "1"
has "and why" "$(cat "$T/err")" "ZZZ9"
done_
