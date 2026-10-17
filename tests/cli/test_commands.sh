#!/bin/sh
# sheet commands: every control function's signature; one described.
. "$(dirname "$0")/lib.sh"
out=$(sheet commands)
has "SETCELL among them" "$out" "SETCELL("
has "SORTRANGE's signature" "$out" "SORTRANGE("
check "one, described" "$(sheet commands getcell | head -1 | cut -c1-8)" "GETCELL("
sheet commands NOSUCH > /dev/null 2>&1; rc=$?
check "an unknown one: exit 1" "$rc" "1"
done_
