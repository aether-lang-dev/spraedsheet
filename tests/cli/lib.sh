# tests/cli/lib.sh: what each test of cli/sheet.js shares. ok/fail lines,
# a folder of its own, and `sheet` (the CLI on node).
ROOT=$(cd "$(dirname "$0")/../.." && pwd)
T=$(mktemp -d "${TMPDIR:-/tmp}/sheet-cli-test.XXXXXX")
trap 'rm -rf "$T"' EXIT
fails=0
sheet() { node "$ROOT/cli/sheet.js" "$@"; }
check() {  # check <what> <got> <want>
    if [ "$2" = "$3" ]; then echo "ok   $1"; else echo "FAIL $1: got [$2], want [$3]"; fails=$((fails + 1)); fi
}
has() {  # has <what> <text> <needle>
    case "$2" in *"$3"*) echo "ok   $1" ;; *) echo "FAIL $1: [$3] not in [$2]"; fails=$((fails + 1)) ;; esac
}
done_() { [ "$fails" -eq 0 ]; }
