#!/bin/sh
# function-status.sh -- docs/function_status.md, generated (docs/comparisons/GNUMERIC.md
# action 1) by tests/status/function_status.ts, run on sae's engine as the
# headless tests are (sae's lower/run-tests.sh --mod; sae built as its AGENTS.md
# says). --check: say whether the committed page is current, and fail if not.
ROOT=$(cd "$(dirname "$0")/.." && pwd)
SAE=$(cd "${SAE_DIR:-$ROOT/sae}" 2>/dev/null && pwd -P) || { echo "function-status.sh: no sae at ${SAE_DIR:-$ROOT/sae}" >&2; exit 2; }
OUT=$ROOT/target/function-status
rm -rf "$OUT"; mkdir -p "$OUT"
# The page has no "// expect:" lines, so the harness calls it a failure; what
# it printed is the page.
SAE_LOWER_OUT=$OUT SAE_PAGE_MEM=536870912 SAE_TIME_LIMIT=60000 "$SAE/lower/run-tests.sh" --mod "$ROOT/tests/status" >"$OUT/harness.log" 2>&1
PAGE=$OUT/mod_function_status.actual
if ! [ -s "$PAGE" ] || ! head -1 "$PAGE" | grep -q '^# Function status$' || grep -q '^Error\|^Uncaught\|^    at ' "$PAGE"; then
    echo "function-status.sh: the generator did not print the page (see $OUT)" >&2
    cat "$PAGE" "$OUT/harness.log" >&2 2>/dev/null | tail -20
    exit 1
fi
DOC=$ROOT/docs/function_status.md
if [ "${1:-}" = --check ]; then
    if cmp -s "$PAGE" "$DOC"; then echo "docs/function_status.md is current"; exit 0; fi
    echo "docs/function_status.md is not current: run tools/function-status.sh and commit it" >&2
    diff "$DOC" "$PAGE" | head -20 >&2
    exit 1
fi
cp "$PAGE" "$DOC"
echo "wrote docs/function_status.md ($(wc -l < "$DOC") lines)"
