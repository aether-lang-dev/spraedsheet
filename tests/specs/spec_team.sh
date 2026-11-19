#!/bin/sh
# spec_team.sh — two sae browsers on one team workbook (spec_team.ae): builds
# and starts sheet-server on the page-server port with a fresh data folder and
# two people, starts Bob's sae-driver on the driver port + 1, and runs the spec
# through sae's run_spec.sh with Ann's sae-driver on the driver port.
#   SAE_TEST_SITE_PORT (default 8091)  the server's port
#   AETHER_UI_TEST_PORT (default 9222) Ann's driver; Bob's is the next one
ROOT=$(cd "$(dirname "$0")/../.." && pwd)
SAE=$(cd "${SAE_DIR:-$ROOT/sae}" && pwd -P)
SITE_PORT=${SAE_TEST_SITE_PORT:-8091}
DRIVER_PORT=${AETHER_UI_TEST_PORT:-9222}
BOB_PORT=$((DRIVER_PORT + 1))
T=$ROOT/target/spec-team
rm -rf "$T" && mkdir -p "$T"
# Built aside and moved into place: macOS kills a binary rewritten in place while a copy of it runs.
ae build "$ROOT/server/sheetserver.ae" -o "$ROOT/target/sheet-server.new" >"$T/build.log" 2>&1 && mv -f "$ROOT/target/sheet-server.new" "$ROOT/target/sheet-server" || { cat "$T/build.log"; exit 2; }
ANN=$("$ROOT/target/sheet-server" add-person "$T/data" ann)
BOB=$("$ROOT/target/sheet-server" add-person "$T/data" bob)
if curl -s -o /dev/null "http://127.0.0.1:$SITE_PORT/" || curl -s -o /dev/null "http://127.0.0.1:$BOB_PORT/widgets"; then
    echo "spec_team.sh: something already answers on $SITE_PORT or $BOB_PORT" >&2
    exit 2
fi
"$ROOT/target/sheet-server" serve "$T/data" "$ROOT/app" "$SITE_PORT" >"$T/server.log" 2>&1 &
SERVER=$!
i=0
while [ $i -lt 100 ] && ! curl -s -o /dev/null "http://127.0.0.1:$SITE_PORT/api/info"; do sleep 0.1; i=$((i + 1)); done
URL="http://127.0.0.1:$SITE_PORT/sheet?book=budget"
AETHER_UI_TEST_PORT=$BOB_PORT SAE_STORAGE_DIR="$T/bob-storage" SAE_PAGE_TIME_MS=5000 \
    "${SAE_TEST_BIN:-$SAE/target/build/bin/sae-driver}" "$URL" >"$T/bob.log" 2>&1 &
BOBSAE=$!
trap 'kill $SERVER $BOBSAE 2>/dev/null || true' EXIT INT TERM
status=0
SHEET_TEAM_ANN=$ANN SHEET_TEAM_BOB=$BOB SHEET_TEAM_BOB_PORT=$BOB_PORT SHEET_TEAM_BOB_LOG="$T/bob.log" \
SHEET_TEAM_DIR="$T" SHEET_TEAM_DATA="$T/data" \
SAE_TEST_SERVER=external SAE_TEST_START="$URL" SAE_STORAGE_DIR="$T/ann-storage" SAE_PAGE_TIME_MS=5000 \
SAE_TEST_LIB_DIR="$ROOT/tests/lib" AETHER_UI_TEST_PORT=$DRIVER_PORT SAE_TEST_SITE_PORT=$SITE_PORT \
    "$SAE/tests/run_spec.sh" "$ROOT/tests/specs/spec_team.ae" "$ROOT/app" || status=$?
echo "spec_team: server log $T/server.log, Bob's sae $T/bob.log, Ann's $SAE/target/spec_team.log, screenshots $T/shots"
exit $status
