#!/bin/sh
# real_keys_x11.sh PORT: the sheet (the installed app) driven by REAL X11
# input (xdotool, XTEST), as a person's keyboard reaches GTK: a click on A6,
# then arrows, Tab and Shift+Tab. After each key the driver reads the Name
# Box's cell and the focus: the cursor moved, and the focus stayed the
# grid's. (A real arrow used to move it once, then GTK moved the focus on to
# the next control; aether-ui's key_claim.) Run under an X server
# (xvfb-run). Exits 0 on a pass, 1 on a fail, 77 without xdotool.
PORT=${1:-9399}; U=http://127.0.0.1:$PORT
command -v xdotool > /dev/null 2>&1 || { echo "  SKIP no xdotool"; exit 77; }
ROOT=$(cd "$(dirname "$0")/../.." && pwd)
cd "$ROOT" || exit 1
AETHER_UI_TEST_PORT=$PORT SAE_PAGE_MEM=536870912 ./sae/target/build/bin/sae-driver --app app > /tmp/real_keys_x11.app.log 2>&1 &
AP=$!
for i in $(seq 1 120); do curl -sf -o /dev/null $U/widgets && break; sleep 0.25; done
sleep 2
WID=$(xdotool search --pid $AP 2>/dev/null | tail -1)
fail=0
q() { curl -s $U/widgets | python3 -c "
import json,re,sys
W=json.load(sys.stdin)
def cell(col,row):
    h=[w['id'] for w in W if 'aui-grid-header' in (w.get('classes') or '') and w.get('text')==col]
    for w in W:
        if h and w.get('a11y_header')==h[-1] and w.get('a11y_row')==row: return w
    return None
if sys.argv[1]=='at':
    c=cell(sys.argv[2],int(sys.argv[3])); print(c['x']+c['w']//2, c['y']+c['h']//2)
else:
    nb=[w.get('text') for w in W if w.get('type')=='textfield' and re.match(r'^[A-Z]+[0-9]+$', w.get('text') or '')]
    print(nb[0] if nb else '-')" "$@"; }
focus() { curl -s $U/focus | python3 -c 'import json,sys; print(json.load(sys.stdin).get("handle"))'; }
run() {   # COL ROW then "KEY WANT" pairs
    col=$1; row=$2; shift 2
    set -- $(q at "$col" "$row") "$@"
    xdotool windowfocus --sync "$WID" 2>/dev/null
    xdotool mousemove --window "$WID" "$1" "$2" click 1
    sleep 0.8
    shift 2
    grid=$(focus)
    for step in "$@"; do
        key=${step% *}; want=${step#* }
        xdotool key "$key"
        sleep 0.5
        got=$(q nb); f=$(focus)
        [ "$got" = "$want" ] || { echo "  FAIL from $col$row, $key: the cursor at $got, not $want"; fail=1; }
        [ "$f" = "$grid" ] || { echo "  FAIL from $col$row, $key: the focus left the grid (on $f, not $grid)"; fail=1; }
    done
}
run A 6 "Up A5" "Up A4" "Up A3"
run A 6 "Down A7" "Down A8" "Down A9"
run B 6 "Left A6" "Right B6" "Right C6"
run B 6 "Tab C6" "Tab D6" "ISO_Left_Tab C6"
kill $AP 2>/dev/null
[ $fail -eq 0 ] && echo "  OK   real arrows, Tab and Shift+Tab moved the cursor; the focus stayed the grid's"
exit $fail
