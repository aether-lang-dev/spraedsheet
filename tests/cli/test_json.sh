#!/bin/sh
# sheet --json: machine-readable results and failures (docs/comparisons/GNUMERIC.md action 5):
# convert says what it left out as a list, a refusal is one JSON object on stderr
# with the command that failed, and the exit status stays 0, 1 (refused) or 2.
. "$(dirname "$0")/lib.sh"
printf 'name,qty\npears,3\n' > "$T/in.csv"
out=$(sheet convert --json "$T/in.csv" "$T/out.xlsx")
check "convert --json: one JSON object, ok, with the files and a list of what was left out" "$(printf '%s' "$out" | node -e 'const o=JSON.parse(require("fs").readFileSync(0,"utf8"));console.log([o.ok,o.command,o.in.endsWith("in.csv"),o.out.endsWith("out.xlsx"),Array.isArray(o.report)].join())')" "true,convert,true,true,true"
# A RexxJS workbook's formulas are not Excel's: the report says so.
sheet run "formulas.use language=RexxJS" "cell.set cell=A1 input=2" 'cell.set cell=B1 input="=A1*2"' --save "$T/rx.json" > /dev/null
out=$(sheet convert --json "$T/rx.json" "$T/rx.xlsx")
has "a loss is listed in the report" "$out" "RexxJS formulas are not Excel"
sheet convert --json "$T/in.csv" "$T/out.doc" > "$T/out" 2> "$T/err"; rc=$?
check "a kind it does not write: exit 1" "$rc" "1"
check "stdout empty and stderr one JSON object: not ok, the command, the error" "$(node -e 'const o=JSON.parse(require("fs").readFileSync(process.argv[1],"utf8"));console.log([o.ok,o.command,o.refused,o.error.includes(".doc")].join())' "$T/err")" "false,convert,true,true"
sheet run --json "cell.set cell=ZZZ9 input=1" > "$T/out" 2> "$T/err"; rc=$?
check "a refused command: exit 1" "$rc" "1"
check "and its line and why, as JSON" "$(node -e 'const o=JSON.parse(require("fs").readFileSync(process.argv[1],"utf8"));console.log([o.ok,o.command,o.line,o.error.includes("ZZZ9")].join())' "$T/err")" "false,run,cell.set cell=ZZZ9 input=1,true"
out=$(sheet run --json "cell.set cell=A1 input=4" "cell.get cell=A1")
check "run --json lists each line with what it gave" "$(printf '%s' "$out" | node -e 'const o=JSON.parse(require("fs").readFileSync(0,"utf8"));console.log([o.ok,o.results.length,o.results[1].line,o.results[1].said].join())')" "true,2,cell.get cell=A1,4"
done_
