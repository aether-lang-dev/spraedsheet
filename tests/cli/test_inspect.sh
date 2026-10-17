#!/bin/sh
# sheet inspect: a workbook's sheets (each used to its last cell), names,
# tables and charts.
. "$(dirname "$0")/lib.sh"
cat > "$T/w.json" <<'JSON'
{"name":"W","formulas":"excel","sheetName":"Costs","cells":{"A1":"Rent","B1":"1200","B2":"300"},"names":{"TOTAL":"B1:B2"},"sheets":[{"name":"Notes","cells":{"A1":"x","C4":"y"}}]}
JSON
out=$(sheet inspect "$T/w.json")
has "opened" "$out" "opened W"
has "the first sheet, used to B2" "$out" "sheet Costs: used to B2"
has "the second, to C4" "$out" "sheet Notes: used to C4"
# A file's "names" are its first sheet's own (bookNames are the workbook's).
has "its name, the sheet's" "$out" "name Costs!TOTAL: B1:B2"
done_
