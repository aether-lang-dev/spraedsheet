#!/bin/sh
# sheet convert: CSV to XLSX and back to CSV, the values kept; to the app's
# JSON, HTML and PDF; a kind it does not write refused.
. "$(dirname "$0")/lib.sh"
printf 'name,qty\npears,3\nplums,5\n' > "$T/in.csv"
sheet convert "$T/in.csv" "$T/out.xlsx" > /dev/null
check "an XLSX written (a zip)" "$(head -c 2 "$T/out.xlsx")" "PK"
sheet convert "$T/out.xlsx" "$T/back.csv" > /dev/null
check "and read back as CSV, the same values" "$(tr -d '\r' < "$T/back.csv")" "$(printf 'name,qty\npears,3\nplums,5')"
sheet convert "$T/in.csv" "$T/out.json" > /dev/null
has "the app's JSON" "$(cat "$T/out.json")" "pears"
sheet convert "$T/in.csv" "$T/out.html" > /dev/null
has "HTML, a table" "$(cat "$T/out.html")" "<table"
sheet convert "$T/in.csv" "$T/out.pdf" > /dev/null
check "a PDF" "$(head -c 5 "$T/out.pdf")" "%PDF-"
sheet convert "$T/in.csv" "$T/out.doc" > /dev/null 2> "$T/err"; rc=$?
check "a kind it does not write: exit 1" "$rc" "1"
done_
