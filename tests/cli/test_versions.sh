#!/bin/sh
# Versions in a git repository (app/sheet/gitstore.ts): two saves from one
# run of the app, then real git reads what the app wrote -- the log, the
# diff, the file at each commit, and git's own consistency check.
. "$(dirname "$0")/lib.sh"
command -v git >/dev/null 2>&1 || { echo "SKIP versions (no git)"; exit 0; }
export SHEET_APPDATA="$T/appdata"
out=$(sheet run "cell.set cell=A1 input=Item" "cell.set cell=B1 input=2" "version.save message=First" "cell.set cell=B1 input=3" "cell.set cell=C3 input=new" "version.diff" "version.save message=Second" "version.list")
has "a save gives the commit's id" "$out" "version.save message=First: $(git --git-dir="$SHEET_APPDATA/versions.git" rev-parse HEAD~1)"
R="$SHEET_APPDATA/versions.git"
check "git fsck: the objects are git's" "$(git --git-dir="$R" fsck 2>&1 | grep -v '^Checking\|^notice' | head -3)" ""
check "git log: two versions, newest first" "$(git --git-dir="$R" log --format=%s)" "$(printf 'Second\nFirst')"
check "git show: the first version's file" "$(git --git-dir="$R" show HEAD~1:workbook.sheet.txt | grep -c '^B1	2$')" "1"
check "git diff names the changed cell line" "$(git --git-dir="$R" diff HEAD~1 HEAD | grep -c '^[-+]B1	')" "2"
check "git diff names the added cell" "$(git --git-dir="$R" diff HEAD~1 HEAD | grep -c '^+C3	new$')" "1"
check "git cat-file: the commit's parent is the first" "$(git --git-dir="$R" rev-parse HEAD~1)" "$(git --git-dir="$R" cat-file -p HEAD | sed -n 's/^parent //p')"
done_
