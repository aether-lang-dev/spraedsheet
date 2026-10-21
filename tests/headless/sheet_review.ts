// One request, one reviewable diff (blog-entry-todo.md B.5, the review.*
// commands): a hold takes the workbook's text as it is, whatever is done
// after it is the diff against that (B.4's diffText, cell by cell), and the
// changes are kept whole or taken back whole.
import { Commands } from "./sheet/commands.ts";
import type { Store } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { changeLine } from "./sheet/gitstore.ts";
import { eq } from "./lib/check.ts";

const store = (): Store => ({ get: () => null, set: () => {}, remove: () => {} });
const wb = new Workbook();
const cmds = new Commands(wb, store());
await cmds.run("formulas.use language=Excel");
await cmds.run("cell.set cell=A1 input=Item");
await cmds.run("cell.set cell=B1 input=2");
await cmds.run('cell.set cell=B2 input="=B1*21"');

eq("no hold: a diff is refused, with why", [(await cmds.execute("review.diff")).ok, (await cmds.execute("review.diff")).message], [false, "nothing is held for review: review.hold first"]);
eq("no hold: keep and take back are refused too", [(await cmds.execute("review.keep")).ok, (await cmds.execute("review.takeback")).ok], [false, false]);

const before = (await cmds.execute("file.exportText")).output;
const held = await cmds.execute("review.hold");
eq("a hold", [held.ok, held.message], [true, "holding the workbook as it is for review"]);
eq("a hold with nothing done: no changes", (await cmds.execute("review.diff")).output as unknown[], []);

await cmds.run("cell.set cell=B1 input=3");
await cmds.run("cell.set cell=C3 input=new");
await cmds.run("cell.comment cell=A1 text=heading");
const diff = await cmds.execute("review.diff");
eq("the diff names the changed cells, old and new", (diff.output as { where: string; was: string; now: string }[]).map(changeLine),
  ["Sheet1!A1: Item → Item {\"comment\":\"heading\"}", "Sheet1!B1: 2 → 3", "Sheet1!C3: added new"]);
eq("held again while one is held: the same hold, not a new baseline", [(await cmds.execute("review.hold")).message, (await cmds.execute("review.diff")).output as unknown[]].map((x) => (Array.isArray(x) ? x.length : x)), ["already held: 3 changes so far", 3]);
const strict = await cmds.execute("review.hold", { strict: "1" });
eq("a strict hold while one is held is refused", [strict.ok, strict.message], [false, "a review is waiting: 3 changes; keep them or take them back first"]);

// Taken back whole.
const back = await cmds.execute("review.takeback");
eq("taken back: the message", [back.ok, back.message], [true, "took back 3 changes"]);
eq("taken back: the workbook as it was held", [wb.sheet.input("B1"), wb.sheet.input("C3"), (await cmds.execute("file.exportText")).output === before], ["2", "", true]);
eq("and the hold is gone", (await cmds.execute("review.diff")).ok, false);

// Kept whole.
await cmds.execute("review.hold");
await cmds.run("cell.set cell=B1 input=5");
const kept = await cmds.execute("review.keep");
eq("kept: the message", [kept.ok, kept.message], [true, "kept 1 change"]);
eq("kept: the workbook as it is, and the hold gone", [wb.sheet.input("B1"), (await cmds.execute("review.diff")).ok], ["5", false]);

// A strict hold with none held is a hold.
eq("a strict hold with none held", [(await cmds.execute("review.hold", { strict: "1" })).ok, (await cmds.execute("review.keep")).message], [true, "kept 0 changes"]);

// REXX.
const said: string[] = [];
await cmds.script('LET h = HOLDREVIEW()\nLET r = SETCELL("B1", 9)\nLET d = REVIEWDIFF()\nSAY d.0 d.1\nLET k = TAKEBACKREVIEW()\nSAY k\nLET v = GETCELL("B1")\nSAY v\nLET h = HOLDREVIEW()\nLET r = SETCELL("B1", 7)\nLET k = KEEPREVIEW()\nSAY k\nSAY GETCELL("B1")', (l) => said.push(l));
eq("HOLDREVIEW, REVIEWDIFF, TAKEBACKREVIEW and KEEPREVIEW from REXX", said, ["1 Sheet1!B1: 5 → 9", "took back 1 change", "5", "kept 1 change", "7"]);

// expect: ok no hold: a diff is refused, with why
// expect: ok no hold: keep and take back are refused too
// expect: ok a hold
// expect: ok a hold with nothing done: no changes
// expect: ok the diff names the changed cells, old and new
// expect: ok held again while one is held: the same hold, not a new baseline
// expect: ok a strict hold while one is held is refused
// expect: ok taken back: the message
// expect: ok taken back: the workbook as it was held
// expect: ok and the hold is gone
// expect: ok kept: the message
// expect: ok kept: the workbook as it is, and the hold gone
// expect: ok a strict hold with none held
// expect: ok HOLDREVIEW, REVIEWDIFF, TAKEBACKREVIEW and KEEPREVIEW from REXX
