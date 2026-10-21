// What a script reads in the middle of a batch (docs/comparisons/HYPERFORMULA.md
// action 2), said and held: every command takes effect at once, and a read after
// a write sees the written state, formulas worked out (read-after-write: the
// committed state, not a snapshot from before the script); under manual
// calculation a formula stays as it was until CALCULATE(); and a script that
// fails puts the whole workbook back, so nothing it read mid-way was kept.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
const said: string[] = [];
const r1 = await cmds.script('LET a = SETCELL("A1", 1)\nLET b = SETCELL("B1", "=A1*2")\nSAY GETCELL("B1")\nLET c = SETCELL("A1", 5)\nSAY GETCELL("B1")\nLET d = SETCELL("C1", "=B1+A1")\nSAY GETCELL("C1")', (l) => said.push(l));
eq("automatic: a read after a write sees the formulas worked out each time", [r1.ok, said], [true, ["2", "10", "15"]]);

// Manual calculation: a formula keeps its old value until it is worked out.
said.length = 0;
await cmds.run("formulas.calculation mode=manual");
const r2 = await cmds.script('LET a = SETCELL("A1", 7)\nSAY GETCELL("B1")\nSAY GETCELL("A1")\nLET c = CALCULATE()\nSAY GETCELL("B1")', (l) => said.push(l));
eq("manual: the write is read at once, the formula that reads it is as it was until CALCULATE()", [r2.ok, said], [true, ["10", "7", "14"]]);
await cmds.run("formulas.calculation mode=automatic");

// A script run from the panel (script()) keeps what it did before it failed, as one step of undo;
// a request from the bus (spreadsheet(), the MCP tools) is all or nothing.
said.length = 0;
const r3 = await cmds.script('LET a = SETCELL("A1", 100)\nSAY GETCELL("B1")\nLET b = NO_SUCH_FUNCTION()', (l) => said.push(l));
eq("a failing script keeps what it did, read at once (200), as one undo step", [r3.ok, said, wb.sheet.input("A1"), wb.sheet.value("B1")], [false, ["200"], "100", 200]);
await cmds.run("edit.undo");
eq("and one Undo takes all of it back", [wb.sheet.input("A1"), wb.sheet.value("B1")], ["7", 14]);

said.length = 0;
const r4 = await cmds.spreadsheet('SETCELL("A1", "100"); BOGUS()', (l) => said.push(l));
eq("a failing bus request: taken back whole, and told so", [r4.ok, r4.message.endsWith("(nothing it did was kept)"), wb.sheet.input("A1"), wb.sheet.value("B1")], [false, true, "7", 14]);

// expect: ok automatic: a read after a write sees the formulas worked out each time
// expect: ok manual: the write is read at once, the formula that reads it is as it was until CALCULATE()
// expect: ok a failing script keeps what it did, read at once (200), as one undo step
// expect: ok and one Undo takes all of it back
// expect: ok a failing bus request: taken back whole, and told so
