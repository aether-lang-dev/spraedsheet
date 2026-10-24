// Replay preview (docs/comparisons/VISIDATA.md action 2): macro.preview shows what a
// macro would do before it is run: each command it runs with the cells it
// changed and whether it was refused, and says whether a failure restores the
// workbook. The workbook is left as it was, with no step of undo.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook(), cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
await cmds.run("cell.set cell=A1 input=1");
await cmds.execute("macro.set", { name: "Fill", text: 'ADDRESS SHEET "cell.set cell=B1 input=5"\nADDRESS SHEET "cell.set cell=B2 input=""=B1*2"""\nADDRESS SHEET "format.font range=B1:B2 bold=1"' });
await cmds.execute("macro.set", { name: "Breaks", text: 'ADDRESS SHEET "cell.set cell=C1 input=9"\nADDRESS SHEET "range.sort range=Z9"\nADDRESS SHEET "cell.set cell=C2 input=10"' });
const undoBefore = wb.undoLabel;

const p = await cmds.execute("macro.preview", { name: "Fill" });
const steps = (p.output as { line: string; ok: boolean; cells: string[] }[]);
eq("a preview lists each command and the cells it changes", [p.ok, steps.map((s) => [s.line.split(" ")[0], s.ok, s.cells])],
  [true, [["cell.set", true, ["B1"]], ["cell.set", true, ["B2"]], ["format.font", true, ["B1", "B2"]]]]);
eq("and says what it did to nothing: the workbook is as it was", [wb.sheet.input("B1"), wb.sheet.input("B2"), wb.undoLabel, wb.canRedo], ["", "", undoBefore, false]);
eq("the message counts them and says nothing was changed", p.message, "Fill would run 3 commands, all accepted; nothing was changed");

const q = await cmds.execute("macro.preview", { name: "Breaks" });
const qs = (q.output as { ok: boolean; message: string }[]);
eq("a macro that fails part way: the failing command is shown, and what a run would keep", [q.ok, qs.map((s) => s.ok), q.message.includes("1 refused"), q.message.includes("a run keeps what it did before the failure")],
  [true, [true, false, true], true, true]);
eq("a preview of a failing macro changes nothing either", [wb.sheet.input("C1"), wb.sheet.input("C2"), wb.undoLabel], ["", "", undoBefore]);

eq("a macro that does not exist", (await cmds.execute("macro.preview", { name: "Nope" })).message, "no macro Nope");
eq("with parameters", (await (async () => {
  await cmds.execute("macro.set", { name: "Put", text: "r = SETCELL(CELL, N * 2)" });
  const o = await cmds.execute("macro.preview", { name: "Put", args: "CELL=H1;N=21" });
  return [o.ok, (o.output as { cells: string[] }[]).map((s) => s.cells), wb.sheet.input("H1")];
})()), [true, [["H1"]], ""]);

// REXX
const said: string[] = [];
await cmds.script('LET m = PREVIEWMACRO("Fill")\nSAY m', (l) => said.push(l));
eq("PREVIEWMACRO from REXX", said, ["Fill would run 3 commands, all accepted; nothing was changed"]);

// expect: ok a preview lists each command and the cells it changes
// expect: ok and says what it did to nothing: the workbook is as it was
// expect: ok the message counts them and says nothing was changed
// expect: ok a macro that fails part way: the failing command is shown, and what a run would keep
// expect: ok a preview of a failing macro changes nothing either
// expect: ok a macro that does not exist
// expect: ok with parameters
// expect: ok PREVIEWMACRO from REXX
