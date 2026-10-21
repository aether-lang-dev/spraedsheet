// Whole-column and whole-row references (SUM(A:A), COUNTIF($1:$1,"x"), Costs!B:B).
// Found by the IronCalc calc_tests workbooks (docs/comparisons/IRONCALC.md action
// 2): they did not parse. A whole column is read as far as the sheet is used, and
// is worked out again when any cell in it changes, even one past where the sheet
// was used; copying and inserting move it along its own axis only.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
const bad: string[] = [];
for (const line of [
  "formulas.use language=Excel",
  "cell.set cell=A1 input=Item", "cell.set cell=B1 input=Qty", "cell.set cell=C1 input=Price",
  "cell.set cell=A2 input=Tea", "cell.set cell=B2 input=3", "cell.set cell=C2 input=2",
  "cell.set cell=A3 input=Cake", "cell.set cell=B3 input=5", "cell.set cell=C3 input=4",
  "cell.set cell=A4 input=Tea", "cell.set cell=B4 input=7", "cell.set cell=C4 input=1",
  "cell.set cell=F1 input==SUM(B:B)",
  "cell.set cell=F2 input==SUM(B:C)",
  'cell.set cell=F3 input="=COUNTIF(A:A,""Tea"")"',
  'cell.set cell=F4 input="=SUMIF(A:A,""Tea"",B:B)"',
  'cell.set cell=F5 input="=MATCH(""Cake"",A:A,0)"',
  "cell.set cell=F6 input==INDEX(B:B,4)",
  "cell.set cell=F7 input==MAX(B:B)-MIN(B:B)",
  'cell.set cell=F8 input="=VLOOKUP(""Cake"",A:C,3,FALSE)"',
  "cell.set cell=F9 input==SUM(2:2)",
  "cell.set cell=F10 input==SUM($B:$B)+SUM(B:B)",
]) {
  const o = await cmds.run(line);
  if (!o.ok) bad.push(`${line}: ${o.message}`);
}
eq("set up", bad, []);
const col = (k: string[]) => k.map((c) => wb.sheet.value(c));
eq("whole columns and a whole row are read as far as the sheet is used", col(["F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10"]), [15, 22, 2, 10, 3, 7, 4, 4, 27, 30]);

// A cell in the column changes, or one is put far below where the sheet was used.
await cmds.run("cell.set cell=B3 input=50");
await cmds.run("cell.set cell=B200 input=1000");
eq("worked out again when a cell in the column changes, or one is added far below", col(["F1", "F2", "F7"]), [1060, 1067, 1000 - 3]);
await cmds.run("cell.set cell=B200 input=");
eq("and when it is cleared again", wb.sheet.value("F1"), 60);
await cmds.run("cell.set cell=E2 input=9");
eq("a row read whole takes in a cell added to its right", wb.sheet.value("F9"), 81);

// Insert and delete: a whole column moves with columns, not with rows; a whole row the other way.
await cmds.run("rows.insert at=1");
eq("a row inserted: the whole columns are as they were, the whole row moves down", [wb.sheet.input("F2"), wb.sheet.input("F3"), wb.sheet.input("F10")], ["=SUM(B:B)", "=SUM(B:C)", "=SUM(3:3)"]);
eq("and their values are the same", col(["F2", "F3"]).concat([wb.sheet.value("F10")]), [60, 67, 81]);
await cmds.run("columns.insert at=B");
eq("a column inserted before B: B:B is now C:C, B:C grows to C:D, $B:$B too", [wb.sheet.input("G2"), wb.sheet.input("G3"), wb.sheet.input("G11")], ["=SUM(C:C)", "=SUM(C:D)", "=SUM($C:$C)+SUM(C:C)"]);
eq("their values are as before", col(["G2", "G3", "G11"]), [60, 67, 120]);
await cmds.run("columns.delete at=C");
eq("the column deleted: what read only it is #REF!, a span of two shrinks", [wb.sheet.input("F2"), wb.sheet.input("F3")], ["=SUM(#REF!)", "=SUM(C:C)"]);
await cmds.run("edit.undo");
await cmds.run("edit.undo");
await cmds.run("edit.undo");
eq("undo, three times: the sheet as it was before the insert", [wb.sheet.input("F1"), wb.sheet.input("F9"), wb.sheet.value("F1")], ["=SUM(B:B)", "=SUM(2:2)", 60]);

// Copying: a relative whole column moves along, a $ one stays.
await cmds.run("edit.copy range=F1");
await cmds.run("edit.paste at=G1");
await cmds.run("edit.copy range=F10");
await cmds.run("edit.paste at=G10");
eq("copied one column to the right: B:B is C:C, $B:$B stays", [wb.sheet.input("G1"), wb.sheet.input("G10")], ["=SUM(C:C)", "=SUM($B:$B)+SUM(C:C)"]);
await cmds.run("edit.copy range=F9");
await cmds.run("edit.paste at=G12");
eq("a whole row copied down moves down", wb.sheet.input("G12"), "=SUM(5:5)");

// On another sheet, with its name quoted or renamed.
await cmds.run("cell.set cell=A30 input=1");
await cmds.run("cell.set cell=B30 input=2");
await cmds.run('sheet.add name="Other sheet"');
await cmds.run("cell.set cell=A1 input==SUM(Sheet1!B:B)");
await cmds.run("cell.set cell=A2 input==SUM('Sheet1'!30:31)");
eq("another sheet's whole column and rows", [wb.page("Other sheet")!.sheet.value("A1"), wb.page("Other sheet")!.sheet.value("A2")], [62, 3]);
await cmds.run("sheet.activate name=Sheet1");
await cmds.run("cell.set cell=B150 input=40");
eq("worked out again when the cell is on the other sheet", wb.page("Other sheet")!.sheet.value("A1"), 102);
await cmds.run("sheet.rename name=Sheet1 to=Main");
eq("renamed: the reference follows", [wb.page("Other sheet")!.sheet.input("A1"), wb.page("Other sheet")!.sheet.input("A2")], ["=SUM(Main!B:B)", "=SUM(Main!30:31)"]);

// A cell reading its own column is circular, as in Excel.
await cmds.run("sheet.activate name=Main");
await cmds.run("cell.set cell=B20 input==SUM(B:B)");
eq("a cell summing its own column is a circular reference", String(wb.sheet.value("B20")).startsWith("#CYCLE!") || (wb.sheet.value("B20") as { code?: string }).code === "#CYCLE!", true);

// Saved and reopened.
const wb2 = new Workbook(), cmds2 = new Commands(wb2);
await cmds2.execute("file.importText", { text: String((await cmds.execute("file.exportText")).output) });
eq("reopened from its text form", [wb2.sheet.input("F1"), wb2.sheet.input("F9"), wb2.page("Other sheet")!.sheet.input("A1")], ["=SUM(B:B)", "=SUM(2:2)", "=SUM(Main!B:B)"]);

// expect: ok set up
// expect: ok whole columns and a whole row are read as far as the sheet is used
// expect: ok worked out again when a cell in the column changes, or one is added far below
// expect: ok and when it is cleared again
// expect: ok a row read whole takes in a cell added to its right
// expect: ok a row inserted: the whole columns are as they were, the whole row moves down
// expect: ok and their values are the same
// expect: ok a column inserted before B: B:B is now C:C, B:C grows to C:D, $B:$B too
// expect: ok their values are as before
// expect: ok the column deleted: what read only it is #REF!, a span of two shrinks
// expect: ok undo, three times: the sheet as it was before the insert
// expect: ok copied one column to the right: B:B is C:C, $B:$B stays
// expect: ok a whole row copied down moves down
// expect: ok another sheet's whole column and rows
// expect: ok worked out again when the cell is on the other sheet
// expect: ok renamed: the reference follows
// expect: ok a cell summing its own column is a circular reference
// expect: ok reopened from its text form
