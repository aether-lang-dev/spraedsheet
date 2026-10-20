// A workbook as a REXX program (blog-entry-todo.md B.2, app/sheet/bookscript.ts):
// WORKBOOKSCRIPT() written, run in a new workbook, and the two workbooks' text
// forms (B.1) are identical, for each sample and for one with something of each part.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const exportText = async (c: Commands) => String((await c.execute("file.exportText")).output);
const scriptOf = async (c: Commands) => String((await c.execute("file.script")).output);

// The program, run in a new workbook: its text form is the first's.
async function remade(label: string, cmds: Commands): Promise<string> {
  const script = await scriptOf(cmds), text = await exportText(cmds);
  const wb2 = new Workbook();
  const cmds2 = new Commands(wb2);
  const r = await cmds2.script(script, () => {});
  const text2 = await exportText(cmds2);
  const a = text.split("\n"), b = text2.split("\n");
  eq(`${label}: run in a new workbook, the same text form (${r.message})`, [r.message, a.filter((l) => !b.includes(l)).concat(b.filter((l) => !a.includes(l)))], ["ran", []]);
  return script;
}

for (const sample of ["budget", "test", "formatted"]) {
  const wb = new Workbook();
  const cmds = new Commands(wb);
  await cmds.execute("file.open", { sample });
  await remade(`sample ${sample}`, cmds);
}

const wb = new Workbook();
const cmds = new Commands(wb);
for (const line of [
  "formulas.use language=Excel",
  "cell.set cell=A1 input=Item", "cell.set cell=B1 input=2", 'cell.set cell=B2 input="=B1*21"', 'cell.set cell=C3 input="=SUM(B1:B2)"', "cell.comment cell=B2 text=answer",
  "cell.format cell=B1 format=number:0.00", "validation.set range=B1 kind=whole min=0", "names.define name=two range=B1 scope=sheet", "names.define name=ZED range=B2",
  "format.fill range=A1 color=yellow", "format.border range=B2 edge=bottom color=red", "cells.merge range=E1:F1",
  "table.define name=Sales range=A1:B2 headers=1", "chart.insert kind=bar range=A1:B2", "format.colorScale range=B1:B2",
  "rows.hide range=A9:A9", "view.freezeRows count=1", "data.filter range=A1:B2 column=B criterion=\">1\"",
  "sheet.add name=Costs", 'cell.set cell=A1 input="=Sheet1!B2+1"', "sheet.activate name=Sheet1", "macro.record name=Twice", "cell.set cell=D1 input=9", "macro.stop",
]) await cmds.run(line);
await cmds.execute("cell.set", { cell: "A5", input: "two\tcolumns\nand \"lines\"" });
await cmds.execute("cell.set", { cell: "A6", input: "\"quoted" });
await cmds.execute("cell.set", { cell: "A7", input: "C:\\new\\table" });
const script = await remade("a workbook with a bit of each part", cmds);
const lines = script.split("\n");
eq("a program: a comment, then calls, a cell a CALL SETCELL", [lines[0].startsWith("/* Untitled: this workbook as a REXX program"), lines.includes('CALL SETCELL "B2", "=B1*21"'), lines.includes('CALL SETCELL "A6", """quoted"')],
  [true, true, true]);
eq("a newline or tab in a cell by its escape, a backslash doubled", [lines.find((l) => l.startsWith('CALL SETCELL "A5"')), lines.find((l) => l.startsWith('CALL SETCELL "A7"'))], ['CALL SETCELL "A5", "two\\tcolumns\\nand ""lines"""', 'CALL SETCELL "A7", "C:\\\\new\\\\table"']);
eq("the head first and the end last", [lines[1].startsWith('ADDRESS SHEET "file.begin json='), lines[lines.length - 2].startsWith('ADDRESS SHEET "file.end json='), lines.filter((l) => l.startsWith('ADDRESS SHEET "sheet.begin ')).length], [true, true, 2]);
eq("written twice, the same bytes", (await scriptOf(cmds)) === script, true);
eq("a filter, frozen rows and a hidden row are in it", [script.includes('\\"filter\\"') || script.includes('""filter""'), script.includes("frozenRows")], [true, true]);

// It makes the same values, and an edit to the program is an edit to the workbook.
const wb3 = new Workbook();
const cmds3 = new Commands(wb3);
await cmds3.script(script.replace('CALL SETCELL "B1", "2"', 'CALL SETCELL "B1", "3"'), () => {});
eq("edited as REXX: B1 is 3, B2 follows", [wb3.sheet.input("B1"), String(wb3.sheet.value("B2")), wb3.sheetNames()], ["3", "63", ["Sheet1", "Costs"]]);

// REXX.
const said: string[] = [];
await cmds.script('LET p = WORKBOOKSCRIPT()\nSAY LENGTH(p)', (l) => said.push(l));
eq("WORKBOOKSCRIPT() from REXX gives the program", said, [String(script.length)]);

// expect: ok sample budget: run in a new workbook, the same text form (ran)
// expect: ok sample test: run in a new workbook, the same text form (ran)
// expect: ok sample formatted: run in a new workbook, the same text form (ran)
// expect: ok a workbook with a bit of each part: run in a new workbook, the same text form (ran)
// expect: ok a program: a comment, then calls, a cell a CALL SETCELL
// expect: ok a newline or tab in a cell by its escape, a backslash doubled
// expect: ok the head first and the end last
// expect: ok written twice, the same bytes
// expect: ok a filter, frozen rows and a hidden row are in it
// expect: ok edited as REXX: B1 is 3, B2 follows
// expect: ok WORKBOOKSCRIPT() from REXX gives the program
