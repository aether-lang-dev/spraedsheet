// A cross-feature structural-edit workbook (docs/comparisons/LIBREOFFICE.md
// action 1, UNIVER.md 3): one workbook with a table read by another sheet, a
// name, validation, a conditional format, a merge and a chart source, a row
// inserted above them all, the other sheet renamed, undo and redo, saved and
// reopened. Every feature must come through each step as one workbook: what
// they cover moves with the row, and the same text form comes out the other end.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
const setupFailed: string[] = [];
const text = async (c: Commands = cmds) => String((await c.execute("file.exportText")).output);
for (const line of [
  "formulas.use language=Excel",
  "cell.set cell=A1 input=Item", "cell.set cell=B1 input=Amount", "cell.set cell=C1 input=Share",
  "cell.set cell=A2 input=Tea", "cell.set cell=B2 input=10", "cell.set cell=A3 input=Cake", "cell.set cell=B3 input=30", "cell.set cell=A4 input=Soup", "cell.set cell=B4 input=60",
  "cell.set cell=C2 input==B2/B5", "cell.set cell=C3 input==B3/B5", "cell.set cell=C4 input==B4/B5",
  'cell.set cell=B5 input="=SUM(B2:B4)"',
  "table.define name=Sales range=A1:C4 headers=1",
  "names.define name=TOTAL range=B5",
  "names.define name=PART range=B3 scope=sheet",
  "validation.set range=B2:B4 kind=whole min=0",
  "format.highlight range=C2:C4 op=greater value=0.5",
  "cells.merge range=E3:F3",
  "chart.insert kind=bar range=A1:B4",
  "cell.comment cell=B3 text=cake",
  "sheet.add name=Report",
  'cell.set cell=A1 input="=SUM(Sheet1!B2:B4)"',
  'cell.set cell=A2 input="=Sheet1!B3*2"',
  'cell.set cell=A3 input="=TOTAL"',
  "sheet.activate name=Sheet1",
]) {
  const o = await cmds.run(line);
  if (!o.ok) setupFailed.push(`${line}: ${o.message}`);
}
eq("the workbook is set up", setupFailed, []);

const report = () => [wb.page("Report")!.sheet.value("A1"), wb.page("Report")!.sheet.value("A2"), wb.page("Report")!.sheet.value("A3")];
eq("before: the report reads a range, a cell and the name", report(), [100, 60, 100]);
const before = await text();

// What one cell of Sheet1 has besides its input, as cell.info lists it.
const info = async (cell: string): Promise<Record<string, unknown>> => {
  const was = wb.active;
  await cmds.run("sheet.activate name=Sheet1");
  const o = (await cmds.execute("cell.info", { cell })).output as Record<string, unknown>;
  await cmds.run(`sheet.activate name=${was}`);
  return o;
};
const names = () => [...wb.names()].map(([n, r]) => `${n}=${r}`).concat(wb.page("Sheet1")!.sheet.names().map(([n, r]) => `Sheet1.${n}=${r}`));
const tableRows = () => wb.page("Sheet1")!.sheet.tables().map((t) => `${t.name} ${t.range[0].row + 1}-${t.range[1].row + 1}`);
const chartRows = () => wb.page("Sheet1")!.charts.map((c) => `${c.range[0].row + 1}-${c.range[1].row + 1}`);
const mergesOf = () => wb.page("Sheet1")!.merges.map((m) => `${m[0].row + 1}-${m[1].row + 1}`);

await cmds.run("sheet.activate name=Sheet1");
await cmds.run('cell.set cell=E1 input="=Report!A1+1"');
const before1 = await text();
const b3 = await info("B3"), c2 = await info("C2");
eq("before: B3 has its comment and validation, C2 its rule", [JSON.stringify(b3).includes("cake"), JSON.stringify(b3).includes("validation"), JSON.stringify(c2).includes("rule")], [true, true, true]);

// Insert a row above Tea.
await cmds.run("rows.insert at=2");
eq("the cells and the formulas that read them moved", [wb.sheet.input("A3"), wb.sheet.input("B6"), wb.sheet.input("C3"), wb.sheet.input("B2")], ["Tea", "=SUM(B3:B5)", "=B3/B6", ""]);
eq("the report, on another sheet, still reads the right cells", [wb.page("Report")!.sheet.input("A1"), wb.page("Report")!.sheet.input("A2"), ...report()], ["=SUM(Sheet1!B3:B5)", "=Sheet1!B4*2", 100, 60, 100]);
eq("the table, the chart and the merge moved or grew", [tableRows(), chartRows(), mergesOf()], [["Sales 1-5"], ["1-5"], ["4-4"]]);
eq("the workbook's name and the sheet's own name followed their cells", names(), ["TOTAL=Sheet1!B6", "Sheet1.PART=B4"]);
const b4 = await info("B4"), b2 = await info("B2"), b6 = await info("B6"), c3 = await info("C5");
eq("the comment and the validation moved with B3 to B4; the new row has neither", [JSON.stringify(b4).includes("cake"), JSON.stringify(b4).includes("validation"), JSON.stringify(b2).includes("validation"), JSON.stringify(b2).includes("comment"), JSON.stringify(b6).includes("validation")], [true, true, false, false, false]);
eq("the conditional format's range moved down a row: C5 in it, the new C2 and C6 not", [JSON.stringify(c3).includes("rule"), JSON.stringify(await info("C2")).includes("rule"), JSON.stringify(await info("C6")).includes("rule")], [true, false, false]);
const afterInsert = await text();

// The other sheet renamed: what reads it follows.
await cmds.run("sheet.rename name=Report to=Summary");
eq("renamed: Sheet1's formula reads Summary, the name and the report are as they were", [wb.sheet.input("E1"), wb.page("Summary")!.sheet.value("A1"), wb.page("Summary")!.sheet.value("A3"), names()[0]], ["=Summary!A1+1", 100, 100, "TOTAL=Sheet1!B6"]);
const afterRename = await text();

// Undo takes back the rename, then the insert, one step each; redo gives them again.
await cmds.run("edit.undo");
eq("undo: the rename is taken back", await text(), afterInsert);
await cmds.run("edit.undo");
eq("undo again: the insert is taken back, every feature with it", [await text(), names(), tableRows(), chartRows(), mergesOf()], [before1, ["TOTAL=Sheet1!B5", "Sheet1.PART=B3"], ["Sales 1-4"], ["1-4"], ["3-3"]]);
await cmds.run("edit.redo");
await cmds.run("edit.redo");
eq("redo twice: the workbook as it was after the rename", await text(), afterRename);

// Saved and reopened, as its text form, as its file and as its program: the same workbook.
const wb2 = new Workbook(), cmds2 = new Commands(wb2);
await cmds2.execute("file.importText", { text: afterRename });
eq("reopened from its text form: the same text", await text(cmds2), afterRename);
const wb3 = new Workbook(), cmds3 = new Commands(wb3);
await cmds3.execute("file.import", { text: String((await cmds.execute("file.export")).output) });
eq("reopened from its file: the same text", await text(cmds3), afterRename);
const wb4 = new Workbook(), cmds4 = new Commands(wb4);
const ran = await cmds4.script(String((await cmds.execute("file.script")).output), () => {});
eq("run as its program in a new workbook: the same text", [ran.ok, await text(cmds4)], [true, afterRename]);
eq("and it still adds up: the summary reads 100 in all of them", [wb2.page("Summary")!.sheet.value("A1"), wb3.page("Summary")!.sheet.value("A1"), wb4.page("Summary")!.sheet.value("A1")], [100, 100, 100]);

// And a column inserted at the left, in the reopened workbook: the same features move right.
await cmds2.run("sheet.activate name=Sheet1");
await cmds2.run("columns.insert at=A");
const w = wb2.page("Sheet1")!;
eq("a column inserted at A: names, table, merge, chart and the summary's reads moved right", [
  [...wb2.names()].map(([n, r]) => `${n}=${r}`).concat(w.sheet.names().map(([n, r]) => `${n}=${r}`)),
  w.sheet.tables().map((t) => `${t.range[0].col}-${t.range[1].col}`), w.merges.map((m) => `${m[0].col}-${m[1].col}`), w.charts.map((c) => `${c.range[0].col}-${c.range[1].col}`),
  wb2.page("Summary")!.sheet.input("A1"), wb2.page("Summary")!.sheet.value("A3"), wb2.page("Summary")!.sheet.value("A1"),
], [["TOTAL=Sheet1!C6", "PART=C4"], ["1-3"], ["5-6"], ["1-2"], "=SUM(Sheet1!C3:C5)", 100, 100]);

// expect: ok the workbook is set up
// expect: ok before: the report reads a range, a cell and the name
// expect: ok before: B3 has its comment and validation, C2 its rule
// expect: ok the cells and the formulas that read them moved
// expect: ok the report, on another sheet, still reads the right cells
// expect: ok the table, the chart and the merge moved or grew
// expect: ok the workbook's name and the sheet's own name followed their cells
// expect: ok the comment and the validation moved with B3 to B4; the new row has neither
// expect: ok the conditional format's range moved down a row: C5 in it, the new C2 and C6 not
// expect: ok renamed: Sheet1's formula reads Summary, the name and the report are as they were
// expect: ok undo: the rename is taken back
// expect: ok undo again: the insert is taken back, every feature with it
// expect: ok redo twice: the workbook as it was after the rename
// expect: ok reopened from its text form: the same text
// expect: ok reopened from its file: the same text
// expect: ok run as its program in a new workbook: the same text
// expect: ok and it still adds up: the summary reads 100 in all of them
// expect: ok a column inserted at A: names, table, merge, chart and the summary's reads moved right
