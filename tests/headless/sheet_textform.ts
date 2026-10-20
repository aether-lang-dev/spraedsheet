// The canonical text form (blog-entry-todo.md B.1, app/sheet/textform.ts):
// one cell a line in a stable order; every sample workbook, and one with
// something of each part, written as text and read back give the same text
// and the same file; written twice is the same bytes; one cell changed is
// one line changed; EXPORTTEXT and IMPORTTEXT from REXX.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { writeSheetFile } from "./sheet/file.ts";
import { fileOfText } from "./sheet/textform.ts";
import { eq } from "./lib/check.ts";

// The file as JSON with its keys sorted (the text sorts them), compared as text.
const sortKeys = (v: any): any => (Array.isArray(v) ? v.map(sortKeys) : v && typeof v === "object" ? Object.fromEntries(Object.keys(v).sort().map((k) => [k, sortKeys(v[k])])) : v);
const fileJson = (wb: Workbook) => JSON.stringify(sortKeys(writeSheetFile(wb)));
const exportText = async (c: Commands) => String((await c.execute("file.exportText")).output);

// Each sample: written, a new workbook, read back.
for (const sample of ["budget", "test", "formatted"]) {
  const wb = new Workbook();
  const cmds = new Commands(wb);
  await cmds.execute("file.open", { sample });
  const text = await exportText(cmds), before = fileJson(wb);
  await cmds.execute("file.new");
  const o = await cmds.execute("file.importText", { text });
  eq(`sample ${sample}: read back (${o.message}): the same text, the same file`, [o.ok, (await exportText(cmds)) === text, fileJson(wb) === before], [true, true, true]);
  eq(`sample ${sample}: written twice, the same bytes`, (await exportText(cmds)) === text, true);
}

// A workbook with something of each part.
const wb = new Workbook();
const cmds = new Commands(wb);
for (const line of [
  "formulas.use language=Excel",
  "cell.set cell=A1 input=Item", "cell.set cell=B1 input=2", 'cell.set cell=B2 input="=B1*21"', 'cell.set cell=C3 input="=SUM(B1:B2)"', "cell.comment cell=B2 text=answer",
  "cell.format cell=B1 format=number:0.00", "validation.set range=B1 kind=whole min=0", "names.define name=two range=B1 scope=sheet", "names.define name=ZED range=B2",
  "format.fill range=A1 color=yellow", "format.border range=B2 edge=bottom color=red", "cells.merge range=E1:F1",
  "table.define name=Sales range=A1:B2 headers=1", "chart.insert kind=bar range=A1:B2", "format.colorScale range=B1:B2",
  "sheet.add name=Costs", 'cell.set cell=A1 input="=Sheet1!B2+1"', "sheet.activate name=Sheet1", "macro.record name=Twice", "cell.set cell=D1 input=9", "macro.stop",
]) await cmds.run(line);
await cmds.execute("cell.set", { cell: "A5", input: "two\tcolumns\nand \"lines\"" });
await cmds.execute("cell.set", { cell: "A6", input: "\"quoted" });
await cmds.execute("cell.set", { cell: "A7", input: " padded " });

const text = await exportText(cmds);
const lines = text.split("\n");
const line = (key: string) => lines.find((l) => l.startsWith(`${key}\t`));
eq("the header, then the workbook's settings, then each sheet", [lines[0], lines.filter((l) => l.startsWith("@sheet")), text.endsWith("\n")],
  ["# rexxsheet text 1", ["@sheet\tSheet1", "@sheet\tCosts"], true]);
eq("a constant is its input; a formula as typed, with no cached value", [lines.includes("C3\t=SUM(B1:B2)"), lines.includes("D1\t9"), line("A1")!.startsWith("A1\tItem\t{")], [true, true, true]);
eq("a formula's line has its expression, its notes after, and no value", line("B2"), 'B2\t=B1*21\t{"comment":"answer","look":{"borders":{"bottom":"#ff0000"}}}');
eq("cells are in rows then columns", lines.filter((l) => /^[A-Z]+[0-9]+\t/.test(l)).slice(0, 7).map((l) => l.split("\t")[0]), ["A1", "B1", "D1", "B2", "C3", "A5", "A6"]);
eq("text with a tab or newline, a leading quote or blanks is a JSON string", [line("A5"), line("A6"), line("A7")],
  ['A5\t"two\\tcolumns\\nand \\"lines\\""', 'A6\t"\\"quoted"', 'A7\t" padded "']);
eq("names sorted, a list's elements a line apiece", [lines.filter((l) => l.startsWith("@bookNames") || l.startsWith("@names")), lines.filter((l) => l.startsWith("@rules[]")).length, lines.filter((l) => l.startsWith("@tables[]")).length, lines.filter((l) => l.startsWith("@charts[]")).length],
  [['@bookNames\tZED\t"Sheet1!B2"', '@names\tTWO\t"B1"'], 1, 1, 1]);

const before = fileJson(wb);
eq("written twice, the same bytes", (await exportText(cmds)) === text, true);

// Read into a fresh workbook: the same text, the same file, the same values.
const wb2 = new Workbook();
const cmds2 = new Commands(wb2);
const o = await cmds2.execute("file.importText", { text });
eq("read back: the same text, the same file", [o.ok, (await exportText(cmds2)) === text, fileJson(wb2) === before], [true, true, true]);
eq("and worked out", [wb2.sheet.input("B2"), String(wb2.sheet.value("B2")), wb2.sheet.input("A5"), wb2.sheet.input("A6"), wb2.sheet.input("A7"), wb2.sheetNames()], ["=B1*21", "42", "two\tcolumns\nand \"lines\"", "\"quoted", " padded ", ["Sheet1", "Costs"]]);

// One cell changed, one line changed.
await cmds.execute("cell.set", { cell: "B1", input: "3" });
const changed = (await exportText(cmds)).split("\n");
eq("one cell changed is one line changed", [changed.length === lines.length, lines.filter((l, i) => l !== changed[i]).length, changed.find((l, i) => l !== lines[i])?.split("\t")[0]], [true, 1, "B1"]);
await cmds.execute("cell.set", { cell: "H9", input: "new" });
const added = (await exportText(cmds)).split("\n");
eq("and a cell added is one line added", [added.length - changed.length, added.filter((l) => !changed.includes(l))], [1, ["H9\tnew"]]);

// Refused, with the line.
for (const [label, bad] of [["not a text", "hello"], ["a cell before a sheet", "# rexxsheet text 1\nA1\tx\n"], ["a bad JSON line", "# rexxsheet text 1\n@sheet\tS\n@names\tX\t{oops\n"], ["no sheet", "# rexxsheet text 1\n@name\t\"x\"\n"], ["not a cell", "# rexxsheet text 1\n@sheet\tS\nnope\tx\n"]]) {
  const r = await cmds2.execute("file.importText", { text: bad });
  eq(`refused: ${label}`, [r.ok, r.message.length > 0], [false, true]);
}
eq("parsed back, a file is readSheetFile's", Object.keys(fileOfText(text)).includes("sheetName"), true);

// REXX.
const said: string[] = [];
const r = await cmds2.script('LET t = EXPORTTEXT()\nSAY LENGTH(t)\nLET r = SETCELL("A1", "changed")\nLET r = IMPORTTEXT(t)\nSAY GETCELL("A1")', (l) => said.push(l));
eq("EXPORTTEXT and IMPORTTEXT from REXX", [r.message, said], ["ran", [String(text.length), "Item"]]);

// expect: ok sample budget: read back (opened Sample Budget: 20 cells, RexxJS formulas): the same text, the same file
// expect: ok sample budget: written twice, the same bytes
// expect: ok sample test: read back (opened Test Sheet: 10 cells, RexxJS formulas): the same text, the same file
// expect: ok sample test: written twice, the same bytes
// expect: ok sample formatted: read back (opened Formatted Spreadsheet Example: 50 cells, RexxJS formulas): the same text, the same file
// expect: ok sample formatted: written twice, the same bytes
// expect: ok the header, then the workbook's settings, then each sheet
// expect: ok a constant is its input; a formula as typed, with no cached value
// expect: ok a formula's line has its expression, its notes after, and no value
// expect: ok cells are in rows then columns
// expect: ok text with a tab or newline, a leading quote or blanks is a JSON string
// expect: ok names sorted, a list's elements a line apiece
// expect: ok written twice, the same bytes
// expect: ok read back: the same text, the same file
// expect: ok and worked out
// expect: ok one cell changed is one line changed
// expect: ok and a cell added is one line added
// expect: ok refused: not a text
// expect: ok refused: a cell before a sheet
// expect: ok refused: a bad JSON line
// expect: ok refused: no sheet
// expect: ok refused: not a cell
// expect: ok parsed back, a file is readSheetFile's
// expect: ok EXPORTTEXT and IMPORTTEXT from REXX
