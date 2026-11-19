// Implicit intersection on import (TODO.md, decided 2026-11-18). A formula a file stores
// with no array flag is intersected implicitly by Excel, which shows it with an @
// (=@A1:A4, =ROMAN(@A1:A4), =@INDEX(...)) and works it out to one value; this app
// spilled it. The importer puts the @ in where Excel's loader does, the writer takes
// it out again, so the formula comes back as the legacy one it was. Found by the
// IronCalc calc_tests workbooks (their 47 `legacy` cells; docs/comparisons/COMPATIBILITY.md).
import { intersected, stored } from "./sheet/implicit.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { readXlsx } from "./sheet/xlsx.ts";
import { writeXlsx } from "./sheet/xlsxwrite.ts";
import { unzip, utf8Bytes, utf8Text, zip } from "./sheet/zip.ts";
import { eq } from "./lib/check.ts";

const names = new Set(["QTY"]);
const cases: [string, string][] = [
  ["A1+B1", "A1+B1"],
  ["A1:A3", "@A1:A3"],
  ["A1:A1", "A1:A1"],
  ["SUM(A1:A3)", "SUM(A1:A3)"],
  ["SUM(A1:A3*2)", "SUM(@A1:A3*2)"],
  ["SUMPRODUCT(A1:A3*B1:B3)", "SUMPRODUCT(A1:A3*B1:B3)"],
  ["ROMAN(H34:H37,1)", "ROMAN(@H34:H37,1)"],
  ["EXP(H:H)", "EXP(@H:H)"],
  ["EXPON.DIST(F:F,G:G,C56)", "EXPON.DIST(@F:F,@G:G,C56)"],
  ["INDEX(A:A,MATCH(1,B:B,0))", "@INDEX(A:A,MATCH(1,B:B,0))"],
  ["SUM(INDEX(A1:C3,0,2))", "SUM(INDEX(A1:C3,0,2))"],
  ["VLOOKUP(A1,B1:C9,2,FALSE)", "VLOOKUP(A1,B1:C9,2,FALSE)"],
  ["IF(A1:A3>1,1,2)", "IF(@A1:A3>1,1,2)"],
  ["CHOOSE(A25:A27,A4,A5,A6)", "CHOOSE(@A25:A27,A4,A5,A6)"],
  ['CONCATENATE(A1:A3,"!")', 'CONCATENATE(@A1:A3,"!")'],
  ['"A1:A3"&A1', '"A1:A3"&A1'],
  ["-A1:A3", "-@A1:A3"],
  ["(A1:A3)", "(@A1:A3)"],
  ["SUM(Sales[Amount])", "SUM(Sales[Amount])"],
  ["Sales[Amount]*2", "@Sales[Amount]*2"],
  ["Sales[@Amount]*2", "Sales[@Amount]*2"],
  ["QTY*2", "@QTY*2"],
  ["SUM(QTY)", "SUM(QTY)"],
  ["@A1:A3+1", "@A1:A3+1"],
  ["SUM(A1:A3", "SUM(A1:A3"],
];
eq("where the loader puts the @", cases.map(([f]) => intersected(f, names)), cases.map(([, want]) => want));

eq("the writer takes the loader's own @s out, and no others", [
  stored("@A1:A3", names), stored("ROMAN(@H34:H37,1)", names), stored("@INDEX(A:A,MATCH(1,B:B,0))", names),
  stored("SUM(@A1:A3)", names), stored("SUM(A1:A3*2)", names), stored("@A1:A3+@B1:B3", names), stored("A1+B1", names),
], [
  { text: "A1:A3", legacy: true }, { text: "ROMAN(H34:H37,1)", legacy: true }, { text: "INDEX(A:A,MATCH(1,B:B,0))", legacy: true },
  { text: "SUM(SINGLE(A1:A3))", legacy: false }, { text: "SUM(A1:A3*2)", legacy: false }, { text: "A1:A3+B1:B3", legacy: true }, { text: "A1+B1", legacy: true },
]);

// A workbook as Excel stores one with legacy formulas: the sheet's XML put in a written file.
const seed = new Workbook(), seeded = new Commands(seed);
await seeded.run("cell.set cell=A1 input=1");
const parts = [...unzip(writeXlsx(seed).bytes)].map(([n, b]): [string, Uint8Array] => [n, n === "xl/worksheets/sheet1.xml" ? utf8Bytes(
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
  + '<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c>'
  + '<c r="B2"><f>A1:A4</f><v>2</v></c><c r="D2" cm="1"><f t="array" ref="D2:D4">A1:A3</f><v>1</v></c><c r="E2"><f>EXP(A:A)</f><v>7.38905609893065</v></c>'
  + '<c r="F2"><f t="shared" ref="F2:F3" si="0">A1:A4</f><v>2</v></c></row>'
  + '<row r="3"><c r="A3"><v>3</v></c><c r="C3" t="str"><f>_xlfn.ROMAN(A1:A4)</f><v>III</v></c><c r="D3"><v>2</v></c><c r="F3"><f t="shared" si="0"/><v>3</v></c></row>'
  + '<row r="4"><c r="A4"><v>4</v></c><c r="D4"><v>3</v></c><c r="G4"><f>SUM(A1:A4*2)</f><v>8</v></c></row></sheetData></worksheet>') : b]);
const read = readXlsx(zip(parts));
const wb = new Workbook(), cmds = new Commands(wb);
const o = await cmds.execute("file.import", { text: JSON.stringify(read.file) });
eq("the file is read", o.ok, true);
const at = (k: string) => wb.sheet.input(k);
eq("the importer puts the @ in a formula with no array flag", [at("B2"), at("C3"), at("E2"), at("F2"), at("F3"), at("G4")],
  ["=@A1:A4", "=ROMAN(@A1:A4)", "=EXP(@A:A)", "=@A1:A4", "=@A2:A5", "=SUM(@A1:A4*2)"]);
eq("an array formula has none, and spills", [at("D2"), wb.sheet.spillOf("D2").length], ["=A1:A3", 2]);
eq("each is one value, as Excel stored it", [wb.sheet.value("B2"), wb.sheet.value("C3"), wb.sheet.value("E2"), wb.sheet.value("F3"), wb.sheet.value("G4")],
  [2, "III", Math.exp(2), 3, 8]);

// Written again: the legacy formulas as they were, with no array flag; one the loader would work out
// otherwise (SUM(A1:A3*2) typed in a dynamic-array Excel, or a written @) as an array formula.
await cmds.run("cell.set cell=H1 input==SUM(@A1:A3)");
await cmds.run("cell.set cell=H2 input==SUM(A1:A3*2)");
await cmds.run("cell.set cell=H3 input==SUM(A1:A3)");
const sheet = utf8Text(unzip(writeXlsx(wb).bytes).get("xl/worksheets/sheet1.xml")!);
const f = (ref: string) => { const m = new RegExp(`<c r="${ref}"[^>]*>\\s*<f((?: [^>]*)?)>([^<]*)</f>`).exec(sheet); return m ? [m[1].trim(), m[2]] : null; };
eq("written back as the legacy formulas they were", [f("B2"), f("C3"), f("E2"), f("F3")],
  [["", "A1:A4"], ["", "ROMAN(A1:A4)"], ["", "EXP(A:A)"], ["", "A2:A5"]]);
eq("an array formula keeps its flag", f("D2"), ['t="array" ref="D2:D4"', "A1:A3"]);
eq("what the loader would work out otherwise is an array formula, a written @ SINGLE", [f("H1"), f("H2"), f("H3")],
  [['t="array" ref="H1"', "SUM(_xlfn.SINGLE(A1:A3))"], ['t="array" ref="H2"', "SUM(A1:A3*2)"], ["", "SUM(A1:A3)"]]);
const back = new Workbook(), cmds2 = new Commands(back);
await cmds2.execute("file.import", { text: JSON.stringify(readXlsx(writeXlsx(wb).bytes).file) });
eq("and is read back as it was", [back.sheet.input("B2"), back.sheet.input("F3"), back.sheet.value("B2"), back.sheet.value("H1"), back.sheet.value("H2"), back.sheet.value("H3")],
  ["=@A1:A4", "=@A2:A5", 2, 1, 12, 6]);
// expect: ok where the loader puts the @
// expect: ok the writer takes the loader's own @s out, and no others
// expect: ok the file is read
// expect: ok the importer puts the @ in a formula with no array flag
// expect: ok an array formula has none, and spills
// expect: ok each is one value, as Excel stored it
// expect: ok written back as the legacy formulas they were
// expect: ok an array formula keeps its flag
// expect: ok what the loader would work out otherwise is an array formula, a written @ SINGLE
// expect: ok and is read back as it was
