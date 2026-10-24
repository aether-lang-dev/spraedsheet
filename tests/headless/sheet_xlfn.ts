// Excel stores a function newer than 2007 under a prefix (_xlfn.TEXTJOIN,
// _xlfn._xlws.SORT) and a spill's reference as ANCHORARRAY(A1). A file Excel
// saved is read with those taken off, and a file written has them put back, so
// Excel recognises the function (without the prefix it shows #NAME?). Found by
// the IronCalc calc_tests workbooks (docs/comparisons/IRONCALC.md action 2).
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { readXlsx, plainFormula } from "./sheet/xlsx.ts";
import { writeXlsx } from "./sheet/xlsxwrite.ts";
import { unzip, utf8Text } from "./sheet/zip.ts";
import { eq } from "./lib/check.ts";

eq("the prefixes come off a stored formula", [
  plainFormula('_xlfn.TEXTJOIN(",",TRUE,A1:A3)'), plainFormula("_xlfn._xlws.SORT(A1:A9)"), plainFormula("LET(_xlpm.x,2,_xlpm.x*3)"),
  plainFormula("SUM(_xlfn.ANCHORARRAY(A1))"), plainFormula('"_xlfn.X in a string"&_xlfn.UNICODE("a")'), plainFormula("A_xlfn.B1"),
], ['TEXTJOIN(",",TRUE,A1:A3)', "SORT(A1:A9)", "LET(x,2,x*3)", "SUM(A1#)", '"_xlfn.X in a string"&UNICODE("a")', "A_xlfn.B1"]);

const wb = new Workbook(), cmds = new Commands(wb);
for (const l of ["formulas.use language=Excel", "cell.set cell=A1 input=Tea", "cell.set cell=A2 input=Cake", "cell.set cell=A3 input=Soup",
  'cell.set cell=B1 input="=TEXTJOIN(""-"",TRUE,A1:A3)"', "cell.set cell=B2 input==SORT(A1:A3)", "cell.set cell=C1 input==SUM(LEN(B2#))",
  "cell.set cell=D1 input==STDEV.S(1,2,3)", "cell.set cell=D2 input==XLOOKUP(\"Cake\",A1:A3,A1:A3)", "cell.set cell=D3 input==SUM(1,2)",
  'cell.set cell=D4 input="=""_xlfn.TEXTJOIN(""&LEN(A1)"']) {
  const o = await cmds.run(l);
  if (!o.ok) print(`FAIL setup ${l}: ${o.message}`);
}
const bytes = writeXlsx(wb).bytes;
const sheet = utf8Text(unzip(bytes).get("xl/worksheets/sheet1.xml")!);
const f = (ref: string) => { const m = new RegExp(`<c r="${ref}"[^>]*>\\s*<f(?: [^>]*)?>([^<]*)</f>`).exec(sheet); return m ? m[1].replace(/&quot;/g, '"').replace(/&amp;/g, "&") : null; };
eq("written with the prefixes Excel needs", [f("B1"), f("B2"), f("C1"), f("D1"), f("D2"), f("D3")],
  ['_xlfn.TEXTJOIN("-",TRUE,A1:A3)', "_xlfn._xlws.SORT(A1:A3)", "SUM(LEN(_xlfn.ANCHORARRAY(B2)))", "_xlfn.STDEV.S(1,2,3)", '_xlfn.XLOOKUP("Cake",A1:A3,A1:A3)', "SUM(1,2)"]);
eq("text in a string is left alone", f("D4"), '"_xlfn.TEXTJOIN("&LEN(A1)');

// And read back: the formulas as they were typed, with their values.
const back = new Workbook(), cmds2 = new Commands(back);
const o = await cmds2.execute("file.import", { text: JSON.stringify(readXlsx(bytes).file) });
eq("read back as typed", [o.ok, back.sheet.input("B1"), back.sheet.input("B2"), back.sheet.input("C1"), back.sheet.input("D1"), back.sheet.input("D2"), back.sheet.value("B1"), back.sheet.value("D2")],
  [true, wb.sheet.input("B1"), "=SORT(A1:A3)", "=SUM(LEN(B2#))", "=STDEV.S(1,2,3)", wb.sheet.input("D2"), "Tea-Cake-Soup", "Cake"]);

// expect: ok the prefixes come off a stored formula
// expect: ok written with the prefixes Excel needs
// expect: ok text in a string is left alone
// expect: ok read back as typed
