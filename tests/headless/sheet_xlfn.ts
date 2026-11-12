// Excel stores a function newer than 2007 under a prefix (_xlfn.TEXTJOIN,
// _xlfn._xlws.SORT) and a spill's reference as ANCHORARRAY(A1). A file Excel
// saved is read with those taken off, and a file written has them put back, so
// Excel recognises the function (without the prefix it shows #NAME?). Found by
// the IronCalc calc_tests workbooks (docs/comparisons/IRONCALC.md action 2).
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { readXlsx, plainFormula, storedFormula } from "./sheet/xlsx.ts";
import { writeXlsx } from "./sheet/xlsxwrite.ts";
import { unzip, utf8Text } from "./sheet/zip.ts";
import { eq } from "./lib/check.ts";

eq("the prefixes come off a stored formula", [
  plainFormula('_xlfn.TEXTJOIN(",",TRUE,A1:A3)'), plainFormula("_xlfn._xlws.SORT(A1:A9)"), plainFormula("LET(_xlpm.x,2,_xlpm.x*3)"),
  plainFormula("SUM(_xlfn.ANCHORARRAY(A1))"), plainFormula('"_xlfn.X in a string"&_xlfn.UNICODE("a")'), plainFormula("A_xlfn.B1"),
], ['TEXTJOIN(",",TRUE,A1:A3)', "SORT(A1:A9)", "LET(x,2,x*3)", "SUM(A1#)", '"_xlfn.X in a string"&UNICODE("a")', "A_xlfn.B1"]);

eq("a LAMBDA's optional parameter, _xlop.b, is [b]", plainFormula("_xlfn.LAMBDA(_xlpm.a,_xlop.b,IF(_xlfn.ISOMITTED(_xlpm.b),_xlpm.a/2,_xlpm.a*_xlpm.b))(A5)"),
  "LAMBDA(a,[b],IF(ISOMITTED(b),a/2,a*b))(A5)");

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

// LET and LAMBDA names go to the file as _xlpm.name (an optional parameter's declaration as _xlop.name),
// as Excel's own files (IronCalc's LET.xlsx, LAMBDA.xlsx, ISOMITTED.xlsx) have them, and come back as typed.
{
  const w = new Workbook(), wc = new Commands(w);
  for (const l of ["formulas.use language=Excel", "cell.set cell=A1 input=7",
    "cell.set cell=B1 input==LET(x,A1,y,2,x*y)", "cell.set cell=B2 input==LAMBDA(a,[b],IF(ISOMITTED(b),a/2,a*b))(A1)",
    "cell.set cell=B3 input==x+LET(x,1,x)", 'cell.set cell=B4 input="=LET(s,""x y"",LEN(s)&""s"")"',
    "cell.set cell=B5 input==LET(area,LAMBDA(x,y,x*y),area(A1,2))", "cell.set cell=B6 input==MAP(A1:A1,LAMBDA(q,q+1))"]) {
    const o = await wc.run(l);
    if (!o.ok) print(`FAIL setup ${l}: ${o.message}`);
  }
  const xb = writeXlsx(w).bytes, xs = utf8Text(unzip(xb).get("xl/worksheets/sheet1.xml")!);
  const g = (ref: string) => { const m = new RegExp(`<c r="${ref}"[^>]*>\\s*<f(?: [^>]*)?>([^<]*)</f>`).exec(xs); return m ? m[1].replace(/&quot;/g, '"').replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">") : null; };
  eq("LET and LAMBDA written with _xlpm. and _xlop.", [g("B1"), g("B2"), g("B3"), g("B4"), g("B5"), g("B6")], [
    "_xlfn.LET(_xlpm.x,A1,_xlpm.y,2,_xlpm.x*_xlpm.y)",
    "_xlfn.LAMBDA(_xlpm.a,_xlop.b,IF(_xlfn.ISOMITTED(_xlpm.b),_xlpm.a/2,_xlpm.a*_xlpm.b))(A1)",
    "x+_xlfn.LET(_xlpm.x,1,_xlpm.x)",
    '_xlfn.LET(_xlpm.s,"x y",LEN(_xlpm.s)&"s")',
    "_xlfn.LET(_xlpm.area,_xlfn.LAMBDA(_xlpm.x,_xlpm.y,_xlpm.x*_xlpm.y),_xlpm.area(A1,2))",
    "_xlfn.MAP(A1:A1,_xlfn.LAMBDA(_xlpm.q,_xlpm.q+1))",
  ]);
  const b2 = new Workbook(), c2 = new Commands(b2);
  await c2.execute("file.import", { text: JSON.stringify(readXlsx(xb).file) });
  eq("LET and LAMBDA read back as typed", ["B1", "B2", "B3", "B4", "B5", "B6"].map((r) => b2.sheet.input(r)), ["B1", "B2", "B3", "B4", "B5", "B6"].map((r) => w.sheet.input(r)));
  eq("and to the same values", ["B1", "B2", "B4", "B5", "B6"].map((r) => b2.sheet.value(r)), ["B1", "B2", "B4", "B5", "B6"].map((r) => w.sheet.value(r)));
  // A named LAMBDA is a defined name's formula.
  const n = new Workbook(), nc = new Commands(n);
  for (const l of ["formulas.use language=Excel", 'names.define name=Twice formula="=LAMBDA(t,t*2)"', "cell.set cell=A1 input==Twice(4)"]) await nc.run(l);
  const nb = writeXlsx(n).bytes, wbx = utf8Text(unzip(nb).get("xl/workbook.xml")!);
  eq("a named LAMBDA written with its prefixes", /<definedName name="TWICE">_xlfn\.LAMBDA\(_xlpm\.t,_xlpm\.t\*2\)<\/definedName>/.test(wbx), true);
}
// The names the spec's future-function list has, and the two it leaves bare.
eq("spec names get the prefix", [storedFormula("IMCOSH(1)"), storedFormula("NETWORKDAYS.INTL(1,2)"), storedFormula("FILTERXML(1,2)")], ["_xlfn.IMCOSH(1)", "NETWORKDAYS.INTL(1,2)", "_xlfn.FILTERXML(1,2)"]);

// A cell of empty text (typed ') is a shared string of nothing in a file, and stays text when
// read: VALUE of it is #VALUE!, where a blank cell's is 0.
{
  const e = new Workbook(), ec = new Commands(e);
  for (const l of ["formulas.use language=Excel", "cell.set cell=A1 input='", "cell.set cell=B1 input==VALUE(A1)", "cell.set cell=B2 input==VALUE(A2)", "cell.set cell=B3 input==ISTEXT(A1)"]) await ec.run(l);
  const bytes2 = writeXlsx(e).bytes, xml = utf8Text(unzip(bytes2).get("xl/worksheets/sheet1.xml")!);
  const again = new Workbook(), ac = new Commands(again);
  await ac.execute("file.import", { text: JSON.stringify(readXlsx(bytes2).file) });
  eq("empty text survives a file", [/<c r="A1" t="s"><v>\d+<\/v><\/c>/.test(xml), again.sheet.input("A1"), again.sheet.value("B1"), again.sheet.value("B2"), again.sheet.value("B3")].map((v) => (typeof v === "object" ? (v as { code: string }).code : v)),
    [true, "'", "#VALUE!", 0, true]);
}

// expect: ok the prefixes come off a stored formula
// expect: ok a LAMBDA's optional parameter, _xlop.b, is [b]
// expect: ok written with the prefixes Excel needs
// expect: ok text in a string is left alone
// expect: ok read back as typed
// expect: ok LET and LAMBDA written with _xlpm. and _xlop.
// expect: ok LET and LAMBDA read back as typed
// expect: ok and to the same values
// expect: ok a named LAMBDA written with its prefixes
// expect: ok spec names get the prefix
// expect: ok empty text survives a file
