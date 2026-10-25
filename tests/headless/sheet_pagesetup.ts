// Page setup (docs/comparisons/LIBREOFFICE.md action 6, print/layout): a sheet's
// own paper, orientation, margin, scale, print area, rows to repeat at the top
// of each page, header and footer, kept with the sheet (undoable, saved in the
// file and in XLSX), and used by Print, Export PDF and Page Breaks unless they
// are told otherwise.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { fromBase64 } from "./sheet/csv.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook(), cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
const text = (o: { output?: unknown }) => Array.from(fromBase64(String(o.output)), (c) => String.fromCharCode(c)).join("");
const count = (s: string, what: string) => s.split(what).length - 1;
const layout = async () => (await cmds.execute("page.layout")).output as Record<string, unknown>;

eq("the layout of a new sheet", await layout(), { paper: "a4", landscape: false, margin: 36, scale: 100, fitWide: 0, fitTall: 0, area: "", titles: "", titleCols: "", header: "", footer: "", gridlines: false, headings: false });
await cmds.run("cell.set cell=A1 input=Item");
await cmds.run("cell.set cell=B1 input=Cost");
for (let r = 2; r <= 120; r++) { await cmds.run(`cell.set cell=A${r} input=row${r}`); await cmds.run(`cell.set cell=B${r} input=${r}`); }
const pages = async (extra = "") => Number((await cmds.run(`file.exportPdf ${extra}`)).message.split(" ")[1]);
const portrait = await pages();
eq("120 rows on A4: 3 pages", portrait, 3);

// Setting it.
const set = await cmds.run("page.setup paper=letter landscape=1 margin=54 scale=80 header=&A footer=\"Page &P of &N\"");
eq("page.setup says what it set", [set.ok, set.message], [true, "page setup: letter, landscape, margin 54, scale 80%, header &A, footer Page &P of &N"]);
eq("and the layout holds it", await layout(), { paper: "letter", landscape: true, margin: 54, scale: 80, fitWide: 0, fitTall: 0, area: "", titles: "", titleCols: "", header: "&A", footer: "Page &P of &N", gridlines: false, headings: false });
const setupPages = await pages();
eq("Export PDF follows the sheet's setup: landscape letter at 80% takes the rows on 4 pages", setupPages, 4);
eq("and a parameter wins over the setup", await pages("landscape=0 scale=100 paper=a4 margin=36"), portrait);

// Scale.
await cmds.run("page.setup paper=a4 landscape=0 margin=36 scale=100 header= footer=");
await cmds.run("page.setup scale=50");
eq("scale 50 prints the same rows on fewer pages", await pages(), 2);
await cmds.run("page.setup scale=100");

// Print area.
await cmds.run("page.setup area=A1:B10");
eq("a print area is what prints", await pages(), 1);
const area = text(await cmds.run("file.exportPdf"));
eq("and nothing beyond it", [area.includes("(row10) Tj"), area.includes("(row11) Tj")], [true, false]);
eq("a range given still wins", await pages("range=A1:B120"), 3);
await cmds.run("page.setup area=none");
eq("area=none takes it away", (await layout()).area, "");

// Rows repeated at the top.
await cmds.run("page.setup titles=1:1");
const titled = text(await cmds.run("file.exportPdf"));
eq("a title row is printed on every page: the header cell once a page", [count(titled, "(Item) Tj"), count(titled, "/Type /Page ")], [3, 3]);
eq("(a page of its own is short a row, so the sheet may take more pages)", (await pages()) >= 3, true);
await cmds.run("page.setup titles=none");
eq("titles=none: the header only once", count(text(await cmds.run("file.exportPdf")), "(Item) Tj"), 1);

// Header and footer.
await cmds.run("page.setup header=\"Costs report\" footer=\"&LLeft&CPage &P of &N&RRight\"");
const hf = text(await cmds.run("file.exportPdf"));
eq("a header on each page; a footer in three parts with the page numbers", [count(hf, "(Costs report) Tj"), hf.includes("(Page 2 of 3) Tj"), hf.includes("(Left) Tj"), hf.includes("(Right) Tj")], [3, true, true, true]);
await cmds.run("page.setup header=\"&A &F\" footer=");
eq("&A is the sheet's name, &F the file's", text(await cmds.run("file.exportPdf")).includes("(Sheet1 Untitled) Tj"), true);

// Page Breaks follows the layout.
await cmds.run("page.setup header= footer= paper=a4 landscape=0");
const portraitBreaks = (await cmds.run("page.breaks")).message;
await cmds.run("page.setup landscape=1");
eq("page.breaks uses the sheet's orientation", (await cmds.run("page.breaks")).message === portraitBreaks, false);
await cmds.run("page.setup landscape=0");

// Undo, redo.
await cmds.run("page.setup scale=70");
await cmds.run("edit.undo");
eq("one step of undo, and redo", [(await layout()).scale, (await cmds.run("edit.redo")).ok, (await layout()).scale], [100, true, 70]);

// The file.
await cmds.run("page.setup paper=letter landscape=1 margin=20 scale=90 area=A1:B5 titles=1:1 header=H footer=F gridlines=1 headings=1");
const saved = (await cmds.run("file.export")).output as string;
await cmds.run("file.new");
eq("a new workbook has the default", (await layout()).paper, "a4");
await cmds.execute("file.import", { text: saved });
eq("a file read back keeps the layout", await layout(), { paper: "letter", landscape: true, margin: 20, scale: 90, fitWide: 0, fitTall: 0, area: "A1:B5", titles: "1:1", titleCols: "", header: "H", footer: "F", gridlines: true, headings: true });
// A file with no layout (one saved before layouts, or by a sheet that has the default) reads with the default.
{
  await cmds.run("file.new");
  await cmds.run("cell.set cell=A1 input=x");
  const plain = (await cmds.run("file.export")).output as string;
  await cmds.run("page.setup paper=letter");
  await cmds.execute("file.import", { text: plain });
  eq("a file with no layout reads with the default", [(await layout()).paper, JSON.parse(plain).view?.layout], ["a4", undefined]);
}

// XLSX.
await cmds.run("page.setup paper=letter landscape=1 margin=20 scale=90 area=A1:B5 titles=1:1 header=H footer=F gridlines=1 headings=1");
const xlsx = await cmds.run("file.exportXlsx");
await cmds.run("file.new");
await cmds.execute("file.importXlsx", { base64: String(xlsx.output) });
eq("XLSX keeps the layout: paper, orientation, margin, scale, area, titles, header, footer, gridlines, headings", await layout(),
  { paper: "letter", landscape: true, margin: 20, scale: 90, fitWide: 0, fitTall: 0, area: "A1:B5", titles: "1:1", titleCols: "", header: "H", footer: "F", gridlines: true, headings: true });

// Refusals, and REXX.
eq("refused", [(await cmds.run("page.setup paper=a3")).message, (await cmds.run("page.setup scale=5")).message, (await cmds.run("page.setup scale=500")).message,
  (await cmds.run("page.setup area=nonsense")).message, (await cmds.run("page.setup titles=0:1")).message, (await cmds.run("page.setup margin=200")).message],
  ["a3 is not a paper size (a4, letter)", "5 is not a scale (10 to 400 percent)", "500 is not a scale (10 to 400 percent)", "nonsense is not a range", "0:1 is not rows to repeat (1:2, or none)", "200 is not a margin (0 to 144 points)"]);
const said: string[] = [];
await cmds.script('LET a = PAGESETUP("paper=letter;scale=60;landscape=1")\nSAY a\nSAY GETPAGESETUP()', (l) => said.push(l));
eq("PAGESETUP and GETPAGESETUP from REXX", said, ["page setup: letter, landscape, margin 20, scale 60%, area A1:B5, titles 1:1, header H, footer F, gridlines, headings",
  "page setup: letter, landscape, margin 20, scale 60%, area A1:B5, titles 1:1, header H, footer F, gridlines, headings"]);

// Columns to repeat at the left, and fit to pages wide by tall (LibreOffice 6, left from the first pass).
// A wide sheet: columns A to T of 60 points each, 30 rows. A4 portrait fits 8 columns across and 41 rows down.
const wide = new Workbook(), wcmds = new Commands(wide);
await wcmds.run("formulas.use language=Excel");
const wtext = (o: { output?: unknown }) => Array.from(fromBase64(String(o.output)), (c) => String.fromCharCode(c)).join("");
const wlayout = async () => (await wcmds.execute("page.layout")).output as Record<string, unknown>;
const cols = "ABCDEFGHIJKLMNOPQRST".split("");
for (let r = 1; r <= 30; r++) for (const c of cols) await wcmds.run(`cell.set cell=${c}${r} input=${c === "A" ? `lab${r}` : c === "B" ? `key${r}` : `${c}${r}`}`);
eq("a wide sheet prints 3 pages across", (await wcmds.run("page.breaks")).output, ["A1:H30", "I1:P30", "Q1:T30"]);

// Columns to repeat.
const setCols = await wcmds.run("page.setup titleCols=A:B");
eq("page.setup titleCols=A:B says so", [setCols.ok, setCols.message], [true, "page setup: a4, portrait, margin 36, scale 100%, title columns A:B"]);
eq("the layout holds them", (await wlayout()).titleCols, "A:B");
eq("the pages after the first have less room for their own columns (403 points after the 120 repeated: 6, not 8)", (await wcmds.run("page.breaks")).output, ["A1:H30", "I1:N30", "O1:T30"]);
const repeated = wtext(await wcmds.run("file.exportPdf"));
eq("the key column is printed on every page; the label column too", [count(repeated, "(lab5) Tj"), count(repeated, "(key5) Tj"), count(repeated, "/Type /Page ")], [3, 3, 3]);
eq("and a column that is not repeated is printed once", count(repeated, "(H5) Tj"), 1);
await wcmds.run("page.setup titles=1:2");
const both = wtext(await wcmds.run("file.exportPdf"));
eq("rows and columns repeated together: the corner (lab1 is A1) on every page, only once a page", [count(both, "(lab1) Tj"), count(both, "(lab2) Tj")], [3, 3]);
await wcmds.run("page.setup titles=none");
await wcmds.run("page.setup titleCols=none");
eq("titleCols=none takes them away", [(await wlayout()).titleCols, count(wtext(await wcmds.run("file.exportPdf")), "(lab5) Tj")], ["", 1]);
await wcmds.run("page.setup titleCols=A:B");
await wcmds.run("edit.undo");
eq("one step of undo, and redo", [(await wlayout()).titleCols, (await wcmds.run("edit.redo")).ok, (await wlayout()).titleCols], ["", true, "A:B"]);
await wcmds.run("columns.hide range=A:A");
eq("a hidden title column is not printed", count(wtext(await wcmds.run("file.exportPdf")), "(lab5) Tj"), 0);
await wcmds.run("columns.unhide range=A:A");
await wcmds.run("page.setup titleCols=none");

// Fit to pages wide by tall.
const fit1 = await wcmds.run("page.setup fitWide=1");
eq("fitWide=1 says so", [fit1.ok, fit1.message], [true, "page setup: a4, portrait, margin 36, scale 100%, fit 1 wide by automatic tall"]);
eq("one page across: the sheet is printed on one page", (await wcmds.run("page.breaks")).output, ["A1:T30"]);
eq("at the largest scale that fits (1,200 points of columns into 523: 43%), said", (await wcmds.run("page.breaks")).message.endsWith("; scale 43%"), true);
eq("Export PDF follows: one page", (await wcmds.run("file.exportPdf")).message.split(" ")[1], "1");
await wcmds.run("page.setup fitWide=2");
eq("two across: the scale where twenty columns take two pages (87%)", [(await wcmds.run("page.breaks")).output, (await wcmds.run("page.breaks")).message.endsWith("; scale 87%")], [["A1:J30", "K1:T30"], true]);
await wcmds.run("page.break cell=F1");
eq("a manual break is ignored while the sheet is fitted, as Excel does", (await wcmds.run("page.breaks")).output, ["A1:J30", "K1:T30"]);
await wcmds.run("page.resetBreaks");
await wcmds.run("page.setup fitWide=0");
eq("fitWide=0 and fitTall=0: not fitted, the scale applies", [(await wlayout()).fitWide, (await wlayout()).fitTall, (await wcmds.run("page.breaks")).output], [0, 0, ["A1:H30", "I1:P30", "Q1:T30"]]);
await wcmds.run("page.setup fitWide=1 fitTall=1");
eq("1 by 1 prints the whole sheet on one page", [(await wcmds.run("file.exportPdf")).message.split(" ")[1], (await wlayout()).fitTall], ["1", 1]);
await wcmds.run("page.setup fitWide=1 fitTall=0");
await wcmds.run("page.setup scale=80");
eq("a scale given on its own takes the fit away (Adjust to ...% of normal size)", [(await wlayout()).fitWide, (await wlayout()).fitTall, (await wlayout()).scale], [0, 0, 80]);
await wcmds.run("page.setup fit=none scale=100");
// A sheet that already fits is not enlarged.
{
  const small = new Workbook(), scmds = new Commands(small);
  await scmds.run("cell.set cell=A1 input=x");
  await scmds.run("page.setup fitWide=1 fitTall=1");
  eq("a sheet that is small enough stays at its size (fit never enlarges)", [(await scmds.run("page.breaks")).output, (await scmds.run("page.breaks")).message.includes("scale")], [["A1:A1"], true]);
  eq("(and says it printed at 100%)", (await scmds.run("page.breaks")).message.endsWith("; scale 100%"), true);
}
// Tall: 120 rows.
{
  const tall = new Workbook(), tcmds = new Commands(tall);
  for (let r = 1; r <= 120; r++) { await tcmds.run(`cell.set cell=A${r} input=row${r}`); await tcmds.run(`cell.set cell=B${r} input=${r}`); }
  await tcmds.run("page.setup fitTall=1 fitWide=0");
  eq("fitTall=1: 120 rows on one page (the largest scale: 34%)", [(await tcmds.run("page.breaks")).output, (await tcmds.run("page.breaks")).message.endsWith("; scale 34%")], [["A1:B120"], true]);
}

// The fit and the columns undone, redone, saved, in XLSX.
await wcmds.run("page.setup fitWide=2 fitTall=0 titleCols=A:B titles=1:1");
await wcmds.run("page.setup fitTall=3");
await wcmds.run("edit.undo");
eq("a fit is a step of undo", [(await wlayout()).fitTall, (await wcmds.run("edit.redo")).ok, (await wlayout()).fitTall], [0, true, 3]);
const wsaved = (await wcmds.run("file.export")).output as string;
await wcmds.run("file.new");
await wcmds.execute("file.import", { text: wsaved });
eq("a file read back keeps the fit and the columns", [(await wlayout()).fitWide, (await wlayout()).fitTall, (await wlayout()).titleCols, (await wlayout()).titles], [2, 3, "A:B", "1:1"]);
const wx = await wcmds.run("file.exportXlsx");
const wxml = wtext(await wcmds.run("file.exportXlsx")).length > 0;
await wcmds.run("file.new");
const imp = await wcmds.execute("file.importXlsx", { base64: String(wx.output) });
eq("XLSX keeps the fit (fitToPage, fitToWidth, fitToHeight) and the columns to repeat (Print_Titles: A:B and 1:1)", [(await wlayout()).fitWide, (await wlayout()).fitTall, (await wlayout()).titleCols, (await wlayout()).titles], [2, 3, "A:B", "1:1"]);
eq("and the import report no longer says columns to repeat were left out", String(imp.message).includes("columns to repeat"), false);
await wcmds.run("page.setup fitTall=0");
const wx2 = await wcmds.run("file.exportXlsx");
await wcmds.run("file.new");
await wcmds.execute("file.importXlsx", { base64: String(wx2.output) });
eq("a fit with an automatic height reads back as 0 (Excel's fitToHeight=0)", [(await wlayout()).fitWide, (await wlayout()).fitTall], [2, 0]);

// Print follows the same setup as Export PDF.
{
  const sent: string[] = [];
  wide.printer = (b64) => { sent.push(wtext({ output: b64 })); return ""; };
  await wcmds.run("page.setup fit=none titles=none titleCols=none");
  await wcmds.run("page.setup fitWide=1 fitTall=1");
  const one = await wcmds.run("file.print");
  eq("Print follows the fit: the whole sheet on one page", [one.message, count(sent[0], "/Type /Page ")], ["1 page to the printer", 1]);
  await wcmds.run("page.setup fit=none titleCols=A:B");
  const three = await wcmds.run("file.print");
  eq("and the columns to repeat: the key column on every page of three", [three.message, count(sent[1], "(key5) Tj")], ["3 pages to the printer", 3]);
  await wcmds.run("page.setup titleCols=none");
}

// Refusals and REXX.
eq("refused", [(await wcmds.run("page.setup fitWide=-1")).message, (await wcmds.run("page.setup fitTall=abc")).message, (await wcmds.run("page.setup fitWide=100")).message, (await wcmds.run("page.setup titleCols=1:2")).message, (await wcmds.run("page.setup titleCols=A")).message],
  ["-1 is not a number of pages (0 for as many as it needs, to 99)", "abc is not a number of pages (0 for as many as it needs, to 99)", "100 is not a number of pages (0 for as many as it needs, to 99)", "1:2 is not columns to repeat (A:B, or none)", "A is not columns to repeat (A:B, or none)"]);
{
  const said2: string[] = [];
  await wcmds.script('LET a = PAGESETUP("fitWide=1;fitTall=2;titleCols=B:C")\nSAY a', (l) => said2.push(l));
  eq("PAGESETUP takes fitWide, fitTall and titleCols", said2, ["page setup: a4, portrait, margin 36, scale 100%, fit 1 wide by 2 tall, title columns B:C"]);
}

// expect: ok the layout of a new sheet
// expect: ok 120 rows on A4: 3 pages
// expect: ok page.setup says what it set
// expect: ok and the layout holds it
// expect: ok Export PDF follows the sheet's setup: landscape letter at 80% takes the rows on 4 pages
// expect: ok and a parameter wins over the setup
// expect: ok scale 50 prints the same rows on fewer pages
// expect: ok a print area is what prints
// expect: ok and nothing beyond it
// expect: ok a range given still wins
// expect: ok area=none takes it away
// expect: ok a title row is printed on every page: the header cell once a page
// expect: ok (a page of its own is short a row, so the sheet may take more pages)
// expect: ok titles=none: the header only once
// expect: ok a header on each page; a footer in three parts with the page numbers
// expect: ok &A is the sheet's name, &F the file's
// expect: ok page.breaks uses the sheet's orientation
// expect: ok one step of undo, and redo
// expect: ok a new workbook has the default
// expect: ok a file read back keeps the layout
// expect: ok a file with no layout reads with the default
// expect: ok XLSX keeps the layout: paper, orientation, margin, scale, area, titles, header, footer, gridlines, headings
// expect: ok refused
// expect: ok PAGESETUP and GETPAGESETUP from REXX
// expect: ok a wide sheet prints 3 pages across
// expect: ok page.setup titleCols=A:B says so
// expect: ok the layout holds them
// expect: ok the pages after the first have less room for their own columns (403 points after the 120 repeated: 6, not 8)
// expect: ok the key column is printed on every page; the label column too
// expect: ok and a column that is not repeated is printed once
// expect: ok rows and columns repeated together: the corner (lab1 is A1) on every page, only once a page
// expect: ok titleCols=none takes them away
// expect: ok one step of undo, and redo
// expect: ok a hidden title column is not printed
// expect: ok fitWide=1 says so
// expect: ok one page across: the sheet is printed on one page
// expect: ok at the largest scale that fits (1,200 points of columns into 523: 43%), said
// expect: ok Export PDF follows: one page
// expect: ok two across: the scale where twenty columns take two pages (87%)
// expect: ok a manual break is ignored while the sheet is fitted, as Excel does
// expect: ok fitWide=0 and fitTall=0: not fitted, the scale applies
// expect: ok 1 by 1 prints the whole sheet on one page
// expect: ok a scale given on its own takes the fit away (Adjust to ...% of normal size)
// expect: ok a sheet that is small enough stays at its size (fit never enlarges)
// expect: ok (and says it printed at 100%)
// expect: ok fitTall=1: 120 rows on one page (the largest scale: 34%)
// expect: ok a fit is a step of undo
// expect: ok a file read back keeps the fit and the columns
// expect: ok XLSX keeps the fit (fitToPage, fitToWidth, fitToHeight) and the columns to repeat (Print_Titles: A:B and 1:1)
// expect: ok and the import report no longer says columns to repeat were left out
// expect: ok a fit with an automatic height reads back as 0 (Excel's fitToHeight=0)
// expect: ok Print follows the fit: the whole sheet on one page
// expect: ok and the columns to repeat: the key column on every page of three
// expect: ok refused
// expect: ok PAGESETUP takes fitWide, fitTall and titleCols
