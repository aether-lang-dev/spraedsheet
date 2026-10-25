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

eq("the layout of a new sheet", await layout(), { paper: "a4", landscape: false, margin: 36, scale: 100, area: "", titles: "", header: "", footer: "", gridlines: false, headings: false });
await cmds.run("cell.set cell=A1 input=Item");
await cmds.run("cell.set cell=B1 input=Cost");
for (let r = 2; r <= 120; r++) { await cmds.run(`cell.set cell=A${r} input=row${r}`); await cmds.run(`cell.set cell=B${r} input=${r}`); }
const pages = async (extra = "") => Number((await cmds.run(`file.exportPdf ${extra}`)).message.split(" ")[1]);
const portrait = await pages();
eq("120 rows on A4: 3 pages", portrait, 3);

// Setting it.
const set = await cmds.run("page.setup paper=letter landscape=1 margin=54 scale=80 header=&A footer=\"Page &P of &N\"");
eq("page.setup says what it set", [set.ok, set.message], [true, "page setup: letter, landscape, margin 54, scale 80%, header &A, footer Page &P of &N"]);
eq("and the layout holds it", await layout(), { paper: "letter", landscape: true, margin: 54, scale: 80, area: "", titles: "", header: "&A", footer: "Page &P of &N", gridlines: false, headings: false });
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
eq("a file read back keeps the layout", await layout(), { paper: "letter", landscape: true, margin: 20, scale: 90, area: "A1:B5", titles: "1:1", header: "H", footer: "F", gridlines: true, headings: true });
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
  { paper: "letter", landscape: true, margin: 20, scale: 90, area: "A1:B5", titles: "1:1", header: "H", footer: "F", gridlines: true, headings: true });

// Refusals, and REXX.
eq("refused", [(await cmds.run("page.setup paper=a3")).message, (await cmds.run("page.setup scale=5")).message, (await cmds.run("page.setup scale=500")).message,
  (await cmds.run("page.setup area=nonsense")).message, (await cmds.run("page.setup titles=0:1")).message, (await cmds.run("page.setup margin=200")).message],
  ["a3 is not a paper size (a4, letter)", "5 is not a scale (10 to 400 percent)", "500 is not a scale (10 to 400 percent)", "nonsense is not a range", "0:1 is not rows to repeat (1:2, or none)", "200 is not a margin (0 to 144 points)"]);
const said: string[] = [];
await cmds.script('LET a = PAGESETUP("paper=letter;scale=60;landscape=1")\nSAY a\nSAY GETPAGESETUP()', (l) => said.push(l));
eq("PAGESETUP and GETPAGESETUP from REXX", said, ["page setup: letter, landscape, margin 20, scale 60%, area A1:B5, titles 1:1, header H, footer F, gridlines, headings",
  "page setup: letter, landscape, margin 20, scale 60%, area A1:B5, titles 1:1, header H, footer F, gridlines, headings"]);

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
