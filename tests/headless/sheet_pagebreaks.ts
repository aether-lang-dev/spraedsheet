// Page breaks and Page Break Preview (6.14b): Insert Page Break at a cell
// (a page begins at its row and its column; in row 1 or column A the one
// alone), removed, reset; the pages the sheet prints as, the manual breaks
// among the automatic; the PDF paginated by them; one step of undo; saved
// with the sheet and in XLSX (rowBreaks, colBreaks); the preview a view;
// and the REXX functions.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
for (let r = 1; r <= 60; r++) await cmds.run(`cell.set cell=A${r} input=${r}`);
await cmds.run("cell.set cell=H1 input=wide");
const before = (await cmds.run("page.breaks")).output as string[];
eq("the pages a sheet prints as, A4 portrait (41 rows of 18 pt, 8 columns of 60)", before, ["A1:H41", "A42:H60"]);
eq("Insert Page Break at a cell", [(await cmds.run("page.break cell=D20")).message, wb.page("Sheet1")!.breaks], ["page breaks: row 20, column D", { rows: [19], cols: [3] }]);
eq("the pages, the manual breaks among them", (await cmds.run("page.breaks")).output, ["A1:C19", "A20:C60", "D1:H19", "D20:H60"]);
eq("in row 1 or column A, the one alone", [(await cmds.run("page.break cell=A40")).message, (await cmds.run("page.break cell=F1")).message], ["page breaks: row 20, row 40, column D", "page breaks: row 20, row 40, column D, column F"]);
const pdf = await cmds.run("file.exportPdf");
eq("the PDF paginated by them", pdf.message.split(",")[0], "Untitled.pdf: 9 pages");
eq("removed, and undone", [(await cmds.run("page.break cell=F1 off=1")).message, (await cmds.run("edit.undo")).ok, wb.page("Sheet1")!.breaks.cols], ["page breaks: row 20, row 40, column D", true, [3, 5]]);
const again = new Workbook(), cmds2 = new Commands(again);
await cmds2.execute("file.import", { text: (await cmds.run("file.export")).output as string });
eq("saved with the sheet", again.page("Sheet1")!.breaks, { rows: [19, 39], cols: [3, 5] });
const x = await cmds.run("file.exportXlsx");
const back = new Workbook(), cmds3 = new Commands(back);
await cmds3.run(`file.importXlsx name=B.xlsx base64=${x.output}`);
eq("and in XLSX", back.page("Sheet1")!.breaks, { rows: [19, 39], cols: [3, 5] });
eq("Reset All Page Breaks", [(await cmds.run("page.resetBreaks")).message, wb.page("Sheet1")!.breaks], ["page breaks: none", { rows: [], cols: [] }]);
eq("Page Break Preview, a view", [(await cmds.run("view.pageBreakPreview")).message, wb.breakPreview, (await cmds.run("view.pageBreakPreview off=1")).message], ["Page Break Preview", true, "Normal view"]);
const said: string[] = [];
await cmds.script('SAY INSERTPAGEBREAK("B30")\nl = LISTPAGEBREAKS()\nSAY l.0 l.1\nSAY REMOVEPAGEBREAK("B30")\nSAY SETVIEW("pagebreaks", 1)\nSAY RESETPAGEBREAKS()', (l) => said.push(l));
eq("from REXX", said, ["page breaks: row 30, column B", "4 A1:A29", "page breaks: none", "1", "page breaks: none"]);

// expect: ok the pages a sheet prints as, A4 portrait (41 rows of 18 pt, 8 columns of 60)
// expect: ok Insert Page Break at a cell
// expect: ok the pages, the manual breaks among them
// expect: ok in row 1 or column A, the one alone
// expect: ok the PDF paginated by them
// expect: ok removed, and undone
// expect: ok saved with the sheet
// expect: ok and in XLSX
// expect: ok Reset All Page Breaks
// expect: ok Page Break Preview, a view
// expect: ok from REXX
