// View > Show (6.14a): a sheet's gridlines and headings, shown as it
// starts, hidden, saved with it (the sheet file and XLSX's sheetView), each
// sheet its own; the formula bar the window's; refusals; SETVIEW and
// GETVIEW from REXX.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
eq("shown as it starts", [wb.page("Sheet1")!.gridlines, wb.page("Sheet1")!.headings, wb.formulaBar], [true, true, true]);
eq("hidden", [(await cmds.run("view.show what=gridlines on=0")).message, (await cmds.run("view.show what=headings on=0")).message, (await cmds.run("view.show what=formulabar on=0")).message, wb.page("Sheet1")!.gridlines, wb.formulaBar],
  ["Sheet1's gridlines hidden", "Sheet1's headings hidden", "the formula bar hidden", false, false]);
await cmds.run("sheet.add name=Other");
eq("each sheet its own", [wb.page("Other")!.gridlines, wb.page("Sheet1")!.gridlines], [true, false]);
eq("refused", (await cmds.run("view.show what=rulers")).message, "rulers is not gridlines, headings or formulabar");
const again = new Workbook(), cmds2 = new Commands(again);
await cmds2.execute("file.import", { text: (await cmds.run("file.export")).output as string });
eq("saved with the file", [again.page("Sheet1")!.gridlines, again.page("Sheet1")!.headings, again.page("Other")!.gridlines], [false, false, true]);
const x = await cmds.run("file.exportXlsx");
const back = new Workbook(), cmds3 = new Commands(back);
await cmds3.run(`file.importXlsx name=V.xlsx base64=${x.output}`);
eq("an XLSX file's sheetView", [back.page("Sheet1")!.gridlines, back.page("Sheet1")!.headings, back.page("Other")!.headings], [false, false, true]);
const said: string[] = [];
await cmds.script('SAY SETVIEW("gridlines", 1) GETVIEW("gridlines")\nSAY SETVIEW("formulabar") GETVIEW("formula bar")\nSAY GETVIEW("headings")', (l) => said.push(l));
eq("from REXX", said, ["1 1", "1 1", "1"]);

// expect: ok shown as it starts
// expect: ok hidden
// expect: ok each sheet its own
// expect: ok refused
// expect: ok saved with the file
// expect: ok an XLSX file's sheetView
// expect: ok from REXX
