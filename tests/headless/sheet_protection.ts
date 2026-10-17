// Protection (6.13): Excel's legacy password hash; locked cells refusing a
// change on a protected sheet (typing, clearing, a paste, a fill) and
// unlocked ones taking it; what the protection allows (formatting, rows)
// and does not; Allow Edit Ranges, open and by password; the workbook's
// structure; Unprotect by password; undo; the sheet file and XLSX round
// trips; and the REXX functions.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { passwordHash, opens } from "./sheet/protection.ts";
import { eq } from "./lib/check.ts";

eq("Excel's legacy hash", [passwordHash(""), passwordHash("test").length, passwordHash("test") === passwordHash("test"), passwordHash("test") === passwordHash("Test"), opens(passwordHash("abc"), "abc"), opens(passwordHash("abc"), "abd"), opens(undefined, "x")],
  ["", 4, true, false, true, false, true]);

const wb = new Workbook();
const cmds = new Commands(wb);
const msg = async (line: string) => (await cmds.run(line)).message;
await cmds.run("cell.set cell=A1 input=10");
await cmds.run("cell.set cell=B1 input=20");
await cmds.run("protect.lock range=B1:B3 locked=0");
const pr = await cmds.run("protect.sheet password=secret allow=formatCells");
eq("Protect Sheet", [pr.message, wb.page("Sheet1")!.protection?.allow, (await cmds.run("protect.state cell=A1")).output], ["Sheet1 protected with a password, allowing formatCells", ["formatCells"],
  { sheet: 1, password: 1, allow: ["formatCells"], ranges: [], workbook: 0, locked: 1 }]);
const PROTECTED = "The cell or chart you're trying to change is on a protected sheet. To make a change, unprotect the sheet (Review > Unprotect Sheet). You might be asked for a password.";
eq("a locked cell refuses a change", [await msg("cell.set cell=A1 input=11"), show(wb.sheet.value("A1")), await msg("range.clear range=A1:B1"), show(wb.sheet.value("B1"))], [PROTECTED, "10", PROTECTED, "20"]);
eq("an unlocked one takes it", [(await cmds.run("cell.set cell=B2 input=5")).ok, show(wb.sheet.value("B2"))], [true, "5"]);
await cmds.run("edit.copy range=A1");
eq("a paste or a fill over a locked cell is taken back", [await msg("edit.paste at=C1"), wb.sheet.input("C1"), await msg("edit.fillDown range=A1:A3"), wb.sheet.input("A2")], [PROTECTED, "", PROTECTED, ""]);
eq("a paste into unlocked cells is not", [(await cmds.run("edit.paste at=B3")).ok, show(wb.sheet.value("B3"))], [true, "10"]);
eq("formatting, as allowed", [(await cmds.run("format.font range=A1 bold=1")).ok, wb.notes.get("A1").look?.bold], [true, true]);
eq("what it does not allow", [await msg("rows.insert at=2"), await msg("cells.merge range=A5:B5"), await msg("chart.insert kind=bar range=A1:B2"), await msg("protect.lock range=A1 locked=0")], [PROTECTED, PROTECTED, PROTECTED, PROTECTED]);
eq("the view and the names are not guarded", [(await cmds.run("view.zoom percent=150")).ok, (await cmds.run("names.define name=TOP range=A1")).ok], [true, true]);
eq("Unprotect Sheet needs its password", [await msg("protect.sheet off=1 password=wrong"), await msg("protect.sheet off=1 password=secret"), (await cmds.run("cell.set cell=A1 input=12")).ok], ["The password you supplied is not correct.", "Sheet1 unprotected", true]);
// Allow Edit Ranges: one open, one by its password.
await cmds.run("protect.allowRange title=Open range=D1:D3");
await cmds.run("protect.allowRange title=Keyed range=E1:E3 password=k");
await cmds.run("protect.sheet");
eq("an edit range's cells take a change", [(await cmds.run("cell.set cell=D2 input=7")).ok, await msg("cell.set cell=E2 input=8")], [true, PROTECTED]);
eq("one with a password, once it is given", [await msg("protect.openRange title=Keyed password=nope"), await msg("protect.openRange title=Keyed password=k"), (await cmds.run("cell.set cell=E2 input=8")).ok],
  ["The password you supplied is not correct.", "Keyed may be edited", true]);
eq("undo of an allowed change", [(await cmds.run("edit.undo")).ok, wb.sheet.input("E2")], [true, ""]);
// The workbook's structure.
await cmds.run("protect.workbook password=wb");
const STRUCTURE = "The workbook's structure is protected: sheets cannot be added, deleted, moved, renamed, hidden or copied (Review > Protect Workbook to unprotect it).";
eq("Protect Workbook: sheets stay as they are", [await msg("sheet.add name=More"), await msg("sheet.rename name=Sheet1 to=Main"), wb.pages().map((p) => p.name)], [STRUCTURE, STRUCTURE, ["Sheet1"]]);
// Saved and opened again; written as XLSX and read back.
const again = new Workbook(), cmds2 = new Commands(again);
await cmds2.execute("file.import", { text: (await cmds.run("file.export")).output as string });
eq("saved with the file", [again.page("Sheet1")!.protection, again.page("Sheet1")!.editRanges.map((r) => r.title), again.bookProtection?.hash === passwordHash("wb"), again.notes.get("B1").unlocked, (await cmds2.run("cell.set cell=A1 input=1")).message],
  [{ allow: [] }, ["Open", "Keyed"], true, true, PROTECTED]);
const x = await cmds.run("file.exportXlsx");
const back = new Workbook(), cmds3 = new Commands(back);
await cmds3.run(`file.importXlsx name=P.xlsx base64=${x.output}`);
eq("an XLSX file's", [back.page("Sheet1")!.protection, back.page("Sheet1")!.editRanges.map((r) => [r.title, r.hash === passwordHash("k")]), back.bookProtection?.hash === passwordHash("wb"), back.notes.get("B1").unlocked, back.notes.get("A1").unlocked],
  [{ allow: [] }, [["Open", false], ["Keyed", true]], true, true, undefined]);
await cmds.run("protect.workbook off=1 password=wb");
await cmds.run("protect.sheet off=1");
const said: string[] = [];
await cmds.script('SAY LOCKCELLS("F1", 0)\nSAY PROTECTSHEET("pw", "sort,insertRows")\nSAY ISPROTECTED() ISLOCKED("F1") ISLOCKED("G1")\nLET r = SETCELL("F1", 3)\nSAY GETCELL("F1")\nSAY UNPROTECTSHEET("pw")\nSAY ISPROTECTED()', (l) => said.push(l));
eq("from REXX", said, ["F1 unlocked", "Sheet1 protected with a password, allowing sort, insertRows", "1 0 1", "3", "Sheet1 unprotected", "0"]);

// expect: ok Excel's legacy hash
// expect: ok Protect Sheet
// expect: ok a locked cell refuses a change
// expect: ok an unlocked one takes it
// expect: ok a paste or a fill over a locked cell is taken back
// expect: ok a paste into unlocked cells is not
// expect: ok formatting, as allowed
// expect: ok what it does not allow
// expect: ok the view and the names are not guarded
// expect: ok Unprotect Sheet needs its password
// expect: ok an edit range's cells take a change
// expect: ok one with a password, once it is given
// expect: ok undo of an allowed change
// expect: ok Protect Workbook: sheets stay as they are
// expect: ok saved with the file
// expect: ok an XLSX file's
// expect: ok from REXX
