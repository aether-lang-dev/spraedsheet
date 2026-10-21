// A read view with an edit mode (blog-entry-todo.md B.7, the view.editMode
// command): the workbook is changed only while edit mode is on. In the read
// view every command that changes the workbook is refused, with why; looking
// (reads, moving between sheets, the view) still works.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
await cmds.run("cell.set cell=A1 input=Item");
await cmds.run("cell.set cell=B1 input=2");
eq("edit mode is on to start with", [wb.editMode, (await cmds.execute("view.editModeState")).output], [true, 1]);

const off = await cmds.execute("view.editMode", { on: "0" });
eq("switched off: the read view", [off.ok, off.message, wb.editMode], [true, "read view: edit mode is off", false]);
const READ = "read view: switch edit mode on to change the workbook";
const typed = await cmds.execute("cell.set", { cell: "A1", input: "changed" });
eq("typing in a cell changes nothing, and says why", [typed.ok, typed.message, wb.sheet.input("A1")], [false, READ, "Item"]);
eq("so do a clear, a format, a fill, a new sheet and a paste", [
  (await cmds.execute("range.clear", { range: "A1:B1" })).message,
  (await cmds.execute("format.fill", { range: "A1", color: "yellow" })).message,
  (await cmds.execute("edit.fillDown", { range: "A1:B2" })).message,
  (await cmds.execute("sheet.add", { name: "Costs" })).message,
], [READ, READ, READ, READ]);
eq("the workbook is as it was", [wb.sheet.input("A1"), wb.sheet.input("B1"), wb.sheetNames().length], ["Item", "2", 1]);

// Looking still works.
eq("a read, a selection, a look at the view", [
  (await cmds.execute("cell.get", { cell: "B1" })).ok,
  (await cmds.execute("view.select", { range: "A1:B1" })).ok,
  (await cmds.execute("view.zoom", { percent: "120" })).ok,
], [true, true, true]);

// Not through a script, either.
const said: string[] = [];
const refused = await cmds.script('LET r = SETCELL("A1", "x")', (l) => said.push(l));
eq("a script that sets a cell is refused whole", [refused.ok, refused.message.includes(READ), wb.sheet.input("A1")], [false, true, "Item"]);

// Switched on, it does.
const on = await cmds.execute("view.editMode", { on: "1" });
eq("switched on", [on.ok, on.message, wb.editMode], [true, "edit mode is on", true]);
await cmds.run("cell.set cell=A1 input=changed");
eq("typing changes the cell", wb.sheet.input("A1"), "changed");

// With no on=, it toggles.
await cmds.execute("view.editMode");
eq("no on=: toggled off", wb.editMode, false);
await cmds.execute("view.editMode");
eq("and on", wb.editMode, true);
eq("refused: on= that is neither 1 nor 0", (await cmds.execute("view.editMode", { on: "maybe" })).message, "on= is 1 (edit mode) or 0 (the read view), not maybe");

// REXX.
const lines: string[] = [];
await cmds.script('LET m = SETEDITMODE(0)\nSAY m\nSAY GETEDITMODE()\nLET m = SETEDITMODE(1)\nSAY m\nSAY GETEDITMODE()\nLET r = SETCELL("C1", 5)\nSAY GETCELL("C1")', (l) => lines.push(l));
eq("SETEDITMODE and GETEDITMODE from REXX; SETCELL works once it is on", lines, ["read view: edit mode is off", "0", "edit mode is on", "1", "5"]);

// A file opened or a workbook made is not a change to the one being read.
await cmds.execute("view.editMode", { on: "0" });
eq("a file can still be opened in the read view", (await cmds.execute("file.open", { sample: "budget" })).ok, true);
eq("and the read view stays", wb.editMode, false);

// expect: ok edit mode is on to start with
// expect: ok switched off: the read view
// expect: ok typing in a cell changes nothing, and says why
// expect: ok so do a clear, a format, a fill, a new sheet and a paste
// expect: ok the workbook is as it was
// expect: ok a read, a selection, a look at the view
// expect: ok a script that sets a cell is refused whole
// expect: ok switched on
// expect: ok typing changes the cell
// expect: ok no on=: toggled off
// expect: ok and on
// expect: ok refused: on= that is neither 1 nor 0
// expect: ok SETEDITMODE and GETEDITMODE from REXX; SETCELL works once it is on
// expect: ok a file can still be opened in the read view
// expect: ok and the read view stays
