// The Extra controls (8.4): the rows under the sheet tabs (the sheet
// buttons, the panels' tabs) shown or not, collapsed as a workbook starts;
// view.extraControls on=1|0 (none toggles); SHOWEXTRACONTROLS and
// GETEXTRACONTROLS from REXX; a view, not a step to undo; saved with the
// workbook and read back, written only when shown.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { readSheetFile, writeSheetFile } from "./sheet/file.ts";
import { MENUS } from "./sheet/menus.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
const x = async (p: Record<string, string>) => (await cmds.execute("view.extraControls", p)).message;
eq("collapsed as it starts", wb.extraControls, false);
eq("on, off, and toggled", [await x({ on: "1" }), await x({ on: "0" }), await x({}), wb.extraControls], ["extra controls shown", "extra controls hidden", "extra controls shown", true]);
await cmds.run("cell.set cell=A1 input=1");
await x({ on: "0" });
eq("a view: not a step to undo", wb.undoLabel, "Set a cell");
const said: string[] = [];
eq("SHOWEXTRACONTROLS and GETEXTRACONTROLS from REXX", [(await cmds.script("SAY SHOWEXTRACONTROLS(1)\nSAY GETEXTRACONTROLS()\nSAY SHOWEXTRACONTROLS(0) GETEXTRACONTROLS()", (l) => said.push(l))).message, said], ["ran", ["1", "1", "0 0"]]);
eq("written only when shown", [writeSheetFile(wb).extraControls, (await x({ on: "1" }), writeSheetFile(wb).extraControls)], [undefined, true]);
eq("and read back", [readSheetFile(JSON.parse(JSON.stringify(writeSheetFile(wb)))).extraControls, readSheetFile({}).extraControls], [true, false]);
wb.reset();
eq("a new workbook starts collapsed", wb.extraControls, false);
eq("the View menu has it, and every panel", MENUS.find((m) => m.label === "View")!.items.filter((e) => e !== "-" && (e.label === "Extra controls" || e.label.endsWith(" panel"))).length, 16);

// expect: ok collapsed as it starts
// expect: ok on, off, and toggled
// expect: ok a view: not a step to undo
// expect: ok SHOWEXTRACONTROLS and GETEXTRACONTROLS from REXX
// expect: ok written only when shown
// expect: ok and read back
// expect: ok a new workbook starts collapsed
// expect: ok the View menu has it, and every panel
