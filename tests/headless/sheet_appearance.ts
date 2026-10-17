// Dark mode (3.0.6): the appearance a view of the workbook (system, light
// or dark), refused otherwise; in dark the sheet's cells drawn white with
// black text (the chrome is the page's), a cell's own fill and colour still
// winning; SETAPPEARANCE and GETAPPEARANCE from REXX.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
eq("system to begin with", [wb.appearance, (await cmds.run("view.appearance mode=dark")).message, wb.appearance, (await cmds.run("view.appearance mode=dusk")).message], ["system", "appearance dark", "dark", "dusk is not system, light or dark"]);
await cmds.run("cell.set cell=A1 input=7");
await cmds.run("format.fill range=B1 color=yellow");
wb.dark = true;
eq("dark: the sheet white with black text, a cell's own fill after it", [wb.styleOf("A1").startsWith("background: #ffffff; color: #000000"), wb.styleOf("B1").indexOf("#ffff00") > wb.styleOf("B1").indexOf("#ffffff")], [true, true]);
wb.dark = false;
eq("light: no ground of its own", wb.styleOf("A1").includes("#ffffff"), false);
eq("View > Animations: on as it starts, off, on", [wb.animations, (await cmds.run("view.animations off=1")).message, wb.animations, (await cmds.run("view.animations")).message], [true, "animations off", false, "animations on"]);
const said: string[] = [];
await cmds.script('SAY SETAPPEARANCE("light")\nSAY GETAPPEARANCE()\nSAY SETAPPEARANCE("system")\nSAY SETANIMATIONS(0) GETANIMATIONS()\nSAY SETANIMATIONS(1) GETANIMATIONS()', (l) => said.push(l));
eq("from REXX", said, ["light", "light light", "system", "0 0", "1 1"]);

// expect: ok system to begin with
// expect: ok dark: the sheet white with black text, a cell's own fill after it
// expect: ok light: no ground of its own
// expect: ok View > Animations: on as it starts, off, on
// expect: ok from REXX
