// An error value typed into a cell is that error, as in Excel (#N/A typed is
// what NA() gives: ISNA is TRUE of it, ERROR.TYPE numbers it), in any case;
// with an apostrophe in front it is text. #NULL! and #GETTING_DATA are
// Excel's too (ERROR.TYPE 1 and 8). (IronCalc's ERROR.TYPE and
// IS_INFORMATION workbooks found it: docs/comparisons/COMPATIBILITY.md.)
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook(), cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
const set = async (k: string, v: string) => { await cmds.run(`cell.set cell=${k} input=${v}`); };
await set("A1", "#N/A"); await set("A2", "#null!"); await set("A3", "'#N/A"); await set("A4", "#GETTING_DATA");
await set("B1", "=ISNA(A1)"); await set("B2", "=ERROR.TYPE(A2)"); await set("B3", "=ISTEXT(A3)"); await set("B4", "=ERROR.TYPE(A4)");
const v = (k: string) => show(wb.sheet.value(k));
eq("#N/A typed is the error, #null! typed is #NULL!", [v("A1"), v("A2"), v("B1"), v("B2")], ["#N/A", "#NULL!", "TRUE", "1"]);
eq("with an apostrophe it is text; #GETTING_DATA is 8", [v("A3"), v("B3"), v("B4")], ["#N/A", "TRUE", "8"]);

// expect: ok #N/A typed is the error, #null! typed is #NULL!
// expect: ok with an apostrophe it is text; #GETTING_DATA is 8
