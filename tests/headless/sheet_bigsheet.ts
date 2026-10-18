// Excel's sheet (1.4): 16,384 columns by 1,048,576 rows, as an address
// space the sheet keeps sparse: its last cell set and read, a formula over
// cells far apart, a whole column cleared and its sum taken where it is
// used, a typed-out column clipped there, a range past what the app works
// over at once refused, the status figures of a whole column, a far cell
// saved and opened again and through XLSX, SETCELL from REXX; each quick.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
const t0 = Date.now();
await cmds.run("formulas.use language=Excel");
eq("its last cell set and read", [(await cmds.run("cell.set cell=XFD1048576 input=7")).ok, show(wb.sheet.value("XFD1048576"))], [true, "7"]);
await cmds.run("cell.set cell=B2 input=5");
await cmds.run("cell.set cell=B500000 input=10");
await cmds.run('cell.set cell=C1 input="=B2+B500000+XFD1048576"');
eq("a formula over cells far apart", show(wb.sheet.value("C1")), "22");
eq("the status figures of a whole column, quickly", [wb.stats([{ col: 1, row: 0 }, { col: 1, row: 1048575 }]).sum, wb.stats([{ col: 1, row: 0 }, { col: 1, row: 1048575 }]).count], [15, 2]);
// A sheet used to row 5,000: a whole column is as far as that.
const used = new Workbook(), uc = new Commands(used);
await uc.run("cell.set cell=B5000 input=4");
await uc.run("cell.set cell=B7 input=3");
eq("a whole column as far as the sheet is used", [used.extent(), (await uc.run("range.clear range=B:B")).message.split(":").slice(0, 2).join(":")], [{ rows: 5000, cols: 2 }, "cleared B1:B5000"]);
await uc.run("cell.set cell=B5000 input=4");
eq("a typed-out column clipped there too", (await uc.run("range.clear range=B1:B1048576")).message.split(":").slice(0, 2).join(":"), "cleared B1:B5000");
eq("cleared", [show(used.sheet.value("B5000")), show(used.sheet.value("B7"))], ["", ""]);
eq("a range past what the app works over at once", (await cmds.run("format.font range=A1:XFD1048576 bold=1")).message, "A1:XFD1048576 is 17,179,869,184 cells where the sheet is used: more than the 5,000,000 this app works over at once");
const again = new Workbook(), cmds2 = new Commands(again);
await cmds2.execute("file.import", { text: (await cmds.run("file.export")).output as string });
eq("saved and opened again", show(again.sheet.value("XFD1048576")), "7");
const x = await cmds.run("file.exportXlsx");
const back = new Workbook(), cmds3 = new Commands(back);
await cmds3.run(`file.importXlsx name=Big.xlsx base64=${x.output}`);
eq("and through XLSX", [show(back.sheet.value("XFD1048576")), back.sheet.input("C1")], ["7", "=B2+B500000+XFD1048576"]);
const said: string[] = [];
await cmds.script('CALL SETCELL "ZZ100000", 1\nSAY GETCELL("ZZ100000")', (l) => said.push(l));
eq("from REXX", said, ["1"]);
eq("all of it in a few seconds", Date.now() - t0 < 20000, true);

// expect: ok its last cell set and read
// expect: ok a formula over cells far apart
// expect: ok the status figures of a whole column, quickly
// expect: ok a whole column as far as the sheet is used
// expect: ok a typed-out column clipped there too
// expect: ok cleared
// expect: ok a range past what the app works over at once
// expect: ok saved and opened again
// expect: ok and through XLSX
// expect: ok from REXX
// expect: ok all of it in a few seconds
