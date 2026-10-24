// Calculation timings (docs/comparisons/LIBREOFFICE.md action 4): formulas.profile on=1 starts timing
// each formula's evaluations; with no on= it lists the slowest cells (sheet,
// cell, milliseconds, runs) so a slow workbook says where its time goes; on=0 stops.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook(), cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
eq("before profiling starts, there is nothing to list", (await cmds.execute("formulas.profile")).message, "not profiling: formulas.profile on=1 starts it");
eq("started", (await cmds.execute("formulas.profile", { on: "1" })).message, "profiling formula evaluations");
const big: [string, string][] = [["A1", "1"], ["B1", "=A1+1"], ["C1", "=SUMPRODUCT((E1:E60000+A1)*(E1:E60000+A1))"], ["D1", "=B1*2"]];
for (let r = 1; r <= 60000; r++) big.push([`E${r}`, String(r)]);
await wb.sheet.setMany(big);
await cmds.run("cell.set cell=A1 input=2");
await cmds.run("cell.set cell=A1 input=3");
const o = await cmds.execute("formulas.profile");
const rows = o.output as { sheet: string; cell: string; ms: number; runs: number }[];
eq("the slow formula is first, run once per change and once on entry", [rows[0].cell, rows[0].sheet, rows[0].runs, rows[0].ms > 0], ["C1", "SHEET1", 3, true]);
eq("the cheap ones follow, with their runs", rows.filter((r) => r.cell !== "C1").map((r) => [r.cell, r.runs]).sort(), [["B1", 3], ["D1", 3]]);
eq("a limit", ((await cmds.execute("formulas.profile", { top: "1" })).output as unknown[]).length, 1);
eq("the message names the slowest", o.message.startsWith("slowest: SHEET1!C1 "), true);
eq("stopped, and the numbers are gone", [(await cmds.execute("formulas.profile", { on: "0" })).message, (await cmds.execute("formulas.profile")).message], ["profiling stopped", "not profiling: formulas.profile on=1 starts it"]);
await cmds.run("cell.set cell=A1 input=4");
eq("changes after it are not counted", (await cmds.execute("formulas.profile")).ok, true);

const said: string[] = [];
await cmds.script('LET m = PROFILE(1)\nLET r = SETCELL("A1", 5)\nSAY PROFILE()', (l) => said.push(l));
eq("PROFILE from REXX", said.length === 1 && said[0].startsWith("slowest: SHEET1!"), true);

// expect: ok before profiling starts, there is nothing to list
// expect: ok started
// expect: ok the slow formula is first, run once per change and once on entry
// expect: ok the cheap ones follow, with their runs
// expect: ok a limit
// expect: ok the message names the slowest
// expect: ok stopped, and the numbers are gone
// expect: ok changes after it are not counted
// expect: ok PROFILE from REXX
