// 6.11c-b: charts and drawings float over the grid at their cells: a chart
// has its cell (two columns right of its range until it is moved), chart.move
// and INSERTCHART's `at` say another, rows and columns inserted move it, a
// file keeps it, undo takes it back.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { readSheetFile, writeSheetFile } from "./sheet/file.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
const msg = async (line: string) => (await cmds.run(line)).message;
for (let r = 1; r <= 3; r++) { await cmds.run(`cell.set cell=A${r} input=x${r}`); await cmds.run(`cell.set cell=B${r} input=${r}`); }
await cmds.run("chart.insert kind=bar range=A1:B3");
eq("a chart's cell until it is moved: two columns right of its range, at its top row", wb.charts().map((c) => c.where), ["D1"]);
eq("chart.move", [await msg("chart.move index=1 at=F5"), wb.charts().map((c) => c.where)], ["Chart 1 moved to F5", ["F5"]]);
eq("INSERTCHART's cell", [await msg("chart.insert kind=line range=A1:B3 at=H2"), wb.charts().map((c) => c.where)], ["inserted a line chart of A1:B3", ["F5", "H2"]]);
eq("refusals", [await msg("chart.move index=9 at=A1"), await msg("chart.move index=1 at=nonsense")], ["no chart 9", "nonsense is not a cell on the sheet"]);
await cmds.run("rows.insert at=1");
eq("a row inserted above moves it", wb.charts().map((c) => c.where), ["F6", "H3"]);
const back = readSheetFile(JSON.parse(JSON.stringify(writeSheetFile(wb))));
eq("a file keeps it", back.sheets[0].charts.map((c) => c.at), [{ col: 5, row: 5 }, { col: 7, row: 2 }]);
await cmds.run("edit.undo");
eq("undo takes the insert back", wb.charts().map((c) => c.where), ["F5", "H2"]);
const said: string[] = [];
eq("MOVECHART from REXX", [(await cmds.script('SAY MOVECHART(2, "C9")', (l) => said.push(l))).message, said, wb.charts().map((c) => c.where)], ["ran", ["Chart 2 moved to C9"], ["F5", "C9"]]);
// expect: ok a chart's cell until it is moved: two columns right of its range, at its top row
// expect: ok chart.move
// expect: ok INSERTCHART's cell
// expect: ok refusals
// expect: ok a row inserted above moves it
// expect: ok a file keeps it
// expect: ok undo takes the insert back
// expect: ok MOVECHART from REXX
