// A recording keeps its context (docs/comparisons/UNIVER.md action 2): a macro
// recorded on one sheet, with the selection somewhere, replays on that sheet
// and at that cell, not on whatever sheet is active and whatever cell is
// selected when it is run.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
await cmds.run("sheet.add name=Costs");
await cmds.run("sheet.activate name=Sheet1");

await cmds.run("view.select range=C3");
await cmds.run("macro.record name=Mark");
await cmds.run("range.clear range=A1:B1");
await cmds.run("cell.set cell=A1 input=marked");
await cmds.run("insert.shape shape=rectangle");
const stopped = await cmds.execute("macro.stop");
eq("two recorded commands and a shape, said as the person did them", stopped.message, "Mark: 3 commands recorded");
const text = String(stopped.output);
eq("the program starts on the sheet it was recorded on", text.split("\n")[1], 'ADDRESS SHEET "sheet.activate name=Sheet1"');
eq("and the shape names the cell it was put at, not the selection of a later run", text.split("\n").filter((l) => l.includes("insert.shape")), ['ADDRESS SHEET "insert.shape shape=rectangle at=C3"']);

// Run from elsewhere: another sheet active, the selection moved.
await cmds.run("sheet.activate name=Costs");
await cmds.run("view.select range=F9");
await cmds.run("cell.set cell=A1 input=costs");
const ran = await cmds.execute("macro.run", { name: "Mark" });
eq("the macro ran", [ran.ok, ran.message], [true, "Mark ran"]);
eq("it acted on the sheet it was recorded on; the other sheet's cell is as it was", [wb.page("Sheet1")!.sheet.input("A1"), wb.page("Costs")!.sheet.input("A1")], ["marked", "costs"]);
eq("the shape is on Sheet1 at C3 again (the recorded one and the replayed), none on Costs", [wb.page("Sheet1")!.drawings.map((d) => d.at), wb.page("Costs")!.drawings.length], [[{ col: 2, row: 2 }, { col: 2, row: 2 }], 0]);
eq("the recorded sheet is the active one after", wb.active, "Sheet1");

// With one sheet there is nothing to say: no context line.
const solo = new Commands(new Workbook());
await solo.run("macro.record name=One");
await solo.run("cell.set cell=A1 input=1");
eq("one sheet: the program is just what was done", String((await solo.execute("macro.stop")).output).split("\n").length, 2);

// A macro recorded with nothing in it stays empty (no context line to run).
await cmds.run("macro.record name=Empty");
const empty = await cmds.execute("macro.stop");
eq("nothing recorded: just the heading", [empty.message, String(empty.output).split("\n").length], ["Empty: 0 commands recorded", 1]);

// A recording that itself switches sheet first is not given a second line.
await cmds.run("macro.record name=Switch");
await cmds.run("sheet.activate name=Costs");
await cmds.run("cell.set cell=B2 input=x");
const sw = String((await cmds.execute("macro.stop")).output).split("\n");
eq("a recording that starts by switching sheet keeps its own first line", [sw.length, sw[1]], [3, 'ADDRESS SHEET "sheet.activate name=Costs"']);

// expect: ok two recorded commands and a shape, said as the person did them
// expect: ok the program starts on the sheet it was recorded on
// expect: ok and the shape names the cell it was put at, not the selection of a later run
// expect: ok the macro ran
// expect: ok it acted on the sheet it was recorded on; the other sheet's cell is as it was
// expect: ok the shape is on Sheet1 at C3 again (the recorded one and the replayed), none on Costs
// expect: ok the recorded sheet is the active one after
// expect: ok one sheet: the program is just what was done
// expect: ok nothing recorded: just the heading
// expect: ok a recording that starts by switching sheet keeps its own first line
