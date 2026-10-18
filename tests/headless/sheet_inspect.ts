// INSPECTCELLS (7.7): a range's cells as two workbooks are compared, the
// one driven headless and the one through the ui: each its input, what it
// shows, its look; the same work comes to the same JSON.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const work = async (cmds: Commands) => {
  await cmds.run("cell.set cell=A1 input=3");
  await cmds.run('cell.set cell=A2 input="=A1*2"');
  await cmds.run("format.font range=A2 bold=1");
};
const one = new Workbook(), c1 = new Commands(one), two = new Workbook(), c2 = new Commands(two);
await work(c1);
await work(c2);
const a = (await c1.run("range.inspect range=A1:B3")).output, b = (await c2.run("range.inspect range=A1:B3")).output;
eq("each cell's input, what it shows, its look", a, [{ cell: "A1", input: "3", shown: "3" }, { cell: "A2", input: "=A1*2", shown: "6", look: { bold: true } }]);
eq("the same work, the same cells", JSON.stringify(a) === JSON.stringify(b), true);
const said: string[] = [];
await c1.script('SAY INSPECTCELLS("A1")', (l) => said.push(l));
eq("from REXX", said, ['[{"cell":"A1","input":"3","shown":"3"}]']);

// expect: ok each cell's input, what it shows, its look
// expect: ok the same work, the same cells
// expect: ok from REXX
