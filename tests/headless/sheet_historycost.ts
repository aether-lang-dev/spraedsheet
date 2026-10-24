// What the history costs (docs/comparisons/IRONCALC.md action 4): the workbook keeps a
// whole snapshot per undo step (HISTORY deep), so this measures, in map
// entries held (Workbook.historyCost: each saved map once however many
// steps share it), what each kind of change leaves behind, before anyone
// chooses patches over snapshots.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const N = 2000;
const wb = new Workbook();
const cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
const rows: [string, string][] = [];
for (let r = 1; r <= N; r++) rows.push([`A${r}`, String(r)]);
await wb.sheet.setMany(rows);
const cost = () => wb.historyCost();
const was = () => cost().entries;
const start = cost();
// The workbook's styles (Normal, Good, Bad ...) are copied into every step too.
const S = wb.styles.size;
eq("a sheet of 2,000 cells opened by a bulk set is one step, holding the styles and the empty sheet", [start.undo, start.redo, start.entries], [1, 0, S]);

// One-cell edit: the whole inputs map again, and the styles.
let before = was();
await cmds.run("cell.set cell=B1 input=x");
eq("a one-cell edit keeps every cell of the sheet (a snapshot, not a patch) and the styles", was() - before, N + S);
before = was();
await cmds.run("cell.set cell=B2 input=y");
eq("and so does the next", was() - before, N + 1 + S);

// A look: the inputs map is new for the first step after an edit, and shared by the second.
await cmds.run("format.fill range=A1:A10 color=yellow");
before = was();
await cmds.run("format.fill range=A1:A10 color=red");
eq("a look after a look keeps the inputs shared: the notes and the styles only", was() - before, 10 + S);

// A large paste.
await cmds.run("edit.copy range=A1:A1000");
before = was();
await cmds.run("edit.paste at=D1");
eq("a large paste: the notes and the styles, the inputs shared with the step before", was() - before, 10 + S);

// A structural edit.
before = was();
await cmds.run("rows.insert at=5 count=1");
eq("an inserted row after a paste: every input again", was() - before, N + 2 + 1000 + S + 20);

// A scripted group is one step.
before = was();
const steps = cost().undo;
const lines: string[] = [];
for (let i = 1; i <= 50; i++) lines.push(`LET r = SETCELL("F${i}", ${i})`);
const ran = await cmds.script(lines.join("\n"), () => {});
eq("a script of fifty changes is one step and one snapshot", [ran.ok, cost().undo - steps, was() - before], [true, 1, N + 2 + 1000 + S + 20]);

// Redo keeps what was undone.
const undone = cost();
await cmds.run("edit.undo");
await cmds.run("edit.undo");
const after = cost();
eq("undo moves a step to redo; the entries held stay about the same", [after.undo, after.redo, after.entries >= undone.entries - N], [undone.undo - 2, 2, true]);

// The bound: HISTORY steps of a snapshot each.
for (let i = 0; i < 120; i++) await cmds.run(`cell.set cell=C${i + 1} input=${i}`);
const full = cost();
eq("steps are bounded; entries are steps times cells", [full.undo, full.entries > 90 * N], [100, true]);

// The bound: past historyEntries the oldest steps go (on a 100,000-cell sheet
// the page ran out of memory at the 47th edit before this).
{
  const small = new Workbook(), sc = new Commands(small);
  await sc.run("formulas.use language=Excel");
  await small.sheet.setMany(rows);
  small.historyEntries = 5 * N;
  for (let i = 0; i < 20; i++) await sc.run(`cell.set cell=B${i + 1} input=${i}`);
  const c = small.historyCost();
  eq("a history over its entry budget drops the oldest steps, and stays under it", [c.undo >= 2 && c.undo <= 5, c.entries <= 5 * N], [true, true]);
  let undone = 0;
  while (small.canUndo) { await sc.run("edit.undo"); undone++; }
  eq("and what is kept still undoes, step by step", [undone === c.undo, small.sheet.input("B20"), small.sheet.input("B1")], [true, "", "0"]);
}

// expect: ok a sheet of 2,000 cells opened by a bulk set is one step, holding the styles and the empty sheet
// expect: ok a one-cell edit keeps every cell of the sheet (a snapshot, not a patch) and the styles
// expect: ok and so does the next
// expect: ok a look after a look keeps the inputs shared: the notes and the styles only
// expect: ok a large paste: the notes and the styles, the inputs shared with the step before
// expect: ok an inserted row after a paste: every input again
// expect: ok a script of fifty changes is one step and one snapshot
// expect: ok undo moves a step to redo; the entries held stay about the same
// expect: ok steps are bounded; entries are steps times cells
// expect: ok a history over its entry budget drops the oldest steps, and stays under it
// expect: ok and what is kept still undoes, step by step
