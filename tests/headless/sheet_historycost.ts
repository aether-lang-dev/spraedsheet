// What the history costs (docs/comparisons/IRONCALC.md action 4, and the mutation
// journal, docs/mutation_journal.md): a step of undo is a list of journal entries
// (journal.ts), so this measures, in map entries held (Workbook.historyCost:
// what each step's entries were and became), what each kind of change leaves
// behind. Before the journal a step was a whole snapshot (a one-cell edit on
// 2,000 cells kept 2,021 entries, an inserted row 3,043, a script of fifty
// changes 3,043); the numbers below are the journal's.
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
eq("a sheet of 2,000 cells opened by a bulk set is one step holding nothing (the set was not a step)", [start.undo, start.redo, start.entries], [1, 0, 0]);

// One-cell edit: the whole inputs map again, and the styles.
let before = was();
await cmds.run("cell.set cell=B1 input=x");
eq("a one-cell edit keeps that cell: what it was and became", was() - before, 1);
before = was();
await cmds.run("cell.set cell=B2 input=y");
eq("and so does the next", was() - before, 1);

// A look: the inputs map is new for the first step after an edit, and shared by the second.
await cmds.run("format.fill range=A1:A10 color=yellow");
before = was();
await cmds.run("format.fill range=A1:A10 color=red");
eq("a look after a look keeps the notes of the ten cells, was and became", was() - before, 20);

// A large paste.
await cmds.run("edit.copy range=A1:A1000");
before = was();
await cmds.run("edit.paste at=D1");
eq("a large paste keeps the cells it set and the notes", was() - before, 1000 + 30);

// A structural edit.
before = was();
await cmds.run("rows.insert at=5 count=1");
eq("an inserted row after a paste: every cell that moved, where it was and where it went", was() - before, 3034);

// A scripted group is one step.
before = was();
const steps = cost().undo;
const lines: string[] = [];
for (let i = 1; i <= 50; i++) lines.push(`LET r = SETCELL("F${i}", ${i})`);
const ran = await cmds.script(lines.join("\n"), () => {});
eq("a script of fifty changes is one step holding the fifty cells", [ran.ok, cost().undo - steps, was() - before], [true, 1, 50]);

// Redo keeps what was undone.
const undone = cost();
await cmds.run("edit.undo");
await cmds.run("edit.undo");
const after = cost();
eq("undo moves a step to redo; the entries held are the same", [after.undo, after.redo, after.entries], [undone.undo - 2, 2, undone.entries]);

// The bound: HISTORY steps.
for (let i = 0; i < 120; i++) await cmds.run(`cell.set cell=C${i + 1} input=${i}`);
const full = cost();
eq("steps are bounded, and a hundred one-cell steps hold a hundred cells' worth, not a hundred sheets", [full.undo, full.entries < 1000], [100, true]);

// The bound: past historyEntries the oldest steps go. A step that moves every
// cell (an inserted row at the top) is the costly one now.
{
  const small = new Workbook(), sc = new Commands(small);
  await sc.run("formulas.use language=Excel");
  await small.sheet.setMany(rows);
  small.historyEntries = 5 * N;
  for (let i = 0; i < 20; i++) await sc.run("rows.insert at=1 count=1");
  const c = small.historyCost();
  eq("a history over its entry budget drops the oldest steps, and stays under it", [c.undo >= 1 && c.undo <= 5, c.entries <= 5 * N], [true, true]);
  let undone = 0;
  while (small.canUndo) { await sc.run("edit.undo"); undone++; }
  eq("and what is kept still undoes, step by step", [undone === c.undo, small.sheet.input(`A${20 - c.undo + 1}`), small.sheet.input(`A${20 - c.undo}`)], [true, "1", ""]);
}

// expect: ok a sheet of 2,000 cells opened by a bulk set is one step holding nothing (the set was not a step)
// expect: ok a one-cell edit keeps that cell: what it was and became
// expect: ok and so does the next
// expect: ok a look after a look keeps the notes of the ten cells, was and became
// expect: ok a large paste keeps the cells it set and the notes
// expect: ok an inserted row after a paste: every cell that moved, where it was and where it went
// expect: ok a script of fifty changes is one step holding the fifty cells
// expect: ok undo moves a step to redo; the entries held are the same
// expect: ok steps are bounded, and a hundred one-cell steps hold a hundred cells' worth, not a hundred sheets
// expect: ok a history over its entry budget drops the oldest steps, and stays under it
// expect: ok and what is kept still undoes, step by step
