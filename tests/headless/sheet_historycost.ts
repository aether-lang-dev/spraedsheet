// What the history costs (docs/comparisons/IRONCALC.md action 4, and the mutation
// journal, docs/mutation_journal.md): a step of undo is a list of journal entries
// (journal.ts), so this measures, in map entries held (Workbook.historyCost:
// what each step's entries were and became), what each kind of change leaves
// behind. Before the journal a step was a whole snapshot (a one-cell edit on
// 2,000 cells kept 2,021 entries, an inserted row 3,043, a script of fifty
// changes 3,043); the numbers below are the journal's.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { copies, DEPTH } from "./sheet/overlay.ts";
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
eq("a look after a look keeps the ten cells' notes, as they were and became", was() - before, 10);

// A large paste.
await cmds.run("edit.copy range=A1:A1000");
before = was();
await cmds.run("edit.paste at=D1");
eq("a large paste keeps the cells it set and the ten notes it added", was() - before, 1000 + 10);

// A structural edit.
before = was();
await cmds.run("rows.insert at=5 count=1");
eq("an inserted row after a paste: every cell that moved, where it was and where it went", was() - before, 3008);

// A cell style, a name and a table: each the one entry that changed.
before = was();
await cmds.run("style.define name=Mine bold=1");
eq("a defined cell style keeps that style, not the workbook's whole gallery", was() - before, 1);
await cmds.run("names.define name=FIRST range=F1:F9");
before = was();
await cmds.run("names.define name=TOTAL range=F1:F9");
eq("a second defined name keeps that name, not the first again", was() - before, 1);
await cmds.run("table.define name=Orders range=D1:D10 headers=0");
before = was();
await cmds.run("table.define name=Sales range=E1:E10 headers=0");
eq("a second defined table keeps that table", was() - before, 1);

// Widths and heights: the columns and rows that changed.
await cmds.run("columns.width range=A:C width=12");
before = was();
await cmds.run("columns.width range=D1 width=20");
eq("a column widened keeps that column, not every width of the sheet", was() - before, 1);
await cmds.run("rows.height range=1:3 height=30");
before = was();
await cmds.run("rows.height range=4:4 height=40");
eq("a row made taller keeps that row", was() - before, 1);

// A scripted group is one step.
before = was();
const steps = cost().undo;
const lines: string[] = [];
for (let i = 1; i <= 50; i++) lines.push(`LET r = SETCELL("F${i}", ${i})`);
const ran = await cmds.script(lines.join("\n"), () => {});
eq("a script of fifty changes is one step holding the fifty cells", [ran.ok, cost().undo - steps, was() - before], [true, 1, 50]);

// Recorded at the mutation (overlay.ts): a command's snapshot is the cells it
// set, layered on the one before, not a copy of the sheet's 2,000 cells; the
// whole map is copied once in DEPTH snapshots (undo and redo the same). Before,
// a hundred one-cell edits made a hundred copies.
{
  const big = new Workbook(), bc = new Commands(big);
  await bc.run("formulas.use language=Excel");
  await big.sheet.setMany(rows);
  await bc.run("cell.set cell=B1 input=0");
  const at = copies.maps;
  for (let i = 1; i <= 100; i++) await bc.run(`cell.set cell=B${i} input=${i}`);
  const made = copies.maps - at, most = Math.ceil(100 / DEPTH);
  if (made > 2 * most) print(`note: ${made} copies`);
  eq("a hundred one-cell edits copy the sheet's cells a few times (once in DEPTH), not a hundred", made <= 2 * most, true);
  const u = copies.maps;
  for (let i = 0; i < 40; i++) await bc.run("edit.undo");
  for (let i = 0; i < 40; i++) await bc.run("edit.redo");
  eq("and forty undos and redos as few", [copies.maps - u <= 2 * Math.ceil(80 / DEPTH), big.sheet.input("B60"), big.sheet.input("B100")], [true, "60", "100"]);
}

// The lists (rules, charts, merges ...) and sets (hidden rows and columns) of
// a sheet: the items that came and went, not the whole list again.
{
  const v = new Workbook(), vc = new Commands(v);
  await vc.run("formulas.use language=Excel");
  for (let i = 1; i <= 20; i++) await vc.run(`format.highlight range=A${i} op=greater value=${i}`);
  for (let i = 1; i <= 40; i++) await vc.run(`rows.hide range=A${2 * i}`);
  for (let i = 1; i <= 10; i++) await vc.run(`cells.merge range=C${2 * i - 1}:D${2 * i - 1}`);
  const w = () => v.historyCost().entries;
  let b = w();
  await vc.run("format.highlight range=B1 op=greater value=0");
  eq("a rule added to twenty keeps that rule", w() - b, 1);
  b = w();
  await vc.run("rows.hide range=A99");
  eq("a row hidden beside forty keeps that row", w() - b, 1);
  b = w();
  await vc.run("cells.merge range=F1:G1");
  eq("a merge added to ten keeps that merge", w() - b, 1);
  await vc.run("edit.undo"); await vc.run("edit.undo"); await vc.run("edit.undo");
  const back = v.page(v.active)!;
  eq("and undo takes each back", [back.rules.length, back.hiddenRows.has(98), back.merges.length], [20, false, 10]);
  await vc.run("edit.redo"); await vc.run("edit.redo"); await vc.run("edit.redo");
  eq("and redo does each again", [back.rules.length, v.page(v.active)!.hiddenRows.size, v.page(v.active)!.merges.length], [21, 41, 11]);
}

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
// expect: ok a look after a look keeps the ten cells' notes, as they were and became
// expect: ok a large paste keeps the cells it set and the ten notes it added
// expect: ok an inserted row after a paste: every cell that moved, where it was and where it went
// expect: ok a defined cell style keeps that style, not the workbook's whole gallery
// expect: ok a second defined name keeps that name, not the first again
// expect: ok a second defined table keeps that table
// expect: ok a column widened keeps that column, not every width of the sheet
// expect: ok a row made taller keeps that row
// expect: ok a script of fifty changes is one step holding the fifty cells
// expect: ok a hundred one-cell edits copy the sheet's cells a few times (once in DEPTH), not a hundred
// expect: ok and forty undos and redos as few
// expect: ok a rule added to twenty keeps that rule
// expect: ok a row hidden beside forty keeps that row
// expect: ok a merge added to ten keeps that merge
// expect: ok and undo takes each back
// expect: ok and redo does each again
// expect: ok undo moves a step to redo; the entries held are the same
// expect: ok steps are bounded, and a hundred one-cell steps hold a hundred cells' worth, not a hundred sheets
// expect: ok a history over its entry budget drops the oldest steps, and stays under it
// expect: ok and what is kept still undoes, step by step
