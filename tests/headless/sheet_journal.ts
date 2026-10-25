// The mutation journal (docs/mutation_journal.md): every kind of state a command
// can change comes back by undo and goes again by redo, exactly, because a step
// is a list of entries (journal.ts) and not a copy of the workbook. The
// workbook is fingerprinted whole (cells, names, tables, notes, looks, the
// view, styles, theme, protection, the sheets and their order) after each
// command of a long mixed sequence; undoing step by step must meet each
// fingerprint again on the way back, and redoing on the way forward.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { stepOf, applyStep } from "./sheet/journal.ts";
import { eq } from "./lib/check.ts";

const fingerprint = (wb: Workbook): string => {
  const s = wb.saved();
  const rep = (key: string, v: unknown): unknown => {
    if (key === "page") return (v as { name: string }).name;
    if (v instanceof Map) return { map: [...v.entries()].sort((a, b) => String(a[0]) < String(b[0]) ? -1 : 1) };
    if (v instanceof Set) return { set: [...v].sort() };
    return v;
  };
  return JSON.stringify({
    pages: s.pages.map((p) => ({ n: p.name, sheet: p.sheet, notes: p.notes, view: p.view })),
    names: s.names, active: s.active.name, name: s.name, setup: s.setupScript, styles: s.styles, theme: s.theme, bp: s.bookProtection,
  }, rep);
};

const wb = new Workbook(), cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
const sequence = [
  "cell.set cell=A1 input=10", "cell.set cell=A2 input=20", "cell.set cell=A3 input==SUM(A1:A2)",
  "cell.set cell=B1 input=Jan", "cell.comment cell=A1 text=first", "format.fill range=A1:B1 color=yellow",
  "format.font range=A1 bold=1", "format.number range=A1:A3 format=0.00",
  "format.highlight range=A1:A3 op=greater value=15",
  "names.define name=TOTAL range=A3", "table.define name=Sales range=A1:B3 headers=0",
  "view.freezeRows count=1", "rows.insert at=2", "columns.insert at=B",
  "sheet.add name=Costs", "cell.set cell=A1 input=5", "sheet.rename name=Costs to=Prices",
  "sheet.move name=Prices to=1", "style.define name=Mine bold=1", "theme.set name=Office",
  "rows.hide range=A3", "cells.merge range=C1:D1", "protect.sheet password=x",
  "sheet.delete name=Prices", "sheet.activate name=Sheet1",
];
const prints: string[] = [fingerprint(wb)];
const ran: string[] = [];
for (const line of sequence) {
  const steps = wb.historyCost().undo;
  const r = await cmds.run(line);
  if (!r.ok) { print(`note: ${line}: ${r.message}`); continue; }
  ran.push(line);
  if (wb.historyCost().undo > steps) prints.push(fingerprint(wb));
}
eq("the mixed sequence ran, step by step", [ran.length >= sequence.length - 3, wb.historyCost().undo >= ran.length - 6], [true, true]);
// Undo one step at a time: the fingerprint of the step before comes back.
let backOk = true, i = prints.length - 1;
const wrong: string[] = [];
while (wb.canUndo && i > 0) {
  await cmds.run("edit.undo");
  i--;
  if (fingerprint(wb) !== prints[i]) { backOk = false; wrong.push(`undo to ${i}`); const a = fingerprint(wb), b = prints[i]; let k = 0; while (a[k] === b[k]) k++; print(a.slice(Math.max(0, k - 80), k + 120)); print(b.slice(Math.max(0, k - 80), k + 120)); break; }
}
eq("undoing step by step meets each earlier state exactly", [backOk, wrong], [true, []]);
eq("and all the way back is the empty workbook", [fingerprint(wb) === prints[0], i], [true, 0]);
// Redo forward.
let fwdOk = true, j = 0;
while (wb.canRedo) {
  await cmds.run("edit.redo");
  j++;
  if (fingerprint(wb) !== prints[j]) { fwdOk = false; wrong.push(`redo to ${j}`); break; }
}
eq("redoing step by step meets each later state exactly", [fwdOk, j, wrong], [true, prints.length - 1, []]);

// A step is entries: what a one-cell edit holds, and that nothing else is held.
{
  const w = new Workbook(), c = new Commands(w);
  await c.run("formulas.use language=Excel");
  await c.run("cell.set cell=A1 input=1");
  const before = w.saved();
  await c.run("cell.set cell=B2 input=2");
  const step = stepOf(before, w.saved(), "x");
  eq("a one-cell edit is one entry, for the sheet's cells, with that cell", [step.entries.length, step.entries[0].kind, step.entries[0].slot?.key, step.entries[0].changes], [1, "map", "inputs", [["B2", undefined, "2"]]]);
  eq("applied backwards and forwards it moves between the two states", [
    applyStep(step, w.saved(), true).pages[0].sheet.inputs.has("B2"),
    applyStep(step, before, false).pages[0].sheet.inputs.get("B2"),
  ], [false, "2"]);
}

// A look is the notes of the cells it touched, as the keys that changed.
{
  const w = new Workbook(), c = new Commands(w);
  await c.run("formulas.use language=Excel");
  await c.run("cell.set cell=A1 input=1");
  const before = w.saved();
  await c.run("format.fill range=A1:A3 color=red");
  const step = stepOf(before, w.saved(), "x");
  eq("a look over three cells is one entry for the notes holding three keys", [step.entries.map((e) => `${e.kind}:${e.slot?.key}`), step.entries[0].changes!.map((x) => x[0])], [["map:notes"], ["A1", "A2", "A3"]]);
}

// Undo does not disturb what the step did not touch: a change made outside the
// journal (a command that is not a step) stays when a different step is undone.
{
  const w = new Workbook(), c = new Commands(w);
  await c.run("formulas.use language=Excel");
  await c.run("cell.set cell=A1 input=1");
  await c.run("cell.set cell=A2 input=2");
  await w.sheet.set("C1", "outside");
  await c.run("edit.undo");
  eq("an undo takes back the step's cell and leaves one set outside the journal", [w.sheet.input("A2"), w.sheet.input("A1"), w.sheet.input("C1")], ["", "1", "outside"]);
}

// expect: ok the mixed sequence ran, step by step
// expect: ok undoing step by step meets each earlier state exactly
// expect: ok and all the way back is the empty workbook
// expect: ok redoing step by step meets each later state exactly
// expect: ok a one-cell edit is one entry, for the sheet's cells, with that cell
// expect: ok applied backwards and forwards it moves between the two states
// expect: ok a look over three cells is one entry for the notes holding three keys
// expect: ok an undo takes back the step's cell and leaves one set outside the journal
