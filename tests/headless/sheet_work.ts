// Calculation work counters (docs/comparisons/HYPERFORMULA.md action 1): a recalculation is
// judged by what it did, not how long it took on this machine: formulas
// evaluated, cells considered, passes and notifications. Each shape of
// workbook below says how much work an edit must cost, and an unaffected
// branch must cost none.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const make = async () => {
  const wb = new Workbook(), cmds = new Commands(wb);
  await cmds.run("formulas.use language=Excel");
  return { wb, cmds };
};
const work = async (wb: Workbook, f: () => Promise<unknown>) => {
  wb.resetWork();
  await f();
  return wb.work();
};

// A chain of 200 formulas beside an unrelated branch of 500.
{
  const { wb, cmds } = await make();
  const e: [string, string][] = [["A1", "1"]];
  for (let r = 2; r <= 200; r++) e.push([`A${r}`, `=A${r - 1}+1`]);
  e.push(["C1", "1"]);
  for (let r = 1; r <= 500; r++) e.push([`D${r}`, `=C1*${r}`]);
  await wb.sheet.setMany(e);
  const w = await work(wb, () => cmds.run("cell.set cell=A1 input=5"));
  eq("a chain of 200: an edit at its head evaluates each formula once", [w.evaluated, wb.sheet.value("A200")], [199, 204]);
  const none = await work(wb, () => cmds.run("cell.set cell=A1 input=5"));
  eq("an unrelated branch of 500 is not evaluated by it (all 199 were the chain's)", [none.evaluated <= 199, wb.sheet.value("D500")], [true, 500]);
  const tail = await work(wb, () => cmds.run("cell.set cell=C1 input=2"));
  eq("and an edit to the branch evaluates it and not the chain", [tail.evaluated, wb.sheet.value("D500")], [500, 1000]);
  eq("one change is one pass", [tail.passes], [1]);
  eq("and one notification per changed cell", tail.notified, 501);
}

// A diamond: four formulas, each once.
{
  const { wb, cmds } = await make();
  await wb.sheet.setMany([["A1", "1"], ["B1", "=A1+1"], ["C1", "=A1+2"], ["D1", "=B1+C1"]]);
  const w = await work(wb, () => cmds.run("cell.set cell=A1 input=10"));
  eq("a diamond: the join is evaluated once, after both sides", [w.evaluated, wb.sheet.value("D1")], [3, 23]);
}

// A running total: an edit halfway evaluates the rest of the total, not the part before.
{
  const { wb, cmds } = await make();
  const e: [string, string][] = [];
  for (let r = 1; r <= 1000; r++) { e.push([`A${r}`, "1"]); e.push([`B${r}`, r === 1 ? "=A1" : `=B${r - 1}+A${r}`]); }
  await wb.sheet.setMany(e);
  const w = await work(wb, () => cmds.run("cell.set cell=A600 input=11"));
  eq("a running total of 1,000: an edit at row 600 evaluates rows 600 to 1000", [w.evaluated, wb.sheet.value("B1000")], [401, 1010]);
}

// A volatile formula is evaluated on every change, and what reads it with it.
{
  const { wb, cmds } = await make();
  await wb.sheet.setMany([["A1", "1"], ["A2", "2"], ["B1", "=OFFSET(A1,1,0)"], ["B2", "=B1*2"], ["C1", "5"], ["C2", "=C1+1"]]);
  const w = await work(wb, () => cmds.run("cell.set cell=C1 input=6"));
  eq("a volatile OFFSET and its reader are evaluated on an edit that touches neither", [w.evaluated, wb.sheet.value("B2"), wb.sheet.value("C2")], [3, 4, 7]);
}

// A sum over a range is evaluated once per edit, not once per cell of the range.
{
  const { wb, cmds } = await make();
  const e: [string, string][] = [];
  for (let r = 1; r <= 300; r++) e.push([`A${r}`, String(r)]);
  e.push(["B1", "=SUM(A1:A300)"]);
  await wb.sheet.setMany(e);
  const w = await work(wb, () => cmds.run("cell.set cell=A150 input=0"));
  eq("SUM over 300 cells: one evaluation per edit", [w.evaluated, wb.sheet.value("B1")], [1, 45150 - 150]);
}

// A cumulative total over a growing range reads every cell of every range (the
// quadratic shape that range-aware aggregates would cut), and says so.
{
  const { wb, cmds } = await make();
  const e: [string, string][] = [];
  for (let r = 1; r <= 200; r++) { e.push([`A${r}`, "1"]); e.push([`B${r}`, `=SUM($A$1:A${r})`]); }
  await wb.sheet.setMany(e);
  const w = await work(wb, () => cmds.run("cell.set cell=A1 input=2"));
  eq("SUM($A$1:An) down 200 rows: 200 evaluations, each reading its whole range", [w.evaluated, w.reads >= 200 * 201 / 2, w.reads < 3 * 200 * 201 / 2, wb.sheet.value("B200")], [200, true, true, 201]);
}

// Reading counters is free of effect: they reset.
{
  const { wb } = await make();
  wb.resetWork();
  eq("counters reset to zero", Object.values(wb.work()), [0, 0, 0, 0, 0, 0]);
}

// expect: ok a chain of 200: an edit at its head evaluates each formula once
// expect: ok an unrelated branch of 500 is not evaluated by it (all 199 were the chain's)
// expect: ok and an edit to the branch evaluates it and not the chain
// expect: ok one change is one pass
// expect: ok and one notification per changed cell
// expect: ok a diamond: the join is evaluated once, after both sides
// expect: ok a running total of 1,000: an edit at row 600 evaluates rows 600 to 1000
// expect: ok a volatile OFFSET and its reader are evaluated on an edit that touches neither
// expect: ok SUM over 300 cells: one evaluation per edit
// expect: ok SUM($A$1:An) down 200 rows: 200 evaluations, each reading its whole range
// expect: ok counters reset to zero
