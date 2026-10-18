// Malformed input and random sequences (7.6): every command, run with its
// required parameters missing, then each given what it cannot be (words
// for a number or a cell, a cell off the sheet, the empty string), then
// one 20,000 characters long, refuses cleanly: an outcome said, nothing
// thrown. Then seeded random sessions of the commands a person uses most,
// each taken back with undo all the way, end where they began: the same
// workbook, file for file.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { writeSheetFile } from "./sheet/file.ts";
import { eq } from "./lib/check.ts";

// Commands left out: none touches the system, but these wait on it (a
// script's setup, a picture's bytes) or are the session itself.
const SKIP = new Set(["setup.run"]);
const wb = new Workbook();
const cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
await cmds.run("cell.set cell=A1 input=1");
await cmds.run("cell.set cell=B2 input=two");
const crashed: string[] = [], unsaid: string[] = [];
const tryIt = async (id: string, p: Record<string, string>, what: string) => {
  try {
    const o = await cmds.execute(id, p);
    if (typeof o?.message !== "string" || (!o.ok && o.message.trim() === "")) unsaid.push(`${id} ${what}`);
  } catch (e) {
    crashed.push(`${id} ${what}: ${String((e as Error)?.message ?? e).slice(0, 120)}`);
  }
};
const catalogue = cmds.catalogue().filter((c) => !SKIP.has(c.id));
for (const c of catalogue) {
  const required = c.params.split(", ").filter((x) => x && !x.endsWith("?"));
  if (required.length) {
    const o = await cmds.execute(c.id, {});
    if (o.ok || !o.message.startsWith(`${c.id} needs `)) unsaid.push(`${c.id} with nothing`);
  }
  const all = c.params.split(", ").filter((x) => x).map((x) => x.replace(/\?$/, ""));
  for (const bad of ["@@@", "", "ZZZZ99999999", "-1", "1e400"]) {
    for (const name of all) await tryIt(c.id, { ...Object.fromEntries(required.map((r) => [r, "A1"])), [name]: bad }, `${name}=${JSON.stringify(bad)}`);
  }
  const huge = "A".repeat(20000);
  for (const name of all) await tryIt(c.id, { ...Object.fromEntries(required.map((r) => [r, "A1"])), [name]: huge }, `${name}=<20000 chars>`);
}
eq("every command, missing, wrong and huge parameters: none throws", crashed, []);
eq("and each says why it refuses", unsaid, []);

// A seeded generator (mulberry32).
function random(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
const steps = (rnd: () => number): string => {
  const pick = <T>(xs: readonly T[]): T => xs[Math.floor(rnd() * xs.length)];
  const cell = () => `${pick(["A", "B", "C", "D", "E", "F"])}${1 + Math.floor(rnd() * 12)}`;
  const span = () => { const a = cell(), b = cell(); return `${a}:${b}`; };
  return pick([
    () => `cell.set cell=${cell()} input=${pick([String(Math.floor(rnd() * 100)), "pears", `"=SUM(${span()})"`, `"=${cell()}*2"`, `"=IF(${cell()}>5,""big"",""small"")"`])}`,
    () => `range.clear range=${span()}`,
    () => `format.font range=${span()} bold=1`,
    () => `format.fill range=${span()} color=yellow`,
    () => `rows.insert at=${1 + Math.floor(rnd() * 10)}`,
    () => `rows.delete at=${1 + Math.floor(rnd() * 10)}`,
    () => `columns.insert at=${pick(["B", "C", "D"])}`,
    () => `columns.delete at=${pick(["B", "C", "D"])}`,
    () => `cells.merge range=${span()}`,
    () => `cells.unmerge range=${span()}`,
    () => `data.sortAscending range=A1:F12 by=${pick(["A", "B", "C"])}`,
    () => `edit.copy range=${span()}`,
    () => `edit.paste at=${cell()}`,
    () => `edit.fillDown range=${pick(["A1:A6", "B2:B9", "C1:C4"])}`,
    () => `cell.comment cell=${cell()} text=noted`,
    () => `names.define name=N${Math.floor(rnd() * 5)} range=${span()}`,
    () => `sheet.add name=S${Math.floor(rnd() * 5)}`,
    () => `format.number range=${span()} format=currency`,
    () => `rows.hide range=${1 + Math.floor(rnd() * 10)}:${1 + Math.floor(rnd() * 10)}`,
    () => `view.freezePanes cell=${cell()}`,
  ])();
};
const results: string[] = [];
for (const seed of [1, 2, 3, 4, 5]) {
  const w = new Workbook(), c = new Commands(w), rnd = random(seed);
  await c.run("formulas.use language=Excel");
  for (let i = 1; i <= 6; i++) await c.run(`cell.set cell=A${i} input=${i * 3}`);
  while (w.canUndo) await c.run("edit.undo");
  const start = JSON.stringify(writeSheetFile(w));
  let ran = 0, threw = "";
  for (let i = 0; i < 80; i++) {
    try { if ((await c.run(steps(rnd))).ok) ran++; } catch (e) { threw = String((e as Error)?.message ?? e); break; }
  }
  let undone = 0;
  while (w.canUndo && undone < 200) { await c.run("edit.undo"); undone++; }
  const same = JSON.stringify(writeSheetFile(w)) === start;
  results.push(`${seed}: ${threw ? `threw ${threw}` : same ? "back where it began" : `differs after ${ran} steps`}`);
}
eq("seeded random sessions, undone all the way, end where they began", results, ["1: back where it began", "2: back where it began", "3: back where it began", "4: back where it began", "5: back where it began"]);

// expect: ok every command, missing, wrong and huge parameters: none throws
// expect: ok and each says why it refuses
// expect: ok seeded random sessions, undone all the way, end where they began
