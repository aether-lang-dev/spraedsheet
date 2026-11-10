// An interrupted recalculation (sae's page time limit killing the promise job
// a formula was awaited in, so it never comes back and no finally runs) must
// not leave the sheet half worked out. Turns.interrupt fails the turn so the
// queue moves on (sheet_turns.ts); here the cells that turn was working out
// are worked out by the next turn, and said to the listener then, as if the
// change had finished.
import { Sheet, Turns } from "./sheet/sheet.ts";
import type { Context } from "./sheet/sheet.ts";
import { EXCEL } from "./sheet/excel.ts";
import type { Language } from "./sheet/language.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const settle = async () => { for (let i = 0; i < 200; i++) await Promise.resolve(); };

// Excel, but a formula whose text is "A1*10" waits for ever while `stuck`:
// the job the time limit killed.
let stuck = false;
const gated: Language = {
  ...EXCEL,
  compile: (...a) => {
    const f = EXCEL.compile(...a);
    return f.kind === "formula" && a[0] === "A1*10"
      ? { ...f, run: (env) => (stuck ? new Promise<never>(() => {}) : f.run(env)) }
      : f;
  },
};

const heard: string[] = [];
const turns = new Turns();
const context: Context = { name: "SHEET1", turns, value: () => null };
const s = new Sheet({ changed: (k) => heard.push(k) }, gated, context);
const shown = (...keys: string[]) => keys.map((k) => show(s.value(k)));

await s.set("A1", "1");
await s.set("B1", "=A1*10");
await s.set("C1", "=B1+1");
eq("worked out", shown("B1", "C1"), ["10", "11"]);

stuck = true;
const edit = s.set("A1", "2").then(() => "done", (e) => `failed ${e.message}`);
await settle();
eq("the edit's recalculation waits on the formula that never comes back", [turns.busy, shown("B1", "C1")], [true, ["10", "11"]]);
turns.interrupt("page time limit");
eq("interrupting fails that edit", await edit, "failed interrupted: page time limit");

stuck = false;
heard.length = 0;
await s.set("D1", "5");
eq("the next edit, of another cell, works out what the interrupted one left", shown("A1", "B1", "C1", "D1"), ["2", "20", "21", "5"]);
eq("and the listener hears those cells then", [...heard].sort(), ["A1", "B1", "C1", "D1"]);

heard.length = 0;
await s.set("D1", "6");
eq("once worked out they are not worked out again", heard, ["D1"]);

// A recalculation that finished leaves nothing behind.
await s.set("A1", "3");
eq("a change after it still flows through", shown("B1", "C1"), ["30", "31"]);

// A cell too heavy for the limit is tried once more, then left stale, so the
// changes after it are not held up by it for good.
{
  stuck = false;
  await s.set("A1", "1");
  heard.length = 0;
  stuck = true;
  const first = s.set("A1", "2").then(() => "done", (e) => `failed ${e.message}`);
  await settle();
  turns.interrupt("page time limit");
  await first;
  const second = s.set("D1", "7").then(() => "done", (e) => `failed ${e.message}`);
  await settle();
  eq("a retry that is stopped too is stopped", [turns.busy, turns.interrupt("page time limit")], [true, true]);
  await second;
  stuck = false;
  eq("nothing is stale before it gives up", s.stale(), []);
  await s.set("E1", "8");
  eq("the third change is not held up: its own cells, and D1, are worked out", shown("D1", "E1"), ["7", "8"]);
  eq("the cells it could not finish are left stale", s.stale(), ["A1", "B1", "C1"]);
  await s.recalculate();
  eq("Calculate works them out and clears them", [shown("A1", "B1", "C1"), s.stale()], [["2", "20", "21"], []]);
}

// expect: ok worked out
// expect: ok the edit's recalculation waits on the formula that never comes back
// expect: ok interrupting fails that edit
// expect: ok the next edit, of another cell, works out what the interrupted one left
// expect: ok and the listener hears those cells then
// expect: ok once worked out they are not worked out again
// expect: ok a change after it still flows through
// expect: ok a retry that is stopped too is stopped
// expect: ok nothing is stale before it gives up
// expect: ok the third change is not held up: its own cells, and D1, are worked out
// expect: ok the cells it could not finish are left stale
// expect: ok Calculate works them out and clears them
