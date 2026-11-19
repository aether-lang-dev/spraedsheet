// Large files (7.8): a sheet of 100,000 cells (50,000 numbers and 50,000
// formulas over them) opens, recalculates and scrolls within stated times;
// past what a sheet holds here (MOST_HELD, 200,000 cells, what the
// installed app's 512 MB page works with) a change is refused whole, with a
// message, before anything is set. Run with the app's heap (run-tests.sh's
// headless entry: SAE_PAGE_MEM, SAE_TIME_LIMIT). The stated times, in ms, are what the development
// Mac needs with room to spare (it measures 2.2 s to save and 2.5 s to open). A slower host is
// allowed as much longer as it is slower: a fixed workload that is none of this app's code (a
// map, strings and a sort, in the same engine) is timed on that host first, against what the Mac
// took, and the stated times are multiplied by the ratio. Never by less than 1: a fast host is
// held to them as stated. The workload is not the code under test, so a slowdown in the code
// (which the workload does not feel) fails on any host.
const STATED = { build: 8000, edit: 1000, recalculate: 6000, save: 4000, open: 8000, scroll: 300 };
const CALIBRATION_MS = 96;
import { Commands } from "./sheet/commands.ts";
import type { Store } from "./sheet/commands.ts";
import { MOST_HELD } from "./sheet/sheet.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const data = new Map<string, string>();
const store: Store = { get: (k) => data.get(k) ?? null, set: (k, v) => { data.set(k, v); }, remove: (k) => { data.delete(k); } };
const wb = new Workbook();
const cmds = new Commands(wb, store);
await cmds.run("formulas.use language=Excel");
const calibrate = (): number => {
  let best = Infinity;
  for (let round = 0; round < 3; round++) {
    const t = Date.now();
    const m = new Map<string, number>();
    for (let k = 0; k < 300_000; k++) m.set(`k${k}`, k * 2);
    let total = 0;
    for (const [key, v] of m) total += key.length + v;
    const xs = Array.from({ length: 150_000 }, (_, k) => (k * 7919) % 150_001);
    xs.sort((a, b) => a - b);
    total += xs.slice(0, 60_000).join(",").split(",").length;
    if (total > 0) best = Math.min(best, Date.now() - t);
  }
  return best;
};
const calibration = calibrate();
const slower = Math.max(1, calibration / CALIBRATION_MS);
const TIMES = Object.fromEntries(Object.entries(STATED).map(([k, v]) => [k, Math.round(v * slower)])) as typeof STATED;
const ms = async (f: () => Promise<unknown>) => { const t = Date.now(); await f(); return Date.now() - t; };
const within = (what: keyof typeof TIMES, t: number) => (t <= TIMES[what] ? `${what} within its time` : `${what} took ${t} ms, over ${TIMES[what]} (stated ${STATED[what]}, host x${slower.toFixed(2)})`);

eq("the host timed against the Mac", calibration > 0 && CALIBRATION_MS > 0, true);

const N = 50_000;
const entries: [string, string][] = [];
for (let r = 1; r <= N; r++) entries.push([`A${r}`, String(r)], [`B${r}`, `=A${r}*2`]);
const tBuild = await ms(() => wb.sheet.setMany(entries));
eq("100,000 cells set and worked out", [wb.sheet.keys().length, show(wb.sheet.value("B50000")), within("build", tBuild)], [100_000, "100000", "build within its time"]);

const tEdit = await ms(() => cmds.run("cell.set cell=A25000 input=1"));
eq("an edit ripples", [show(wb.sheet.value("B25000")), within("edit", tEdit)], ["2", "edit within its time"]);

const tCalc = await ms(() => cmds.run("formulas.recalculate"));
eq("the whole sheet recalculated", within("recalculate", tCalc), "recalculate within its time");

const tSave = await ms(() => cmds.run("file.save name=Large"));
eq("saved", [data.has("sheet.file.Large"), within("save", tSave)], [true, "save within its time"]);

await cmds.run("sheet.clear");
const tOpen = await ms(() => cmds.run("file.open name=Large"));
eq("opened again, worked out", [wb.sheet.keys().length, show(wb.sheet.value("B49999")), within("open", tOpen)], [100_000, "99998", "open within its time"]);

// Scrolling: what the virtualised grid asks of the sheet for a window far
// down (20 rows by 10 columns at row 40,000), its text and its look.
const tScroll = await ms(async () => {
  for (let r = 40_000; r < 40_020; r++) for (const c of "ABCDEFGHIJ") { wb.textOf(`${c}${r}`); wb.styleOf(`${c}${r}`); }
});
eq("a window far down drawn", [wb.textOf("B40000"), within("scroll", tScroll)], ["80000", "scroll within its time"]);

// Past what a sheet holds: refused whole, nothing set.
const more: [string, string][] = [];
for (let r = 1; r <= MOST_HELD - 100_000 + 1; r++) more.push([`D${r}`, "1"]);
let refused = "";
try { await wb.sheet.setMany(more); } catch (e) { refused = String((e as Error).message); }
eq("a change past it refused, with why, and nothing set", [refused, wb.sheet.keys().length, show(wb.sheet.value("D1"))],
  ["that would leave the sheet holding 200,001 cells: the most a sheet holds here is 200,000", 100_000, ""]);
eq("and through a command, said", (await cmds.execute("cell.set", { cell: "D1", input: "1" })).ok, true);

// expect: ok the host timed against the Mac
// expect: ok 100,000 cells set and worked out
// expect: ok an edit ripples
// expect: ok the whole sheet recalculated
// expect: ok saved
// expect: ok opened again, worked out
// expect: ok a window far down drawn
// expect: ok a change past it refused, with why, and nothing set
// expect: ok and through a command, said
