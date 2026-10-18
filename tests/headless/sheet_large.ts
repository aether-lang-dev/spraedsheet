// Large files (7.8): a sheet of 100,000 cells (50,000 numbers and 50,000
// formulas over them) opens, recalculates and scrolls within stated times;
// past what a sheet holds here (MOST_HELD, 200,000 cells, what the
// installed app's 512 MB page works with) a change is refused whole, with a
// message, before anything is set. Run with the app's heap (run-tests.sh's
// headless entry: SAE_PAGE_MEM, SAE_TIME_LIMIT). The stated times:
const TIMES = { build: 8000, edit: 1000, recalculate: 6000, save: 4000, open: 8000, scroll: 300 };
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
const ms = async (f: () => Promise<unknown>) => { const t = Date.now(); await f(); return Date.now() - t; };
const within = (what: keyof typeof TIMES, t: number) => (t <= TIMES[what] ? `${what} within ${TIMES[what]} ms` : `${what} took ${t} ms, over ${TIMES[what]}`);

const N = 50_000;
const entries: [string, string][] = [];
for (let r = 1; r <= N; r++) entries.push([`A${r}`, String(r)], [`B${r}`, `=A${r}*2`]);
const tBuild = await ms(() => wb.sheet.setMany(entries));
eq("100,000 cells set and worked out", [wb.sheet.keys().length, show(wb.sheet.value("B50000")), within("build", tBuild)], [100_000, "100000", "build within 8000 ms"]);

const tEdit = await ms(() => cmds.run("cell.set cell=A25000 input=1"));
eq("an edit ripples", [show(wb.sheet.value("B25000")), within("edit", tEdit)], ["2", "edit within 1000 ms"]);

const tCalc = await ms(() => cmds.run("formulas.recalculate"));
eq("the whole sheet recalculated", within("recalculate", tCalc), "recalculate within 6000 ms");

const tSave = await ms(() => cmds.run("file.save name=Large"));
eq("saved", [data.has("sheet.file.Large"), within("save", tSave)], [true, "save within 4000 ms"]);

await cmds.run("sheet.clear");
const tOpen = await ms(() => cmds.run("file.open name=Large"));
eq("opened again, worked out", [wb.sheet.keys().length, show(wb.sheet.value("B49999")), within("open", tOpen)], [100_000, "99998", "open within 8000 ms"]);

// Scrolling: what the virtualised grid asks of the sheet for a window far
// down (20 rows by 10 columns at row 40,000), its text and its look.
const tScroll = await ms(async () => {
  for (let r = 40_000; r < 40_020; r++) for (const c of "ABCDEFGHIJ") { wb.textOf(`${c}${r}`); wb.styleOf(`${c}${r}`); }
});
eq("a window far down drawn", [wb.textOf("B40000"), within("scroll", tScroll)], ["80000", "scroll within 300 ms"]);

// Past what a sheet holds: refused whole, nothing set.
const more: [string, string][] = [];
for (let r = 1; r <= MOST_HELD - 100_000 + 1; r++) more.push([`D${r}`, "1"]);
let refused = "";
try { await wb.sheet.setMany(more); } catch (e) { refused = String((e as Error).message); }
eq("a change past it refused, with why, and nothing set", [refused, wb.sheet.keys().length, show(wb.sheet.value("D1"))],
  ["that would leave the sheet holding 200,001 cells: the most a sheet holds here is 200,000", 100_000, ""]);
eq("and through a command, said", (await cmds.execute("cell.set", { cell: "D1", input: "1" })).ok, true);

// expect: ok 100,000 cells set and worked out
// expect: ok an edit ripples
// expect: ok the whole sheet recalculated
// expect: ok saved
// expect: ok opened again, worked out
// expect: ok a window far down drawn
// expect: ok a change past it refused, with why, and nothing set
// expect: ok and through a command, said
