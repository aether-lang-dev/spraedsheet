// sheet-server (server/sheetserver.ae), headless: the real binary, real files
// under target/server-tests/, real HTTP. Each test starts its own server on
// its own data folder. node --test tests/server/server.test.js
//
//   SHEET_SERVER   the binary (default target/sheet-server)
//   SHEET_TEST_PORT  the first port to use (default 18393; each test takes the next)
import { test } from "node:test";
import assert from "node:assert/strict";
import { spawn, spawnSync } from "node:child_process";
import { appendFileSync, existsSync, mkdirSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { applyChanges, changesBetween, keyedLines, textOfLines } from "../../app/sheet/teamtext.ts";

const ROOT = join(dirname(fileURLToPath(import.meta.url)), "..", "..");
const BIN = process.env.SHEET_SERVER ?? join(ROOT, "target", "sheet-server");
const WORK = join(ROOT, "target", "server-tests");
let nextPort = Number(process.env.SHEET_TEST_PORT ?? 18393);
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// A data folder with two people; their tokens.
function people(name) {
  const data = join(WORK, name);
  rmSync(data, { recursive: true, force: true });
  mkdirSync(data, { recursive: true });
  const token = (who) => spawnSync(BIN, ["add-person", data, who], { encoding: "utf8" }).stdout.trim();
  return { data, ann: token("ann"), bob: token("bob") };
}

// The server on `data`, until stop(); its console in .out.
async function start(data, port = nextPort++) {
  const proc = spawn(BIN, ["serve", data, join(ROOT, "app"), String(port)], { stdio: ["ignore", "pipe", "pipe"] });
  const s = { proc, port, out: "", url: `http://127.0.0.1:${port}` };
  proc.stdout.on("data", (d) => { s.out += d; });
  proc.stderr.on("data", (d) => { s.out += d; });
  for (let i = 0; i < 100; i++) {
    try { if ((await fetch(`${s.url}/api/info`)).ok) return s; } catch { /* not up yet */ }
    await sleep(50);
  }
  proc.kill("SIGKILL");
  throw new Error(`sheet-server did not start: ${s.out}`);
}
// A child that has ended (on its own, or by a signal: exitCode stays null then).
const gone_ = (proc) => proc.exitCode !== null || proc.signalCode !== null;
async function stop(s, signal = "SIGTERM") {
  if (gone_(s.proc)) return;
  const gone = new Promise((r) => s.proc.once("exit", r));
  s.proc.kill(signal);
  await gone;
}

const call = async (s, token, method, path, body) => {
  const res = await fetch(`${s.url}${path}`, { method, headers: { Authorization: `Bearer ${token}` }, body: body === undefined ? undefined : typeof body === "string" ? body : JSON.stringify(body) });
  const text = await res.text();
  let json = null;
  try { json = JSON.parse(text); } catch { /* text */ }
  return { status: res.status, json, text };
};

// The workbook every test starts from: Sheet1 with A1..C3, Costs with B2.
const START = "# rexxsheet text 1\n@formulas\t\"excel\"\n@sheet\tSheet1\nA1\tItem\nB1\t5\nC1\t=B1*2\nA2\tTea\nB2\t3\n@names\tQTY\t\"B1:B2\"\n@sheet\tCosts\nB2\t7\n";
const asTriples = (changes) => changes.map((c) => [c.key, c.was, c.now]);
let ids = 0;
const op = (base, cmd, changes, extra = {}) => ({ id: `t${process.pid}-${++ids}`, base, cmd, structure: [], changes: asTriples(changes), ...extra });
// One cell's change against the lines as they are.
const cellChange = (lines, sheet, cell, input) => {
  const key = `${sheet}\t${cell}`;
  return { key, was: lines.get(key) ?? "", now: input === "" ? "" : `${cell}\t${input}` };
};
async function create(s, token, name, text = START) {
  return call(s, token, "POST", `/api/books/${name}`, op(0, "create", changesBetween(new Map(), keyedLines(text))));
}

test("a workbook made, read back as the text it was made from", async () => {
  const p = people("create");
  const s = await start(p.data);
  try {
    const c = await create(s, p.ann, "budget");
    assert.equal(c.status, 200, c.text);
    assert.equal(c.json.rev, 1);
    const g = await call(s, p.ann, "GET", "/api/books/budget");
    assert.equal(g.json.rev, 1);
    assert.equal(g.json.text, textOfLines(keyedLines(START)));
    assert.equal((await create(s, p.bob, "budget")).status, 409, "a second workbook of the same name is refused");
    assert.equal((await call(s, p.ann, "GET", "/api/books")).text, '[{"name":"budget","rev":1}]');
    assert.equal(readFileSync(join(p.data, "books", "budget", "head.sheet.txt"), "utf8"), textOfLines(keyedLines(START), 1));
  } finally { await stop(s); }
});

test("no token, or one the server does not know: 401, nothing done", async () => {
  const p = people("auth");
  const s = await start(p.data);
  try {
    assert.equal((await call(s, "nobody", "GET", "/api/books")).status, 401);
    assert.equal((await create(s, "nobody", "x")).status, 401);
    assert.equal(existsSync(join(p.data, "books", "x")), false);
    assert.equal((await call(s, p.bob, "GET", "/api/me")).json.name, "bob");
  } finally { await stop(s); }
});

test("two edits from one revision: different cells both kept, the same cell refused, saying whose change won", async () => {
  const p = people("cas");
  const s = await start(p.data);
  try {
    await create(s, p.ann, "b");
    const lines = keyedLines(START);
    const a = await call(s, p.ann, "POST", "/api/books/b/ops", op(1, "cell.set cell=B1 input=6", [cellChange(lines, "Sheet1", "B1", "6")]));
    const b = await call(s, p.bob, "POST", "/api/books/b/ops", op(1, "cell.set cell=B2 input=4", [cellChange(lines, "Sheet1", "B2", "4")]));
    assert.deepEqual([a.status, a.json.rev, b.status, b.json.rev], [200, 2, 200, 3], "disjoint cells from revision 1: both kept, in order");
    const c = await call(s, p.bob, "POST", "/api/books/b/ops", op(1, "cell.set cell=B1 input=9", [cellChange(lines, "Sheet1", "B1", "9")]));
    assert.equal(c.status, 409);
    assert.deepEqual({ ...c.json.conflict, at: typeof c.json.conflict.at }, { key: "Sheet1\tB1", why: "changed", rev: 2, who: "ann", at: "number", cmd: "cell.set cell=B1 input=6" });
    const text = (await call(s, p.ann, "GET", "/api/books/b")).json.text;
    assert.match(text, /\nB1\t6\n/);
    assert.match(text, /\nB2\t4\n/);
    assert.equal((await call(s, p.ann, "GET", "/api/books/b")).json.rev, 3, "the refused op took no revision");
  } finally { await stop(s); }
});

test("a structural change and any other change to the same sheet from one revision conflict; another sheet is not touched", async () => {
  const p = people("shape");
  const s = await start(p.data);
  try {
    await create(s, p.ann, "b");
    const lines = keyedLines(START);
    // Ann inserts a row above row 2 on Sheet1 (A2/B2 move to row 3).
    const moved = new Map(lines);
    applyChanges(moved, [cellChange(lines, "Sheet1", "A2", ""), cellChange(lines, "Sheet1", "B2", ""), { key: "Sheet1\tA3", was: "", now: "A3\tTea" }, { key: "Sheet1\tB3", was: "", now: "B3\t3" }]);
    const ins = await call(s, p.ann, "POST", "/api/books/b/ops", op(1, "rows.insert at=2", changesBetween(lines, moved), { structure: ["Sheet1"] }));
    assert.equal(ins.status, 200, ins.text);
    // Bob, from revision 1, types into C9: an empty cell then and now, but the rows under it moved.
    const late = await call(s, p.bob, "POST", "/api/books/b/ops", op(1, "cell.set cell=C9 input=1", [cellChange(lines, "Sheet1", "C9", "1")]));
    assert.equal(late.status, 409);
    assert.deepEqual([late.json.conflict.why, late.json.conflict.who, late.json.conflict.cmd], ["reshaped", "ann", "rows.insert at=2"]);
    const other = await call(s, p.bob, "POST", "/api/books/b/ops", op(1, "cell.set cell=Costs!B2 input=8", [cellChange(lines, "Costs", "B2", "8")]));
    assert.equal(other.status, 200, "another sheet's cell is not in the way");
    // And the other way round: a sort from revision 2 after Bob's change to Costs (revision 3) is refused there.
    const sort = await call(s, p.ann, "POST", "/api/books/b/ops", op(2, "data.sort range=B2:B2", [{ key: "Costs\tB2", was: "B2\t8", now: "B2\t8 " }], { structure: ["Costs"] }));
    assert.equal(sort.status, 409);
    assert.equal(sort.json.conflict.why, "edited");
  } finally { await stop(s); }
});

test("an op sent twice (its answer lost) is kept once, and answered with its revision", async () => {
  const p = people("again");
  const s = await start(p.data);
  try {
    await create(s, p.ann, "b");
    const o = op(1, "cell.set cell=B1 input=6", [cellChange(keyedLines(START), "Sheet1", "B1", "6")]);
    const first = await call(s, p.ann, "POST", "/api/books/b/ops", o);
    const second = await call(s, p.ann, "POST", "/api/books/b/ops", o);
    assert.deepEqual([first.status, first.json.rev, second.status, second.json.rev, second.json.again], [200, 2, 200, 2, true]);
    assert.equal(readFileSync(join(p.data, "books", "b", "ops.log"), "utf8").trim().split("\n").length, 2);
  } finally { await stop(s); }
});

test("a long-poll is answered as soon as another person's op is kept, with who is here", async () => {
  const p = people("poll");
  const s = await start(p.data);
  try {
    await create(s, p.ann, "b");
    const t0 = Date.now();
    const waiting = call(s, p.bob, "GET", "/api/books/b/ops?since=1&wait=8&cell=Sheet1%09A2");
    await sleep(300);
    const sent = Date.now();
    await call(s, p.ann, "POST", "/api/books/b/ops", op(1, "cell.set cell=B1 input=6", [cellChange(keyedLines(START), "Sheet1", "B1", "6")]));
    const got = await waiting;
    const lag = Date.now() - sent;
    assert.equal(got.json.rev, 2);
    assert.equal(got.json.ops.length, 1);
    assert.equal(got.json.ops[0].who, "ann");
    assert.ok(lag < 400, `answered ${lag} ms after the op (a poll every 100 ms)`);
    assert.ok(Date.now() - t0 < 2000);
    assert.deepEqual(got.json.people, [{ name: "bob", cell: "Sheet1\tA2" }]);
    // Nothing new: held for the wait, then answered with no ops.
    const t1 = Date.now();
    const idle = await call(s, p.bob, "GET", "/api/books/b/ops?since=2&wait=1");
    assert.deepEqual([idle.status, idle.json.ops.length], [200, 0]);
    assert.ok(Date.now() - t1 >= 900, "held for about its wait");
    console.log(`long-poll: answered ${lag} ms after the op was kept`);
  } finally { await stop(s); }
});

test("history: every op that changed a cell; the text at any revision", async () => {
  const p = people("history");
  const s = await start(p.data);
  try {
    await create(s, p.ann, "b");
    let lines = keyedLines(START);
    for (const [who, v] of [[p.ann, "6"], [p.bob, "7"], [p.ann, "8"]]) {
      const ch = [cellChange(lines, "Sheet1", "B1", v)];
      const rev = (await call(s, p.ann, "GET", "/api/books/b")).json.rev;
      assert.equal((await call(s, who, "POST", "/api/books/b/ops", op(rev, `cell.set cell=B1 input=${v}`, ch))).status, 200);
      applyChanges(lines, ch);
    }
    const h = await call(s, p.ann, "GET", `/api/books/b/history?key=${encodeURIComponent("Sheet1\tB1")}`);
    assert.deepEqual(h.json.map((e) => [e.rev, e.who, e.was, e.now]), [[1, "ann", "", "B1\t5"], [2, "ann", "B1\t5", "B1\t6"], [3, "bob", "B1\t6", "B1\t7"], [4, "ann", "B1\t7", "B1\t8"]]);
    const at3 = await call(s, p.ann, "GET", "/api/books/b/text?rev=3");
    assert.match(at3.text, /^# rexxsheet text 1\n# revision 3\n/);
    assert.match(at3.text, /\nB1\t7\n/);
    assert.equal((await call(s, p.ann, "GET", "/api/books/b/text?rev=9")).status, 400);
    const cp = await call(s, p.ann, "POST", "/api/books/b/checkpoints/before-q3");
    assert.equal(cp.json.rev, 4);
    assert.match(readFileSync(join(p.data, "books", "b", "checkpoints", "before-q3.sheet.txt"), "utf8"), /# revision 4\n[\s\S]*\nB1\t8\n/);
  } finally { await stop(s); }
});

test("a torn last line (a write cut short) is cut off at open, said, and the log goes on from the last whole op", async () => {
  const p = people("torn");
  let s = await start(p.data);
  await create(s, p.ann, "b");
  assert.equal((await call(s, p.ann, "POST", "/api/books/b/ops", op(1, "cell.set cell=B1 input=6", [cellChange(keyedLines(START), "Sheet1", "B1", "6")]))).status, 200);
  await stop(s, "SIGKILL");
  const log = join(p.data, "books", "b", "ops.log");
  const whole = statSync(log).size;
  appendFileSync(log, '{"rev":3,"base":2,"id":"torn","who":"ann","at":1,"cmd":"cell.set cell=B1 inp');
  s = await start(p.data);
  try {
    assert.match(s.out, /b: the log's last line has no end \(a write cut short\): cut off/);
    assert.equal(statSync(log).size, whole, "the file cut back to its last whole line");
    const g = await call(s, p.ann, "GET", "/api/books/b");
    assert.equal(g.json.rev, 2);
    assert.match(g.json.text, /\nB1\t6\n/);
    const next = await call(s, p.ann, "POST", "/api/books/b/ops", op(2, "cell.set cell=B1 input=7", [{ key: "Sheet1\tB1", was: "B1\t6", now: "B1\t7" }]));
    assert.equal(next.json.rev, 3);
    assert.deepEqual(readFileSync(log, "utf8").trim().split("\n").map((l) => JSON.parse(l).rev), [1, 2, 3]);
  } finally { await stop(s); }
});

test("killed mid-stream (SIGKILL while two people send): every acknowledged op is there after a restart, and the text is the log folded", async () => {
  const p = people("kill");
  let s = await start(p.data);
  await create(s, p.ann, "b");
  const acked = new Map();   // op id -> rev
  let stopping = false;
  // Two writers on disjoint cells, each sending as fast as answers come.
  const writer = async (token, col) => {
    let rev = 1, n = 0;
    const mine = new Map();
    while (!stopping) {
      n++;
      const key = `Sheet1\t${col}${10 + (n % 50)}`;
      const o = op(rev, `cell.set cell=${col}${10 + (n % 50)} input=${n}`, [{ key, was: mine.get(key) ?? "", now: `${col}${10 + (n % 50)}\t${n}` }]);
      try {
        const r = await call(s, token, "POST", "/api/books/b/ops", o);
        if (r.status === 200) { acked.set(o.id, r.json.rev); mine.set(key, `${col}${10 + (n % 50)}\t${n}`); rev = r.json.rev; }
        else if (r.status === 409) rev = r.json.rev;
      } catch { break; }
    }
  };
  const ws = [writer(p.ann, "D"), writer(p.bob, "E")];
  await sleep(700);
  s.proc.kill("SIGKILL");
  stopping = true;
  await Promise.all(ws);
  await new Promise((r) => (gone_(s.proc) ? r() : s.proc.once("exit", r)));
  assert.ok(acked.size > 5, `ops acknowledged before the kill: ${acked.size}`);
  s = await start(p.data);
  try {
    const log = readFileSync(join(p.data, "books", "b", "ops.log"), "utf8").trim().split("\n").map((l) => JSON.parse(l));
    const inLog = new Map(log.map((o) => [o.id, o.rev]));
    for (const [id, rev] of acked) assert.equal(inLog.get(id), rev, `acknowledged op ${id} (revision ${rev}) is in the log`);
    assert.deepEqual(log.map((o) => o.rev), log.map((_, i) => i + 1), "revisions run 1, 2, 3 ... with no gap");
    const folded = new Map();
    for (const o of log) applyChanges(folded, o.changes.map(([key, was, now]) => ({ key, was, now })));
    const g = await call(s, p.ann, "GET", "/api/books/b");
    assert.equal(g.json.rev, log.length);
    assert.equal(g.json.text, textOfLines(folded), "the server's text is its log folded (replay)");
    console.log(`killed after ${acked.size} acknowledged ops; ${log.length} revisions in the log`);
  } finally { await stop(s); }
});

test("two people's ops interleaved, some on the same cells: every kept op is in the text, every refused one was told why, and a third reader's fold agrees", async () => {
  const p = people("interleave");
  const s = await start(p.data);
  try {
    await create(s, p.ann, "b");
    // Each keeps its own view of the workbook, as a sae does: the lines at the revision it last heard of.
    const person = (token) => ({ token, rev: 1, lines: keyedLines(START), kept: 0, refused: 0 });
    const ann = person(p.ann), bob = person(p.bob);
    const catchUp = async (who) => {
      const r = await call(s, who.token, "GET", `/api/books/b/ops?since=${who.rev}`);
      for (const o of r.json.ops) applyChanges(who.lines, o.changes.map(([key, was, now]) => ({ key, was, now })));
      who.rev = r.json.rev;
    };
    const cells = ["B1", "B2", "F1", "F2", "F3"];
    for (let i = 0; i < 60; i++) {
      for (const who of [ann, bob]) {
        const cell = cells[(i * (who === ann ? 1 : 3)) % cells.length];
        const ch = [cellChange(who.lines, "Sheet1", cell, String(i))];
        const r = await call(s, who.token, "POST", "/api/books/b/ops", op(who.rev, `cell.set cell=${cell} input=${i}`, ch));
        if (r.status === 200) who.kept++;
        else { assert.equal(r.status, 409); assert.ok(r.json.conflict.who); who.refused++; }
        if (i % 3 === 0) await catchUp(who);
      }
    }
    await catchUp(ann);
    await catchUp(bob);
    const g = await call(s, p.ann, "GET", "/api/books/b");
    assert.equal(g.json.rev, 1 + ann.kept + bob.kept);
    assert.equal(textOfLines(ann.lines), g.json.text);
    assert.equal(textOfLines(bob.lines), g.json.text);
    assert.ok(ann.refused + bob.refused > 0, "some edits met a newer change to their cell");
    console.log(`interleaved: ann kept ${ann.kept} refused ${ann.refused}, bob kept ${bob.kept} refused ${bob.refused}`);
  } finally { await stop(s); }
});

test("the app's pages are served from the same origin as the API", async () => {
  const p = people("pages");
  const s = await start(p.data);
  try {
    const page = await fetch(`${s.url}/sheet`);
    assert.equal(page.status, 200);
    assert.equal(await page.text(), readFileSync(join(ROOT, "app", "sheet.ts"), "utf8"));
    assert.equal((await fetch(`${s.url}/sheet/teamtext.ts`)).status, 200);
    assert.equal((await fetch(`${s.url}/sheet/view/grid.ts`)).status, 200);
    assert.equal((await fetch(`${s.url}/../people`)).status, 404);
  } finally { await stop(s); }
});
