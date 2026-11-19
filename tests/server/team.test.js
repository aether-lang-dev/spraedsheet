// Two people's spreadsheets on one team workbook, headless: each a real
// Workbook and Commands (app/sheet/, run on node) with a TeamBook
// (app/sheet/team.ts) talking to the real sheet-server over real HTTP. What
// the sae page does, without the window: the spec on a lab host
// (tests/specs/spec_team.ae) drives the same through two sae-drivers.
//   node --test tests/server/team.test.js
import { test } from "node:test";
import assert from "node:assert/strict";
import { spawn, spawnSync } from "node:child_process";
import { mkdirSync, rmSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

globalThis.print ??= () => {};
const { Commands } = await import("../../app/sheet/commands.ts");
const { Workbook } = await import("../../app/sheet/workbook.ts");
const { TeamBook } = await import("../../app/sheet/team.ts");
const { keyedLines, textOfLines } = await import("../../app/sheet/teamtext.ts");
const { writeSheetFile } = await import("../../app/sheet/file.ts");
const { textOfFile } = await import("../../app/sheet/textform.ts");

const ROOT = join(dirname(fileURLToPath(import.meta.url)), "..", "..");
const BIN = process.env.SHEET_SERVER ?? join(ROOT, "target", "sheet-server");
const WORK = join(ROOT, "target", "team-tests");
let nextPort = Number(process.env.SHEET_TEST_PORT ?? 18393) + 40;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function server(name) {
  const data = join(WORK, name);
  rmSync(data, { recursive: true, force: true });
  mkdirSync(data, { recursive: true });
  const token = (who) => spawnSync(BIN, ["add-person", data, who], { encoding: "utf8" }).stdout.trim();
  const tokens = { ann: token("ann"), bob: token("bob") };
  const port = nextPort++;
  const proc = spawn(BIN, ["serve", data, join(ROOT, "app"), String(port)], { stdio: "ignore" });
  const url = `http://127.0.0.1:${port}`;
  for (let i = 0; i < 100; i++) {
    try { if ((await fetch(`${url}/api/info`)).ok) break; } catch { /* not yet */ }
    await sleep(50);
  }
  return { url, tokens, data, stop: () => new Promise((r) => { proc.once("exit", r); proc.kill(); }) };
}

// A person: their own workbook and commands, and the team book over HTTP.
// While `deaf` is set, the answers to their long-polls are held back (they
// have not heard yet), which is how a test makes an edit from an older revision.
async function person(s, who, book, setup = []) {
  const ears = { deaf: false };
  const wb = new Workbook();
  const cmds = new Commands(wb);
  for (const line of setup) await cmds.run(line);
  const send = async (method, path, body) => {
    if (path.includes("/ops?") && path.includes("wait=")) while (ears.deaf) await sleep(20);
    try {
      const res = await fetch(`${s.url}${path}`, { method, headers: { Authorization: `Bearer ${s.tokens[who]}`, "Content-Type": "application/json" }, body });
      return { status: res.status, text: await res.text() };
    } catch (e) { return { status: 0, text: "", error: String(e) }; }
  };
  const team = new TeamBook(cmds, wb, send, book, who);
  const log = [];
  team.log = (l) => log.push(l);
  const why = await team.open(true);
  assert.equal(why, "", `${who} opened ${book}`);
  cmds.team = team;
  team.follow();
  return { who, wb, cmds, team, log, ears, run: (line) => cmds.run(line), text: () => textOfLines(keyedLines(textOfFile(writeSheetFile(wb)))) };
}

async function until(what, fn, ms = 5000) {
  const t0 = Date.now();
  while (Date.now() - t0 < ms) {
    if (await fn()) return Date.now() - t0;
    await sleep(20);
  }
  assert.fail(`waited ${ms} ms for ${what}`);
}

const serverText = async (s, book) => {
  const r = await fetch(`${s.url}/api/books/${book}`, { headers: { Authorization: `Bearer ${s.tokens.ann}` } });
  return (await r.json()).text;
};
const settled = (p) => p.team.waiting === 0;

test("Ann's edit reaches Bob's workbook, its formulas worked out there; both are the server's text", async () => {
  const s = await server("reach");
  const ann = await person(s, "ann", "budget", ["formulas.use language=Excel", "cell.set cell=A1 input=Tea", "cell.set cell=B1 input=3", 'cell.set cell=C1 input="=B1*2"']);
  const bob = await person(s, "bob", "budget");
  try {
    assert.equal(String(bob.wb.sheet.value("C1")), "6", "Bob opened Ann's workbook, worked out");
    const t0 = Date.now();
    await ann.run("cell.set cell=B1 input=5");
    const lag = await until("Bob to see B1 = 5", () => bob.wb.sheet.input("B1") === "5");
    assert.equal(String(bob.wb.sheet.value("C1")), "10", "Bob's formula follows Ann's cell");
    await bob.run('cell.set cell=D1 input="=C1+1"');
    await until("Ann to see D1", () => String(ann.wb.sheet.value("D1")) === "11");
    await until("both settled", () => settled(ann) && settled(bob) && ann.team.rev === bob.team.rev);
    const text = await serverText(s, "budget");
    assert.equal(ann.text(), text);
    assert.equal(bob.text(), text);
    console.log(`an edit reached the other person in ${lag} ms (${Date.now() - t0} ms with the edit)`);
  } finally { ann.team.close(); bob.team.close(); await s.stop(); }
});

test("a paste (Ann's own clipboard) reaches Bob as what it did, not as a paste of Bob's clipboard", async () => {
  const s = await server("paste");
  const ann = await person(s, "ann", "b", ["cell.set cell=A1 input=one", "cell.set cell=A2 input=two", "format.fill range=A1 color=yellow"]);
  const bob = await person(s, "bob", "b");
  try {
    await ann.run("edit.copy range=A1:A2");
    await ann.run("edit.paste at=C5");
    await until("Bob to see the paste", () => bob.wb.sheet.input("C6") === "two");
    assert.equal(bob.wb.sheet.input("C5"), "one");
    await until("both settled", () => settled(ann) && ann.team.rev === bob.team.rev);
    assert.equal(bob.text(), await serverText(s, "b"));
    assert.equal(ann.text(), bob.text(), "the look came with it too");
  } finally { ann.team.close(); bob.team.close(); await s.stop(); }
});

test("the same cell from both at once: one kept, the other told plainly whose change won, and shown it", async () => {
  const s = await server("same");
  const ann = await person(s, "ann", "b", ["cell.set cell=B2 input=1"]);
  const bob = await person(s, "bob", "b");
  try {
    // Both type into B2 before either hears of the other.
    await Promise.all([ann.run("cell.set cell=B2 input=100"), bob.run("cell.set cell=B2 input=200")]);
    await until("both settled on one value", () => settled(ann) && settled(bob) && ann.team.rev === bob.team.rev && ann.wb.sheet.input("B2") === bob.wb.sheet.input("B2"), 8000);
    const won = ann.wb.sheet.input("B2");
    const loser = won === "100" ? bob : ann, winner = won === "100" ? "ann" : "bob";
    assert.match(loser.team.notice, new RegExp(`^Not kept: your B2 = ${won === "100" ? 200 : 100}\\. ${winner} changed Sheet1!B2 \\(B2 = ${won}\\) at \\d\\d:\\d\\d \\(revision \\d+\\) after you began; the sheet shows theirs now\\. Make your change again to replace it\\.$`));
    assert.match(loser.team.status, /^b, revision \d+/, "the team line goes on beside it");
    assert.ok(loser.log.some((l) => /team: refused cell\.set cell=B2/.test(l)));
    assert.equal(ann.text(), await serverText(s, "b"));
    assert.equal(bob.text(), ann.text());
    console.log(`conflict: ${loser.who} saw "${loser.team.notice}"`);
    await loser.run("cell.set cell=Z9 input=next");
    assert.equal(loser.team.notice, "", "gone at their next edit");
  } finally { ann.team.close(); bob.team.close(); await s.stop(); }
});

test("a row inserted by Ann while Bob types further down: Bob's edit is refused (its cells moved), not put in the wrong row", async () => {
  const s = await server("shape");
  const ann = await person(s, "ann", "b", ["cell.set cell=A1 input=head", "cell.set cell=A5 input=five"]);
  const bob = await person(s, "bob", "b");
  try {
    // Bob has not heard of the insert when he types into A9.
    bob.ears.deaf = true;
    await sleep(50);
    await ann.run("rows.insert at=2 count=1");
    await until("Ann's insert kept", () => settled(ann));
    await bob.run("cell.set cell=A9 input=nine");
    bob.ears.deaf = false;
    await until("Bob's edit refused", () => /^Not kept: your A9 = nine\. ann changed the shape of Sheet1 \(rows\.insert/.test(bob.team.notice));
    await until("Bob settled", () => settled(bob) && bob.team.rev === ann.team.rev);
    assert.equal(bob.wb.sheet.input("A6"), "five", "Bob's workbook now has the inserted row");
    assert.equal(bob.wb.sheet.input("A9"), "", "and not his A9");
    assert.equal(bob.text(), await serverText(s, "b"));
  } finally { ann.team.close(); bob.team.close(); await s.stop(); }
});

test("undo takes back your own last kept edit; refused when someone changed that cell since", async () => {
  const s = await server("undo");
  const ann = await person(s, "ann", "b", ["cell.set cell=A1 input=1"]);
  const bob = await person(s, "bob", "b");
  try {
    await ann.run("cell.set cell=A1 input=2");
    await bob.run("cell.set cell=C3 input=bob");
    await until("settled", () => settled(ann) && settled(bob) && ann.team.rev === bob.team.rev);
    const u = await ann.cmds.execute("edit.undo");
    assert.equal(u.ok, true, u.message);
    await until("Bob sees Ann's undo", () => bob.wb.sheet.input("A1") === "1");
    assert.equal(ann.wb.sheet.input("C3"), "bob", "Bob's later edit is not undone with it");
    await ann.cmds.execute("edit.redo");
    await until("Bob sees Ann's redo", () => bob.wb.sheet.input("A1") === "2");
    // Bob changes A1 after Ann: her undo would overwrite his change, so it is refused.
    await until("settled", () => settled(ann) && ann.team.rev === bob.team.rev);
    await bob.run("cell.set cell=A1 input=3");
    await until("Ann hears", () => ann.wb.sheet.input("A1") === "3");
    await ann.cmds.execute("edit.undo");
    await until("Ann's undo refused", () => /^Not undone: your Undo A1 = 2\. bob changed Sheet1!A1 \(A1 = 3\)/.test(ann.team.notice));
    assert.equal(ann.wb.sheet.input("A1"), "3");
    assert.equal(ann.text(), await serverText(s, "b"));
  } finally { ann.team.close(); bob.team.close(); await s.stop(); }
});

test("a command that would put another workbook in the team one's place is refused", async () => {
  const s = await server("replace");
  const ann = await person(s, "ann", "b", ["cell.set cell=A1 input=1"]);
  try {
    const o = await ann.cmds.execute("file.open", { sample: "budget" });
    assert.deepEqual([o.ok, o.message], [false, "b is a team workbook: open another one beside it instead (New Workbook)"]);
    assert.equal(ann.wb.sheet.input("A1"), "1");
  } finally { ann.team.close(); await s.stop(); }
});

test("many edits from both, interleaved: they converge on the server's text", async () => {
  const s = await server("many");
  const ann = await person(s, "ann", "b", ["formulas.use language=Excel", 'cell.set cell=Z1 input="=SUM(A1:B20)"']);
  const bob = await person(s, "bob", "b");
  try {
    for (let i = 1; i <= 20; i++) {
      await ann.run(`cell.set cell=A${i} input=${i}`);
      await bob.run(`cell.set cell=B${i} input=${i * 10}`);
      if (i % 5 === 0) await bob.run(`format.fill range=A${i} color=yellow`);
    }
    // Bob's fills of A5, A10 ... met Ann's newer values there: refused, said, and the rest kept.
    await until("settled", async () => settled(ann) && settled(bob) && ann.team.rev === bob.team.rev && JSON.parse(await (await fetch(`${s.url}/api/books/b`, { headers: { Authorization: `Bearer ${s.tokens.ann}` } })).text()).rev === ann.team.rev, 15000);
    const text = await serverText(s, "b");
    assert.equal(ann.text(), text);
    assert.equal(bob.text(), text);
    assert.equal(String(ann.wb.sheet.value("Z1")), String(20 * 21 / 2 * 11));
    assert.equal(String(bob.wb.sheet.value("Z1")), String(20 * 21 / 2 * 11));
  } finally { ann.team.close(); bob.team.close(); await s.stop(); }
});

test("history: who changed a cell, oldest first; the workbook at an earlier revision, read back as a workbook", async () => {
  const s = await server("history");
  const ann = await person(s, "ann", "b", ["cell.set cell=B2 input=1"]);
  const bob = await person(s, "bob", "b");
  try {
    await ann.run("cell.set cell=B2 input=2");
    await until("Bob hears", () => bob.wb.sheet.input("B2") === "2");
    await bob.run("cell.set cell=B2 input=3");
    await until("Ann hears", () => ann.wb.sheet.input("B2") === "3");
    const h = await ann.team.history("Sheet1\tB2");
    assert.deepEqual(h.map((e) => [e.rev, e.who, e.was, e.now]), [[1, "ann", "", "B2\t1"], [2, "ann", "B2\t1", "B2\t2"], [3, "bob", "B2\t2", "B2\t3"]]);
    const text = await ann.team.textAt(2);
    const past = new Workbook(), pc = new Commands(past);
    assert.equal((await pc.execute("file.importText", { text })).ok, true);
    assert.equal(past.sheet.input("B2"), "2", "revision 2 as it was");
    assert.equal(await ann.team.textAt(99), null);
    assert.match(ann.team.status, /^no revision 99: revision 99 is not one of this workbook's \(1 to 3\)/);
  } finally { ann.team.close(); bob.team.close(); await s.stop(); }
});
