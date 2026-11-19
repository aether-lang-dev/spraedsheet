// A team workbook (docs/design/teamware-over-http.md): this workbook is one a
// sheet-server keeps, which a small team edits together. Each change a
// person makes here is applied at once, then sent as an operation (the
// command, and the lines of the workbook's text it changed, teamtext.ts); the
// server keeps it, in its order, or refuses it when someone changed the same
// lines (or the sheet's shape) first. Others' operations come back on a
// long-poll and are applied here as they arrive. Undo takes back this
// person's own last kept change, as a new operation.
//
// What reaches the server is a Transport (the page's `http`, the token sent
// by the page itself; a stand-in in the tests), so nothing here is the page's.
import { applyChanges, changesBetween, inverse, keyedLines, sheetOfKey, textOfLines } from "./teamtext.ts";
import type { Change } from "./teamtext.ts";
import { writeSheetFile } from "./file.ts";
import { field, textOfFile } from "./textform.ts";
import type { Commands, Outcome, Params } from "./commands.ts";
import type { Workbook } from "./workbook.ts";
import { parseLine } from "./recorder.ts";

export interface Answer {
  readonly status: number;   // 0: nothing came back
  readonly text: string;
  readonly error?: string;
}
export type Transport = (method: string, path: string, body?: string) => Promise<Answer>;

// A transport over the page's `http`, its token in the Authorization header
// (sent by the page, never by sae on its behalf).
export function pageTransport(token: string): Transport {
  return async (method, path, body) => {
    const res = await http.fetch({ method, url: path, headers: { Authorization: `Bearer ${token}` }, body, contentType: "application/json" });
    return { status: res.status, text: res.text, error: res.error };
  };
}

// The commands whose change moves cells about on their sheet: an edit made
// from before one of these is refused there (its addresses moved under it).
export const RESHAPING = new Set(["rows.insert", "rows.delete", "columns.insert", "columns.delete", "columns.move", "data.sort", "data.sortAscending",
  "data.sortDescending", "table.sort", "data.subtotal", "data.removeDuplicates"]);

// The commands that would put another workbook in this one's place.
const REPLACING = new Set(["file.open", "file.new", "file.importText", "file.import"]);

interface Op {
  readonly id: string;
  readonly cmd: string;
  readonly structure: readonly string[];
  readonly changes: readonly Change[];
  readonly undo?: true;       // a take-back of this person's own edit
  readonly redo?: true;
}

interface Kept { readonly rev: number; readonly op: Op }

// Someone else here, and the cell they are on.
export interface Person { readonly name: string; readonly cell: string }

// A command that may have changed the workbook: its id, the sheet it ran
// on, and (a paste) whether it moved what was cut.
export interface TeamEdit { readonly id: string; readonly sheet: string; readonly cut: boolean }

// What the Commands it is attached to ask of it (commands.ts's `team`).
export interface TeamHook {
  before(id: string, params: Params): Promise<Outcome | null>;
  edited(cmd: string, done: readonly TeamEdit[]): void;
}

let opSeq = 0;

export class TeamBook implements TeamHook {
  readonly #cmds: Commands;
  readonly #wb: Workbook;
  readonly #send: Transport;
  readonly name: string;
  readonly #me: string;
  // The server's workbook as far as this page has heard: its revision and lines.
  rev = 0;
  #lines = new Map<string, string>();
  // This person's edits not yet kept, in order, and the lines the workbook here shows (the server's, with them).
  readonly #pending: Op[] = [];
  #local = new Map<string, string>();
  // This person's kept edits, newest last (what Undo takes back), and the undone ones (Redo).
  readonly #mine: Kept[] = [];
  readonly #undone: Op[] = [];
  // What the team line says (and the last refusal, notice, which stays until
  // this person's next edit), who is here, and where this person is (a key:
  // sheet, tab, cell), sent with each poll.
  status = "";
  #notice = "";
  people: Person[] = [];
  where: () => string = () => "";
  #sending = false;
  #following = false;
  #closed = false;
  // Every change of state goes through this, one at a time.
  #queue: Promise<unknown> = Promise.resolve();
  // Told when the team line changed (onChange), and when the workbook here
  // was read again from the server's (onShown): only then does the page
  // show it again, so what a person is typing in the formula bar stays.
  onChange: () => void = () => {};
  onShown: () => void = () => {};
  // The log of what happened, for the specs (print in the page).
  log: (line: string) => void = () => {};

  constructor(cmds: Commands, wb: Workbook, send: Transport, name: string, me: string) {
    this.#cmds = cmds;
    this.#wb = wb;
    this.#send = send;
    this.name = name;
    this.#me = me;
  }

  #serial<T>(fn: () => Promise<T>): Promise<T> {
    const next = this.#queue.then(fn, fn);
    this.#queue = next.catch(() => {});
    return next;
  }

  // The workbook from the server, made if `create` and it is not there yet
  // (from what this page shows now). "" when open, else why not.
  open(create = false): Promise<string> {
    return this.#serial(async () => {
      let a = await this.#send("GET", `/api/books/${this.name}`);
      if (a.status === 404 && create) {
        // Made from what this page shows, under the team workbook's name.
        this.#wb.name = this.name;
        const lines = keyedLines(this.#text());
        const made = await this.#send("POST", `/api/books/${this.name}`, JSON.stringify(this.#wire({ id: this.#id(), cmd: "create", structure: [], changes: changesBetween(new Map(), lines) }, 0)));
        if (made.status !== 200) return `could not make ${this.name}: ${this.#why(made)}`;
        a = await this.#send("GET", `/api/books/${this.name}`);
      }
      if (a.status !== 200) return `could not open ${this.name}: ${this.#why(a)}`;
      const got = JSON.parse(a.text) as { rev: number; text: string };
      this.rev = got.rev;
      this.#lines = keyedLines(got.text);
      await this.#show(new Map(this.#lines));
      this.#say(`${this.name}, revision ${this.rev}`);
      return "";
    });
  }

  // Others' edits, as they come, until close().
  follow(): void {
    if (this.#following) return;
    this.#following = true;
    void (async () => {
      while (!this.#closed) {
        const a = await this.#send("GET", `/api/books/${this.name}/ops?since=${this.rev}&wait=10&cell=${encodeURIComponent(this.where())}`);
        if (this.#closed) break;
        if (a.status === 200) await this.#serial(() => this.#heard(JSON.parse(a.text)));
        else {
          this.#say(`not connected (${this.#why(a)}); trying again`);
          await new Promise((r) => setTimeout(r, 2000));
        }
      }
    })();
  }

  close(): void { this.#closed = true; }

  // Asked before each command: refused in a team workbook, or Undo and Redo
  // done here (they take back this person's own kept edits).
  async before(id: string, _params: Params): Promise<Outcome | null> {
    if (REPLACING.has(id)) return { ok: false, message: `${this.name} is a team workbook: open another one beside it instead (New Workbook)`, changed: [] };
    if (id === "edit.undo") return this.#serial(() => this.#takeBack());
    if (id === "edit.redo") return this.#serial(() => this.#doAgain());
    return null;
  }

  // Told after each command that may have changed the workbook (the
  // outermost one, a script's or a macro's whole run once): its lines now,
  // against what the workbook showed before it, are the op. At once, before
  // anything else runs: the workbook and #local agree until then (Commands
  // hold() keeps the server's versions from landing while a command runs).
  edited(cmd: string, done: readonly TeamEdit[]): void {
    const now = keyedLines(this.#text());
    const changes = changesBetween(this.#local, now);
    if (!changes.length) return;
    this.#local = now;
    this.#undone.length = 0;
    if (this.#notice) { this.#notice = ""; this.onChange(); }
    const structure = [...new Set(done.filter((d) => RESHAPING.has(d.id) || d.cut).map((d) => field(d.sheet)))];
    this.#pending.push({ id: this.#id(), cmd, structure, changes });
    this.#pump();
  }

  // How many edits are on their way.
  get waiting(): number { return this.#pending.length; }

  // ---- sending

  #pump(): void {
    if (this.#sending || !this.#pending.length) return;
    this.#sending = true;
    void (async () => {
      while (this.#pending.length && !this.#closed) {
        const op = this.#pending[0];
        const a = await this.#send("POST", `/api/books/${this.name}/ops`, JSON.stringify(this.#wire(op, this.rev)));
        if (a.status === 200) {
          // Kept: heard back in order (with whatever others did first) before the next is sent.
          const kept = JSON.parse(a.text) as { rev: number };
          const r = await this.#send("GET", `/api/books/${this.name}/ops?since=${this.rev}`);
          await this.#serial(async () => {
            if (r.status === 200) await this.#heard(JSON.parse(r.text));
            if (op.undo) this.#undone.push({ ...op, undo: undefined } as Op);
            else this.#mine.push({ rev: kept.rev, op });
          });
        } else if (a.status === 409) {
          await this.#serial(() => this.#refused(op, JSON.parse(a.text)));
        } else {
          // Not reached, or the server failed: try the same op again (the server knows it by its id).
          this.#say(`not sent (${this.#why(a)}); trying again`);
          await new Promise((r) => setTimeout(r, 2000));
        }
      }
      this.#sending = false;
    })();
  }

  // The server refused `op`: what it said, the workbook put back to the
  // server's (with this person's other edits still on their way).
  async #refused(op: Op, answer: { conflict: { key: string; why: string; rev: number; who: string; at: number; cmd: string } }): Promise<void> {
    // Only this op: the later ones still waiting are sent as they are, and the
    // server refuses any whose lines this one had changed.
    const at = this.#pending.indexOf(op);
    if (at >= 0) this.#pending.splice(at, 1);
    const c = answer.conflict;
    const where = c.key.endsWith("\t") ? `the sheet ${c.key.slice(0, -1)}` : describeKey(c.key);
    const when = c.at ? new Date(c.at).toTimeString().slice(0, 5) : "";
    const theirs = describeCmd(c.cmd);
    const what = c.why === "reshaped" ? `${c.who} changed the shape of ${sheetOfKey(c.key) || "the workbook"} (${theirs}) at ${when} (revision ${c.rev}) after you began, so its cells moved`
      : c.why === "edited" ? `${c.who} changed ${where} (${theirs}) at ${when} (revision ${c.rev}) after you began, and yours moves its cells`
      : `${c.who} changed ${where} (${theirs}) at ${when} (revision ${c.rev}) after you began`;
    const verb = op.undo ? "Not undone" : "Not kept";
    await this.#show(this.#serverWithPending());
    this.#notice = `${verb}: your ${describeCmd(op.cmd)}. ${what}; the sheet shows theirs now. ${op.undo ? "" : "Make your change again to replace it."}`.replace(/\s+$/, "");
    this.#say("");
    this.log(`team: refused ${op.cmd}: ${c.why} ${c.key.replace("\t", "!")} by ${c.who} r${c.rev}`);
  }

  // ---- hearing

  async #heard(got: { rev: number; ops: { rev: number; id: string; who: string; cmd: string; changes: [string, string, string][] }[]; people?: Person[] }): Promise<void> {
    if (got.people) this.people = got.people.filter((p) => p.name !== this.#me);
    let others = 0;
    for (const o of got.ops) {
      if (o.rev <= this.rev) continue;
      applyChanges(this.#lines, o.changes.map(([key, was, now]) => ({ key, was, now })));
      this.rev = o.rev;
      const mine = this.#pending.findIndex((p) => p.id === o.id);
      if (mine === 0) this.#pending.shift();
      else {
        others++;
        this.log(`team: r${o.rev} ${o.who}: ${o.cmd}`);
      }
    }
    if (others) await this.#show(this.#serverWithPending());
    this.#say(others ? `${this.name}, revision ${this.rev}: ${got.ops.filter((o) => o.rev <= this.rev).slice(-1)[0]?.who ?? ""} changed it` : `${this.name}, revision ${this.rev}`);
  }

  // ---- undo and redo: this person's own kept edits

  async #takeBack(): Promise<Outcome> {
    const last = this.#mine.pop();
    if (!last) return { ok: false, message: "nothing of yours to undo in this team workbook", changed: [] };
    return this.#local_(last.op.changes, `Undo ${last.op.cmd}`, true);
  }

  async #doAgain(): Promise<Outcome> {
    const op = this.#undone.pop();
    if (!op) return { ok: false, message: "nothing to redo", changed: [] };
    return this.#local_(inverse(op.changes), op.cmd.replace(/^Undo /, ""), false, true);
  }

  // A take-back (or redo) applied here and sent as an op.
  async #local_(changes: readonly Change[], cmd: string, undo: boolean, redo = false): Promise<Outcome> {
    const back = undo ? inverse(changes) : [...changes];
    const shown = new Map(this.#local);
    applyChanges(shown, back);
    await this.#show(shown);
    const op: Op = { id: this.#id(), cmd, structure: [], changes: back, ...(undo ? { undo: true as const } : {}), ...(redo ? { redo: true as const } : {}) };
    this.#pending.push(op);
    this.#pump();
    return { ok: true, message: `${cmd}: sent`, changed: back.map((c) => describeKey(c.key)) };
  }

  // ---- the workbook here

  #text(): string { return textOfFile(writeSheetFile(this.#wb)); }

  // The server's lines, and this person's edits still on their way over them.
  #serverWithPending(): Map<string, string> {
    const out = new Map(this.#lines);
    for (const p of this.#pending) applyChanges(out, p.changes);
    return out;
  }

  // The workbook shown as `lines` (#local from then on): read again from its
  // text with no command running, on the sheet this person was on, its
  // selection kept.
  async #show(lines: Map<string, string>): Promise<void> {
    await this.#cmds.hold(async () => {
      const sheet = this.#wb.active, selection = this.#wb.selection;
      this.#local = lines;
      const o = await this.#cmds.runHeld("file.importText", { text: textOfLines(lines) });
      if (!o.ok) this.log(`team: the workbook could not be read back: ${o.message}`);
      if (this.#wb.sheetNames().includes(sheet)) await this.#cmds.runHeld("sheet.activate", { name: sheet });
      this.#wb.selection = selection;
    });
    this.onShown();
  }

  #wire(op: Op, base: number): object {
    return { id: op.id, base, cmd: op.cmd, structure: op.structure, changes: op.changes.map((c) => [c.key, c.was, c.now]) };
  }

  #id(): string { return `${this.#me}-${Date.now().toString(36)}-${(++opSeq).toString(36)}-${Math.floor(Math.random() * 1e9).toString(36)}`; }

  #why(a: Answer): string {
    if (a.status === 0) return a.error || "no answer";
    try { return JSON.parse(a.text).error ?? `HTTP ${a.status}`; } catch { return `HTTP ${a.status}`; }
  }

  #say(s: string): void {
    if (s) this.status = s;
    this.onChange();
  }

  // The last refusal, in words, until this person's next edit ("" when none).
  get notice(): string { return this.#notice; }
}

// A key as a person reads it: "Sheet1!B2", "Sheet1 names QTY", "the workbook's formulas".
export function describeKey(key: string): string {
  const [sheet, part, entry] = key.split("\t");
  if (!part) return sheet;
  if (!part.startsWith("@")) return `${sheet}!${part}`;
  const what = part.slice(1).replace(/\[\]$/, "");
  return `${sheet ? `${sheet} ` : "the workbook's "}${what}${entry ? ` ${entry}` : ""}`;
}

// A command as a person reads it: a cell set as "B2 = 5", any other as its line.
export function describeCmd(cmd: string): string {
  const { id, params } = parseLine(cmd);
  if (id === "cell.set" && params.cell !== undefined && Object.keys(params).length === 2) return `${params.cell} = ${params.input ?? ""}`;
  if (cmd.startsWith("Undo ")) return `Undo ${describeCmd(cmd.slice(5))}`;
  return cmd;
}
