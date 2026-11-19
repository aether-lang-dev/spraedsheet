// The team line (docs/design/teamware-over-http.md): on a page that came
// from a sheet-server (http://<server>/sheet?book=<name>), this workbook is
// the team's. The line says the workbook, its revision, and what happened
// (a refusal, in words, stays until this person's next edit), and who else
// has it open and where; the first time, it asks for this person's token,
// which the page keeps in its origin's storage and sends itself.
import type { SheetView } from "./view.ts";
import { TeamBook, describeKey, pageTransport } from "../team.ts";
import { field } from "../textform.ts";

const TOKEN = "team-token";

// The team workbook a page's address names: its ?book=, on an http(s) page
// that may reach its own server (seeks outgoing-http). Null for any other.
export function teamBookOf(url: string): string | null {
  if (typeof http === "undefined" || !/^https?:\/\//.test(url)) return null;
  const m = /[?&]book=([A-Za-z0-9_-]{1,64})(?:&|#|$)/.exec(url);
  return m ? m[1] : null;
}

export function teamBar(v: SheetView, book: string): void {
  const line = state(`Team workbook ${book}: connecting`);
  const here = state("");
  const notice = state("");
  let field_ = 0;
  ui.hstack(6, () => {
    ui.text("Team:");
    bind(ui.text(""), line);
    ui.spacer();
    bind(ui.text(""), here);
  });
  // A refusal, in words, on a line of its own until this person's next edit.
  const told = bind(ui.text(""), notice);
  ui.set_visible(told, false);
  // History: who changed the selected cell, and any revision as a workbook of
  // its own beside this one, in the read view.
  const changes = state<string[]>([]);
  let revField = 0;
  const past = ui.hstack(6, () => {
    v.btn("Who changed this cell", async () => {
      const team = v.team;
      if (!team) return;
      const cell = v.here(), got = await team.history(`${field(v.wb.active)}\t${cell}`);
      if (typeof got === "string") { changes.set([`History of ${cell}: ${got}`]); return; }
      const shown = (line: string) => (line === "" ? "(empty)" : line.split("\t")[1] ?? "");
      changes.set(got.length ? [`${cell}, oldest first:`, ...got.map((h) => `revision ${h.rev}, ${h.who}, ${new Date(h.at).toTimeString().slice(0, 5)}: ${shown(h.now)}${h.was === "" ? "" : ` (was ${shown(h.was)})`}`)]
        : [`${cell} has not been changed since the workbook was made`]);
      print(`sheet: history of ${cell}: ${got.length} change${got.length === 1 ? "" : "s"}`);
    });
    revField = ui.textfield("revision", () => {});
    ui.a11y_label(revField, "revision to look at");
    v.btn("Look at revision", async () => {
      const team = v.team, rev = Number(ui.get_text(revField).trim());
      if (!team || !Number.isInteger(rev) || rev < 1) return;
      const text = await team.textAt(rev);
      if (text === null) return;
      // A workbook of its own, beside the team one, in the read view.
      await v.run("book.new", { name: `${book} at revision ${rev}` });
      const o = await v.cmds.execute("file.importText", { text });
      v.wb.name = `${book} at revision ${rev}`;
      await v.run("view.editMode", { on: "0" });
      print(`sheet: ${book} at revision ${rev}: ${o.message}, read view`);
    });
  });
  ui.set_visible(past, false);
  each(changes, (l: string) => l, (l: string) => ui.text(l));
  const row = ui.hstack(6, () => {
    field_ = ui.textfield("your team token (from whoever runs the server)", () => {});
    ui.a11y_label(field_, "your team token");
    v.btn("Sign in", () => signIn(ui.get_text(field_).trim()));
  });
  ui.set_visible(row, false);

  const signIn = async (token: string): Promise<void> => {
    if (!token) return;
    const send = pageTransport(token);
    const me = await send("GET", "/api/me");
    if (me.status !== 200) {
      line.set(me.status === 401 ? "this server does not know that token: ask for yours" : `cannot reach the server (${me.error || `HTTP ${me.status}`})`);
      ui.set_visible(row, true);
      return;
    }
    storage.set(TOKEN, token);
    ui.set_visible(row, false);
    const name = String(JSON.parse(me.text).name);
    const team = new TeamBook(v.cmds, v.wb, send, book, name);
    team.log = (l) => print(`sheet: ${l}`);
    team.where = () => `${field(v.wb.active)}\t${v.address.value}`;
    team.onChange = () => {
      line.set(`${team.status}${team.waiting ? ` (${team.waiting} of your edits on the way)` : ""}`);
      if (notice.value !== team.notice) { notice.set(team.notice); ui.set_visible(told, team.notice !== ""); }
      here.set(team.people.length ? `Here: ${team.people.map((p) => `${p.name}${p.cell ? ` on ${describeKey(p.cell)}` : ""}`).join(", ")}` : "");
    };
    let shown = false;
    // The workbook read again: the grid and the panels show it.
    team.onShown = () => {
      team.onChange();
      if (shown) { v.refresh(); if (v.grid) ui.datagrid_refresh(v.grid); }
    };
    const why = await team.open(true);
    if (why) { line.set(why); return; }
    shown = true;
    v.cmds.team = team;
    v.team = team;
    ui.set_visible(past, true);
    v.refresh();
    if (v.grid) ui.datagrid_refresh(v.grid);
    print(`sheet: team workbook ${book} open as ${name}, revision ${team.rev}`);
    team.follow();
  };

  const kept = storage.get(TOKEN);
  if (kept) void signIn(kept);
  else {
    line.set(`Team workbook ${book}: sign in with your token`);
    ui.set_visible(row, true);
  }
}
