// The command bar: the seam under the page, in reach. A command line
// ("data.sortDescending range=F4:G7 by=G") runs as the panels' buttons run
// theirs; a RexxJS script (a textarea: a clause a line, or ';' between
// them) drives the same commands through ADDRESS SHEET, SETCELL and GETCELL.
// What came back, and the script's SAY lines, show below. Specs drive the
// sheet through it as a person would.
import type { SheetView } from "./view.ts";

export function commandBar(v: SheetView): void {
  const line = state(""), script = state(""), said = state("");
  ui.hstack(() => {
    ui.on_submit(bind(ui.textfield("command", () => {}), line), async () => {
      const o = await v.runLine(line.value);
      said.set(o.output === undefined ? o.message : `${o.message} → ${JSON.stringify(o.output)}`);
    });
    bind(ui.textarea("RexxJS script"), script);
    v.btn("Run script", async () => {
      const lines: string[] = [];
      const o = await v.cmds.script(script.value, (l) => {
        lines.push(l);
        print(`sheet: say ${l}`);
      });
      v.refresh();
      if (!o.ok) print(`sheet: script failed: ${o.message}`);
      said.set([...lines, o.ok ? "" : `script failed: ${o.message}`].filter((l) => l).join(" | "));
    });
    v.btn("Commands", () => said.set(v.cmds.catalogue().map((c) => c.id).join(", ")));
  });
  bind(ui.text_wrapped("", 720), said);
}

// Record Macro (Script > Record macro): the name field's macro recorded
// from now, stopped, run; its REXX shown in the text area to be edited and
// saved; the workbook's macros listed beside.
export function macroBar(v: SheetView): void {
  const name = state("Macro1"), text = state(""), listed = state("");
  const list = async () => { const o = await v.run("macro.list"); listed.set(`Macros: ${o.message}`); };
  ui.hstack(() => {
    bind(ui.textfield("macro name", () => {}), name);
    v.btn("Record macro", () => v.run("macro.record", { name: name.value }));
    v.btn("Stop recording", async () => {
      const o = await v.run("macro.stop");
      if (o.ok) text.set(String(o.output));
      await list();
    });
    v.btn("Run macro", async () => {
      const o = await v.run("macro.run", { name: name.value });
      for (const l of (o.output as string[] | undefined) ?? []) print(`sheet: say ${l}`);
    });
    v.btn("Show macro", async () => { const o = await v.run("macro.get", { name: name.value }); if (o.ok) text.set(String(o.output)); });
    v.btn("Save macro", async () => { await v.run("macro.set", { name: name.value, text: text.value }); await list(); });
    v.btn("Delete macro", async () => { await v.run("macro.delete", { name: name.value }); await list(); });
    // The macro run whenever the workbook opens, or after a change to the
    // selection's cells.
    v.btn("Run macro on open", () => v.run("macro.trigger", { name: name.value, on: "open" }));
    v.btn("Run macro on a change here", () => v.run("macro.trigger", { name: name.value, on: "change", range: v.range() }));
    bind(ui.text(""), listed);
  });
  bind(ui.textarea("macro REXX"), text);
}
