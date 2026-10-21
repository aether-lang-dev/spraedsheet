// Formula auditing, as Excel's Formulas tab has it: trace what the selected
// cell reads (blue) or what reads it (orange), a level further each time;
// remove the arrows; Show Formulas (on, then off); Evaluate Formula's steps; and the Watch
// Window, the cells watched on any sheet as they change. And Calculation
// Options: manual calculation (on, then off), Calculate now (F9) and
// Calculate sheet (Shift+F9), iterative calculation (on, then off). Error
// checking lists what to check on the sheet, each with Go to and Ignore.
import { parseRange } from "../address.ts";
import type { Finding } from "../errorcheck.ts";
import type { Watch } from "../workbook.ts";
import type { SheetView } from "./view.ts";

export function auditBar(v: SheetView): void {
  const steps = state(""), info = state<string[]>([]);
  const findings = state<Finding[]>([]);
  const check = async () => {
    const o = await v.run("formulas.errorCheck");
    findings.set(o.ok ? o.output as Finding[] : []);
  };
  ui.wrap(() => {
    v.btn("Trace precedents", () => v.run("formulas.tracePrecedents", { cell: v.here() }));
    v.btn("Trace dependents", () => v.run("formulas.traceDependents", { cell: v.here() }));
    v.btn("Remove arrows", () => v.run("formulas.removeArrows"));
    v.btn("Show formulas", () => v.run("formulas.showFormulas", v.wb.showFormulas ? { off: "1" } : {}));
    v.btn("Evaluate formula", async () => {
      const o = await v.run("formulas.evaluate", { cell: v.here() });
      steps.set(o.ok ? o.message : "");
    });
    // Everything non-default about the selected cell, a part to a line.
    v.btn("Cell info", async () => {
      const o = await v.run("cell.info", { cell: v.here() });
      const about = o.output as Record<string, unknown>;
      info.set(o.ok ? [`Cell info ${about.cell}:`, ...Object.entries(about).filter(([k]) => k !== "cell").map(([k, x]) => `${k}: ${typeof x === "string" ? x : JSON.stringify(x)}`)] : []);
    });
    v.btn("Add watch", () => v.run("formulas.watch", { cell: v.here() }));
  });
  ui.wrap(() => {
    v.btn("Manual calculation", () => v.run("formulas.calculation", { mode: v.wb.calculation.manual ? "automatic" : "manual" }));
    v.btn("Calculate now", () => v.run("formulas.recalculate"));
    v.btn("Calculate sheet", () => v.run("formulas.calculateSheet"));
    v.btn("Iterative calculation", () => v.run("formulas.iteration", v.wb.calculation.iterate ? { off: "1" } : {}));
    v.btn("Error checking", check);
    v.btn("1904 date system", () => v.run("workbook.dateSystem", { system: v.wb.dateSystem === 1904 ? "1900" : "1904" }));
  });
  each(findings, (f: Finding) => `${f.cell} ${f.check}`, (f: Finding) => ui.wrap(() => {
    ui.text(`${f.cell} ${f.check}: ${f.message}`);
    ui.btn(`Go to ${f.cell}`, () => v.select(parseRange(f.cell)!));
    ui.btn(`Ignore ${f.cell}`, async () => {
      await v.run("formulas.ignoreError", { cell: f.cell });
      await check();
    });
  }));
  // Checks (B.6): formulas that must be TRUE, worked out on every recalculation;
  // a failing one is named here with its message, and in the status bar.
  const cname = state(""), cformula = state(""), cmessage = state("");
  ui.wrap(() => {
    bind(ui.textfield("check name", () => {}), cname);
    bind(ui.textfield("check formula (must be TRUE)", () => {}), cformula);
    bind(ui.textfield("check message", () => {}), cmessage);
    v.btn("Add check", () => v.run("check.add", { name: cname.value, formula: cformula.value, message: cmessage.value }));
  });
  each(v.checks, (c: { name: string; ok: boolean; shown: string }) => `${c.name}|${c.ok}|${c.shown}`, (c: { name: string; formula: string; message: string; ok: boolean; shown: string }) => ui.wrap(() => {
    ui.text(c.ok ? `Check ${c.name}: passing` : `Check ${c.name}: FAILING, ${c.message} (it gave ${c.shown})`);
    ui.btn(`Remove check ${c.name}`, () => v.run("check.remove", { name: c.name }));
  }));
  bind(ui.text_wrapped("", 720), steps);
  each(info, (l: string) => l, (l: string) => ui.text(l));
  each(v.watches, (w: Watch) => `${w.sheet}!${w.cell}`, (w: Watch) => ui.wrap(() => {
    ui.text(`Watch ${w.sheet}!${w.cell}: ${w.input} = ${w.shown}`);
    ui.btn(`Unwatch ${w.sheet}!${w.cell}`, () => v.run("formulas.unwatch", { cell: `${w.sheet}!${w.cell}` }));
  }));
}
