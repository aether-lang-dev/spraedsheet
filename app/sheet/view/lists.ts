// What the sheet says about itself: the status line, the functions its
// language has, the setup script, and the problems (each cell whose value
// is an error), with a count.
import type { Problem } from "../sheet.ts";
import type { SheetView } from "./view.ts";
import { sliderToZoom, zoomToSlider } from "./zoom.ts";

export function statusLine(v: SheetView): void {
  bind(ui.text(""), v.status);
}

// The status bar under the grid, as Excel's: Ready, or the selection's
// average, count, numerical count, min, max and sum; the sheet's cells;
// Calculate when manual calculation has left cells stale; the control bus
// when it is on. At its right the zoom slider (zoom.ts).
export function statusBar(v: SheetView): void {
  ui.hstack(6, () => {
    bind(ui.text(""), v.figures);
    ui.spacer();
    const out = ui.btn("−", () => v.run("view.zoom", { by: "out" }));
    ui.a11y_label(out, "Zoom out a step");
    v.zoomSlider = ui.slider(0, 100, zoomToSlider(v.wb.zoom), (s: number) => {
      const z = sliderToZoom(s);
      if (z !== v.wb.zoom) v.run("view.zoom", { percent: String(z) });
    });
    ui.a11y_label(v.zoomSlider, "Zoom");
    ui.set_width(v.zoomSlider, 120);
    const inn = ui.btn("+", () => v.run("view.zoom", { by: "in" }));
    ui.a11y_label(inn, "Zoom in a step");
    v.zoomLabel = ui.text(`${v.wb.zoom}%`);
  });
}

export function functionList(v: SheetView): void {
  bind(ui.text_wrapped("", 720), computed(() => {
    const own = v.formulas.value === "RexxJS" ? ", and RexxJS's own (UPPER, SUBSTR, WORD, LENGTH, ...)" : "";
    return `${v.formulas.value} functions: ${v.wb.sheet.language.functions().join(", ")}${own}`;
  }, v.formulas));
}

export function setupScript(v: SheetView): void {
  bind(ui.text_wrapped("", 720), v.setup);
}

export function problemList(v: SheetView): void {
  bind(ui.text(""), computed(() => {
    const n = v.problems.value.length;
    return n === 0 ? "No problems" : `${n} problem${n === 1 ? "" : "s"}`;
  }, v.problems));
  each(v.problems, "key", (p: Problem) => ui.text(`${p.key}: ${p.message}`));
}
