// Rows and columns out of view, over the selection, as Excel has them: Hide
// and Unhide rows and columns; a filter on the selected column of the
// selection's table (the criterion as COUNTIF writes one: >100, north,
// <>x), Clear filter; Freeze top row and Unfreeze; the selection's rows set
// to a height (empty or 0: the grid's).
import { addrOf, colName } from "../address.ts";
import type { SheetView } from "./view.ts";

export function viewBar(v: SheetView): void {
  const criterion = state(""), height = state("");
  ui.hstack(() => {
    v.btn("Hide rows", () => v.run("rows.hide", { range: v.range() }));
    v.btn("Unhide rows", () => v.run("rows.unhide", { range: v.range() }));
    v.btn("Hide columns", () => v.run("columns.hide", { range: v.range() }));
    v.btn("Unhide columns", () => v.run("columns.unhide", { range: v.range() }));
    bind(ui.textfield("filter by", () => {}), criterion);
    v.btn("Filter", () => v.run("data.filter", { range: v.range(), column: colName(addrOf(v.here())!.col), criterion: criterion.value }));
    v.btn("Clear filter", () => v.run("data.clearFilter"));
    v.btn("Freeze top row", () => v.run("view.freezeRows", { count: "1" }));
    v.btn("Split", () => v.run("view.split", v.wb.split ? { off: "1" } : {}));
    v.btn("Freeze first column", () => v.run("view.freezeCols", { count: "1" }));
    v.btn("Freeze panes", () => v.run("view.freezePanes", { cell: v.here() }));
    v.btn("Unfreeze", async () => { await v.run("view.freezeRows", { count: "0" }); await v.run("view.freezeCols", { count: "0" }); });
    bind(ui.textfield("row height", () => {}), height);
    v.btn("Zoom in", () => v.run("view.zoom", { by: "in" }));
    v.btn("Zoom out", () => v.run("view.zoom", { by: "out" }));
    v.btn("Zoom 100%", () => v.run("view.zoom", { percent: "100" }));
    // View > Animations: the selection's glide on or off.
    v.btn("Animations", () => v.run("view.animations", v.wb.animations ? { off: "1" } : {}));
    // Dark mode: the chrome dark or light, or the system's.
    v.btn("Dark mode", () => v.run("view.appearance", { mode: "dark" }));
    v.btn("Light mode", () => v.run("view.appearance", { mode: "light" }));
    v.btn("System appearance", () => v.run("view.appearance", { mode: "system" }));
    // Gridlines, Headings and the Formula Bar, each shown or not.
    v.btn("Gridlines", () => v.run("view.show", { what: "gridlines", on: v.wb.page(v.wb.active)!.gridlines ? "0" : "1" }));
    v.btn("Headings", () => v.run("view.show", { what: "headings", on: v.wb.page(v.wb.active)!.headings ? "0" : "1" }));
    v.btn("Formula bar", () => v.run("view.show", { what: "formulabar", on: v.wb.formulaBar ? "0" : "1" }));
    // Page Break Preview, and the page breaks at the selected cell.
    v.btn("Page break preview", () => v.run("view.pageBreakPreview", v.wb.breakPreview ? { off: "1" } : {}));
    v.btn("Insert page break", () => v.run("page.break", { cell: v.here() }));
    v.btn("Remove page break", () => v.run("page.break", { cell: v.here(), off: "1" }));
    v.btn("Reset page breaks", () => v.run("page.resetBreaks"));
    v.btn("AutoFit columns", () => v.run("columns.autofit", { range: v.range() }));
    v.btn("AutoFit rows", () => v.run("rows.autofit", { range: v.range() }));
    v.btn("Row height", () => v.run("rows.height", { range: v.range(), height: height.value.trim() || "0" }));
  });
}
