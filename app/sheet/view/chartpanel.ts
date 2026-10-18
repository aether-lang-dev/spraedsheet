// Charts of the selection, each floating over the grid at its cell (floats.ts:
// drawn in vg from the shapes charts.ts lays out, redrawn as its cells
// change), each with buttons here, to move it to the selected cell, delete it, and the rest; and a sparkline of the selection, as text in
// the cell to its right.
import { keyOf } from "../address.ts";
import type { ChartView } from "../workbook.ts";
import type { SheetView } from "./view.ts";

export function chartPanel(v: SheetView): void {
  const insert = (kind: string) => () => v.run("chart.insert", { kind, range: v.range() });
  ui.hstack(() => {
    v.btn("Bar chart", insert("bar"));
    v.btn("Line chart", insert("line"));
    v.btn("Pie chart", insert("pie"));
    v.btn("Stacked column chart", insert("barStacked"));
    v.btn("Horizontal bar chart", insert("hbar"));
    v.btn("Area chart", insert("area"));
    v.btn("Doughnut chart", insert("doughnut"));
    v.btn("Scatter chart", insert("scatter"));
    v.btn("Radar chart", insert("radar"));
    v.btn("Histogram", insert("histogram"));
    v.btn("Waterfall chart", insert("waterfall"));
    v.btn("Funnel chart", insert("funnel"));
    v.btn("Treemap", insert("treemap"));
    v.btn("Sunburst", insert("sunburst"));
    v.btn("Box and whisker", insert("box"));
    v.btn("Stock chart", insert("stock"));
    v.btn("Combo chart", insert("combo"));
    // A sparkline in the cell right of the selection's top row.
    const spark = (type: string, options = "") => () => {
      const s = v.selection();
      if (!s) return;
      const right = Math.max(s[0].col, s[1].col) + 1, top = Math.min(s[0].row, s[1].row);
      return v.run("insert.sparkline", { range: v.range(), at: keyOf({ col: right, row: top }), ...(type !== "line" ? { type } : {}), ...(options ? { options } : {}) });
    };
    v.btn("Sparkline", spark("line"));
    v.btn("Win/loss sparkline", spark("winloss", "axis"));
    v.btn("Sparkline with high and low", spark("line", "high;low"));
  });
  each(v.charts, (c: ChartView & { n: number }) => `${c.n}`, (c: ChartView & { n: number }) => {
    ui.hstack(() => {
      ui.text(`${c.title}: ${c.kind} chart of ${c.range}`);
      ui.btn(`Move chart ${c.n} here`, () => v.run("chart.move", { index: String(c.n), at: v.here() }));
      ui.btn(`Delete chart ${c.n}`, () => v.run("chart.delete", { index: String(c.n) }));
      // Chart Elements and Switch Row/Column, on and off.
      const o = v.wb.page(v.wb.active)?.charts[c.n - 1]?.options ?? {};
      ui.btn(`Legend right ${c.n}`, () => v.run("chart.option", { index: String(c.n), legend: o.legend === "right" ? "bottom" : "right" }));
      ui.btn(`Data labels ${c.n}`, () => v.run("chart.option", { index: String(c.n), labels: o.labels ? "0" : "1" }));
      ui.btn(`Gridlines ${c.n}`, () => v.run("chart.option", { index: String(c.n), gridlines: o.gridlines === false ? "1" : "0" }));
      ui.btn(`Switch row/column ${c.n}`, () => v.run("chart.option", { index: String(c.n), byRows: o.byRows ? "0" : "1" }));
      ui.btn(`Use selection ${c.n}`, () => v.run("chart.data", { index: String(c.n), range: v.range() }));
      ui.btn(`Linear trendline ${c.n}`, () => v.run("chart.trendline", { index: String(c.n), series: "1", type: o.trendlines?.length ? "none" : "linear", equation: "1", r2: "1" }));
    });
  });
}
