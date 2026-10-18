// Drawings and charts floating over the grid at their cells (6.11c-b): each a
// container aether-ui's datagrid_float holds over the grid body at its cell,
// scrolled and clipped with the cells, the pointer passing through to them,
// its scene drawn in vg. Reconciled after every refresh: what is the same
// stays, what changed (or is above something that changed, so that the
// drawing order stands) is made again, what moved is moved.
import { fromBase64 } from "../csv.ts";
import type { Shape } from "../charts.ts";
import { CHART_H, CHART_W } from "../workbook.ts";
import type { SheetView } from "./view.ts";

export function draw(s: Shape): void {
  switch (s.kind) {
    case "rect": vg.rect(s.x, s.y, s.w, s.h, () => vg.fill(s.fill)); return;
    case "line": vg.line(s.x1, s.y1, s.x2, s.y2, () => vg.stroke(s.stroke, s.width)); return;
    case "polyline": vg.polyline([...s.points], () => { vg.fill("none"); vg.stroke(s.stroke, s.width); vg.linejoin("round"); }); return;
    case "circle": vg.circle(s.cx, s.cy, s.r, () => vg.fill(s.fill)); return;
    case "path": vg.path(s.d, () => { vg.fill(s.fill); vg.stroke(s.stroke, 1); }); return;
    case "text": vg.text_anchored(s.x, s.y, s.size, s.anchor, s.text, () => vg.fill(s.fill)); return;
  }
}

interface Wanted { readonly key: string; readonly sig: string; readonly row: number; readonly col: number; readonly w: number; readonly h: number; readonly paint: () => void }
interface Live { readonly key: string; readonly sig: string; readonly handle: number; pos: string }

const live = new Map<number, Live[]>();   // by grid

export function floatCount(v: SheetView): number {
  return (live.get(v.grid) ?? []).length;
}

export function syncFloats(v: SheetView): void {
  if (!v.grid || typeof ui.datagrid_float !== "function") return;
  const z = v.wb.zoom / 100;
  const px = (n: number) => Math.max(1, Math.round(n * z));
  const want: Wanted[] = [];
  v.charts.value.forEach((c) => {
    const sig = `${z}|${c.title}|${JSON.stringify(c.shapes)}`;
    want.push({ key: `chart ${c.n}`, sig, row: c.at.row, col: c.at.col + 1, w: px(CHART_W), h: px(CHART_H), paint: () => {
      vg.scene(`0 0 ${CHART_W} ${CHART_H}`, px(CHART_W), px(CHART_H), () => {
        vg.rect(0, 0, CHART_W, CHART_H, () => vg.fill("#ffffff"));
        c.shapes.forEach(draw);
      });
    } });
  });
  for (const d of v.drawings.value) {
    const sig = `${z}|${d.kind}|${d.shape ?? ""}|${d.w}x${d.h}|${d.fill ?? ""}|${d.line ?? ""}|${d.text ?? ""}|${d.picture?.length ?? 0}|${d.shapes.length}`;
    want.push({ key: `drawing ${d.name}`, sig, row: d.at.row, col: d.at.col + 1, w: px(d.w), h: px(d.h), paint: () => {
      vg.scene(`0 0 ${d.w} ${d.h}`, px(d.w), px(d.h), () => {
        if (d.kind === "picture" && d.picture) {
          vg.image(fromBase64(d.picture), () => { vg.box(0, 0, d.w, d.h); vg.fit("contain"); });
          return;
        }
        d.shapes.forEach(draw);
      });
    } });
  }
  const have = live.get(v.grid) ?? [];
  let k = 0;
  while (k < have.length && k < want.length && have[k].key === want[k].key && have[k].sig === want[k].sig) k++;
  // Something gone: all are made again (the page's newest scene is then a live one, not
  // the gone one's, which the driver still lists).
  if (have.length > want.length) k = 0;
  for (let i = have.length - 1; i >= k; i--) ui.datagrid_float_remove(v.grid, have[i].handle);
  have.length = k;
  for (let i = 0; i < want.length; i++) {
    const w = want[i], pos = `${w.row},${w.col},${w.w},${w.h}`;
    if (i < k) {
      if (have[i].pos !== pos) {
        ui.datagrid_float_move(v.grid, have[i].handle, w.row, w.col, 0, 0, w.w, w.h);
        have[i].pos = pos;
      }
      continue;
    }
    const f = ui.datagrid_float(v.grid, w.row, w.col, 0, 0, w.w, w.h);
    if (!f) break;
    ui.into(f, w.paint);
    have.push({ key: w.key, sig: w.sig, handle: f, pos });
  }
  live.set(v.grid, have);
}
