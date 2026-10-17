// One toolbar row under the menus, as old Excel's Standard and Formatting
// toolbars had it (not a ribbon): Undo, Redo; Cut, Copy, Paste; Bold,
// Italic, Underline; Align Left, Center, Align Right; Comma, Currency and
// Percent Style; Borders; Fill Color (yellow) and Font Color (red), as
// Excel's buttons start; Sort Ascending, AutoFilter (the filter field's
// criterion); a chart of the selection. Each a short label with its full
// name and shortcut as its tooltip, and each running what the panels' and
// keys' do (their actions, by label, at the click: the panels come after).
import { toggleFont } from "./lookbar.ts";
import type { SheetView } from "./view.ts";
import { ICONS } from "./icons.ts";

// The groups, by their first tool's label: a wider gap before each.
export const GROUPS = ["Undo", "Cut", "B", "Left", ",", "Borders", "Fill", "A→Z", "Chart"];

export const TOOLS: readonly (readonly [string, string, string])[] = [
  // [label, tooltip, the action it runs or the command]
  ["Undo", "Undo (Ctrl+Z)", "Undo"], ["Redo", "Redo (Ctrl+Y)", "Redo"],
  ["Cut", "Cut (Ctrl+X)", "Cut"], ["Copy", "Copy (Ctrl+C)", "Copy"], ["Paste", "Paste (Ctrl+V)", "Paste"],
  ["B", "Bold (Ctrl+B)", "Bold"], ["I", "Italic (Ctrl+I)", "Italic"], ["U", "Underline (Ctrl+U)", "Underline"],
  ["Left", "Align Left", "Align left"], ["Center", "Center", "Align center"], ["Right", "Align Right", "Align right"],
  [",", "Comma Style", "Comma style"], ["$", "Currency Style", "Currency style"], ["%", "Percent Style", "Percent style"],
  ["Borders", "Borders: an outline round the selection", "Outline borders"],
  ["Fill", "Fill Color: yellow", "Fill yellow"], ["Color", "Font Color: red", "Text red"],
  ["A→Z", "Sort Ascending, by the selected cell's column", "Sort A→Z"], ["AutoFilter", "AutoFilter: the selection's rows by the filter field's criterion", "Filter"],
  ["Chart", "Chart of the selection", "Bar chart"],
];

export function toolBar(v: SheetView): void {
  // The toolbar's own actions, for the menus too.
  v.action("Align left", () => v.run("format.align", { range: v.range(), align: "left" }));
  v.action("Align center", () => v.run("format.align", { range: v.range(), align: "center" }));
  v.action("Align right", () => v.run("format.align", { range: v.range(), align: "right" }));
  v.action("Comma style", () => v.run("format.number", { range: v.range(), format: "number" }));
  v.action("Currency style", () => v.run("format.number", { range: v.range(), format: "currency" }));
  v.action("Percent style", () => v.run("format.number", { range: v.range(), format: "percent" }));
  v.action("Fill yellow", () => v.run("format.fill", { range: v.range(), color: v.wb.notes.get(v.here()).look?.fill ? "none" : "yellow" }));
  v.action("Text red", () => v.run("format.color", { range: v.range(), color: v.wb.notes.get(v.here()).look?.color ? "none" : "red" }));
  v.action("Bold", () => toggleFont(v, "bold"));
  v.action("Italic", () => toggleFont(v, "italic"));
  v.action("Underline", () => toggleFont(v, "underline"));
  // Each group its own row, close set; the groups set apart.
  const groups = GROUPS.map((first, i) => TOOLS.slice(TOOLS.findIndex(([l]) => l === first), i + 1 < GROUPS.length ? TOOLS.findIndex(([l]) => l === GROUPS[i + 1]) : TOOLS.length));
  ui.hstack(14, () => {
    for (const group of groups) {
      ui.hstack(2, () => {
        for (const [label, tip, action] of group) {
          const run = () => {
            if (v.actions.has(action)) return v.perform(action);
            v.status.set(`${tip}: not here`);
          };
          tool(label, tip, run);
        }
      });
    }
  });
}

// A tool: its icon (icons.ts) in a 26 px square, the click running it; its tooltip the tool's name and shortcut,
// and to a screen reader a button named so.
// The ink, light or dark (inkTools): each inked shape, and how.
export const INK = { light: "#33373d", dark: "#e8e8e8" };
const inked: { readonly h: number; readonly stroke?: number; readonly fill?: boolean }[] = [];
function tool(label: string, tip: string, run: () => unknown): number {
  const icon = ICONS[label] ?? { text: label };
  const ink = INK.light;
  const scene: number = vg.scene("0 0 24 24", 26, 26, () => {
    if (icon.strokes) inked.push({ h: vg.path(icon.strokes, () => { vg.fill("none"); vg.stroke(ink, 1.8); vg.linecap("round"); vg.linejoin("round"); }), stroke: 1.8 });
    if (icon.fills) inked.push({ h: vg.path(icon.fills, () => vg.fill(ink)), fill: true });
    if (icon.text) {
      const size = icon.text.length > 1 ? 9 : 15;
      inked.push({ h: vg.text_anchored(12, icon.text.length > 1 ? 15 : 17, size, "middle", icon.text, () => {
        vg.fill(ink);
        if (icon.weight === "bold") vg.stroke(ink, 0.9);
        if (icon.weight === "italic") vg.transform("skewX(-14) translate(4 0)");
      }), fill: true, ...(icon.weight === "bold" ? { stroke: 0.9 } : {}) });
      if (icon.weight === "underline") inked.push({ h: vg.line(7, 20, 17, 20, () => vg.stroke(ink, 1.6)), stroke: 1.6 });
    }
    if (icon.bar) vg.rect(3, 20.5, 18, 3, () => vg.fill(icon.bar!));
    // Over the glyph, unseen: what the pointer clicks and rests on.
    vg.rrect(0, 0, 24, 24, 3, () => {
      vg.fill("#ffffff");
      vg.opacity(0);
      vg.tooltip(tip);
      vg.cursor("pointer");
      vg.on_click(() => run());
    });
  });
  // Its 26 px: a scene's canvas otherwise takes the row's spare width.
  ui.set_width(scene, 26);
  ui.a11y_role(scene, "button");
  ui.a11y_label(scene, tip);
  v_tools.set(label, scene);
  return scene;
}

// The tools' scenes by label, for a spec or the menus to find.
export const v_tools = new Map<string, number>();

// Every icon inked for a dark ground or a light one (dark mode).
export function inkTools(dark: boolean): void {
  const c = dark ? INK.dark : INK.light;
  for (const i of inked) vg.set(i.h, { ...(i.fill ? { fill: c } : {}), ...(i.stroke ? { stroke: [c, i.stroke] } : {}) });
}
