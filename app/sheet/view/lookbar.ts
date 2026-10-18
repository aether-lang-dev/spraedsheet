// A cell's look, over the selection, as Excel's Home tab has it: Bold,
// Italic, Underline and Strikethrough (each on, or off when the selected
// cell has it), outline, all, inside, bottom, thick outline or no borders, Wrap text (its rows as tall as
// their text), and Merge & center, Merge and Unmerge (a merge keeps its
// top-left's value), and Format painter. Ctrl+B, Ctrl+I and Ctrl+U on the grid do the first three.
// Then Excel's Format Cells > Font and Alignment: Double underline,
// Superscript and Subscript, a font family (the field, empty for the
// theme's), the text's place in the cell's height (Top, Middle, Bottom),
// its indent (Increase and Decrease) and its angle (counterclockwise,
// clockwise, up, down, horizontal).
import type { SheetView } from "./view.ts";

export type FontPart = "bold" | "italic" | "underline" | "strike";

// The part toggled over the selection, by what the selected cell has.
export const toggleFont = (v: SheetView, part: FontPart) =>
  v.run("format.font", { range: v.range(), [part]: v.wb.notes.get(v.here()).look?.[part] ? "0" : "1" });

// The font part set over the selection, or taken off when the selected cell has it.
const fontPart = (v: SheetView, key: "underline" | "script", on: string) => {
  const look = v.wb.notes.get(v.here()).look;
  const has = key === "underline" ? look?.underline === "double" : look?.script === on;
  return v.run("format.font", { range: v.range(), [key]: has ? (key === "underline" ? "0" : "none") : on });
};

export function lookBar(v: SheetView): void {
  const family = state("");
  // Format Painter: the first click takes the selection's format, the
  // second paints it over the selection then.
  let painting: string | null = null;
  ui.wrap(() => {
    v.btn("Format painter", async () => {
      if (painting === null) {
        painting = v.range();
        v.status.set(`Format painter: ${painting}'s format taken; select where to paint it, then Format painter again`);
        print(`sheet: format painter took ${painting}`);
        return;
      }
      const from = painting;
      painting = null;
      await v.run("format.painter", { from, to: v.range() });
    });
    v.btn("Bold", () => toggleFont(v, "bold"));
    v.btn("Italic", () => toggleFont(v, "italic"));
    v.btn("Underline", () => toggleFont(v, "underline"));
    v.btn("Strikethrough", () => toggleFont(v, "strike"));
    v.btn("Outline borders", () => v.run("format.border", { range: v.range(), edge: "outline" }));
    v.btn("All borders", () => v.run("format.border", { range: v.range(), edge: "all" }));
    v.btn("Inside borders", () => v.run("format.border", { range: v.range(), edge: "inside" }));
    v.btn("Bottom border", () => v.run("format.border", { range: v.range(), edge: "bottom" }));
    v.btn("Thick outline", () => v.run("format.border", { range: v.range(), edge: "outline", style: "thick" }));
    v.btn("No borders", () => v.run("format.border", { range: v.range(), edge: "none" }));
    v.btn("Wrap text", () => v.run("format.wrap", v.wb.notes.get(v.here()).look?.wrap ? { range: v.range(), off: "1" } : { range: v.range() }));
    v.btn("Merge & center", () => v.run("cells.merge", { range: v.range(), center: "1" }));
    v.btn("Merge", () => v.run("cells.merge", { range: v.range() }));
    v.btn("Unmerge", () => v.run("cells.unmerge", { range: v.range() }));
  });
  // Fonts and alignment beyond the toolbar's (4.2).
  ui.wrap(() => {
    v.btn("Double underline", () => fontPart(v, "underline", "double"));
    v.btn("Superscript", () => fontPart(v, "script", "super"));
    v.btn("Subscript", () => fontPart(v, "script", "sub"));
    bind(ui.textfield("font family", () => {}), family);
    v.btn("Font family", () => v.run("format.font", { range: v.range(), family: family.value.trim() }));
    v.btn("Align top", () => v.run("format.align", { range: v.range(), vertical: "top" }));
    v.btn("Align middle", () => v.run("format.align", { range: v.range(), vertical: "middle" }));
    v.btn("Align bottom", () => v.run("format.align", { range: v.range(), vertical: "bottom" }));
    const indent = (by: number) => () => v.run("format.align", { range: v.range(), indent: String(Math.max(0, Math.min(15, (v.wb.notes.get(v.here()).look?.indent ?? 0) + by))) });
    v.btn("Increase indent", indent(1));
    v.btn("Decrease indent", indent(-1));
    const turn = (rotation: number) => () => v.run("format.align", { range: v.range(), rotation: String(rotation) });
    v.btn("Angle counterclockwise", turn(45));
    v.btn("Angle clockwise", turn(-45));
    v.btn("Rotate text up", turn(90));
    v.btn("Rotate text down", turn(-90));
    v.btn("Horizontal text", turn(0));
  });
  // Cell Styles: the gallery's commonest, on the selection.
  ui.wrap(() => {
    for (const name of ["Normal", "Good", "Bad", "Neutral", "Title", "Heading 1", "Total", "Input", "Note"]) v.btn(`Style ${name}`, () => v.run("style.apply", { range: v.range(), name }));
    // Themes: what names a theme's colours re-coloured.
    for (const name of ["Office", "Greyscale", "Harbour", "Ember"]) v.btn(`Theme ${name}`, () => v.run("theme.set", { name }));
  });
}
