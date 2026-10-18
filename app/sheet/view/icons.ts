// The toolbar's icons: each a glyph on a 24 x 24 grid, drawn in vg (no
// image files, the same on every platform, recoloured with a fill), as
// strokes of round-capped lines; a few letters and signs are text, as
// Excel's own buttons show them (B, I, U, $, %, the comma). A colour bar
// under a glyph (Fill Color, Font Color) is the colour it applies.
export interface Icon {
  readonly strokes?: string;       // path data, stroked
  readonly fills?: string;         // path data, filled
  readonly text?: string;          // a letter or sign, centred
  readonly weight?: "bold" | "italic" | "underline" | "double";
  readonly size?: number;          // a text's size, where not the default
  readonly bar?: string;           // a colour bar along the foot
}

export const ICONS: Readonly<Record<string, Icon>> = {
  "Undo": { strokes: "M9 13 L4 8 L9 3 M4 8 H14 A6 6 0 0 1 14 20 H10" },
  "Redo": { strokes: "M15 13 L20 8 L15 3 M20 8 H10 A6 6 0 0 0 10 20 H14" },
  "Cut": { strokes: "M4 18 A3 3 0 1 0 10 18 A3 3 0 1 0 4 18 M14 18 A3 3 0 1 0 20 18 A3 3 0 1 0 14 18 M8.5 15.5 L18 3 M15.5 15.5 L6 3" },
  "Copy": { strokes: "M9 8 H20 V21 H9 Z M5 17 H4 V4 H15 V5" },
  "Paste": { strokes: "M9 3 H15 V6 H9 Z M7 4.5 H5 V21 H19 V4.5 H17 M9 11 H15 M9 15 H15" },
  "B": { text: "B", weight: "bold" },
  "I": { text: "I", weight: "italic" },
  "U": { text: "U", weight: "underline" },
  "Left": { strokes: "M4 6 H20 M4 10 H13 M4 14 H20 M4 18 H13" },
  "Center": { strokes: "M4 6 H20 M7.5 10 H16.5 M4 14 H20 M7.5 18 H16.5" },
  "Right": { strokes: "M4 6 H20 M11 10 H20 M4 14 H20 M11 18 H20" },
  ",": { text: "000" },
  "$": { text: "$" },
  "%": { text: "%" },
  // 4.2: fonts and alignment on the toolbar.
  "Font": { text: "Aa", size: 12 },
  "Double U": { text: "U", weight: "double" },
  "Super": { strokes: "M3 11 L10 21 M10 11 L3 21 M14 8.5 C14 3 20 3 20 6 C20 8 14 9 14 10.5 H20" },
  "Sub": { strokes: "M3 3 L10 13 M10 3 L3 13 M14 20.5 C14 15 20 15 20 18 C20 20 14 21 14 22.5 H20" },
  "Top": { strokes: "M4 4 H20 M9 8 H15 V20 H9 Z" },
  "Middle": { strokes: "M4 12 H20 M9 5 H15 V19 H9 Z" },
  "Bottom": { strokes: "M4 20 H20 M9 4 H15 V16 H9 Z" },
  "Indent+": { strokes: "M4 5 H20 M12 10 H20 M12 15 H20 M4 20 H20 M4 9 L8 12.5 L4 16" },
  "Indent-": { strokes: "M4 5 H20 M12 10 H20 M12 15 H20 M4 20 H20 M8 9 L4 12.5 L8 16" },
  "Angle": { strokes: "M4 20 L17 7 M10 7 H17 V14 M3 4 H9" },
  "Borders": { strokes: "M4 4 H20 V20 H4 Z M4 12 H20 M12 4 V20" },
  "Fill": { strokes: "M5 11 L11 5 L18 12 L12 18 Z M5 11 H18 M20 13 C21 15 21.5 16 21.5 17 A1.5 1.5 0 0 1 18.5 17 C18.5 16 19 15 20 13", bar: "#ffff00" },
  "Color": { strokes: "M6 17 L12 3 L18 17 M8.4 11.5 H15.6", bar: "#ff0000" },
  "A→Z": { strokes: "M18 4 V19 M15 16 L18 19 L21 16 M4 10 L6.5 3 L9 10 M5 7.5 H8 M4 13 H9 L4 20 H9" },
  "AutoFilter": { strokes: "M3 5 H21 L14 13 V19 L10 21 V13 Z" },
  "Chart": { strokes: "M3 20 H21 M6 20 V13 M10.5 20 V7 M15 20 V11 M19.5 20 V4" },
};
