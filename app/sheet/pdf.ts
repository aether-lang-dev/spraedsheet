// PDF export: cells laid out on pages as Excel prints them (down, then
// over; A4 or Letter, portrait or landscape, its margins), each cell's text
// as it shows in Helvetica (bold, italic, its size and colour; numbers to
// the right as the grid has them), its fill, its borders, the gridlines
// and the row and column headings when asked, the sheet's name and the
// page number at the foot. A PDF 1.4 file of the standard fonts, written
// here (no library). Pure: cells in, bytes out.

// Helvetica's widths (thousandths of the size) for character codes 32 to
// 126, its AFM's; a character past them is as wide as an 'n'.
const W = [278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556,
  1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556,
  333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556, 556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584];
export const textWidth = (s: string, size: number, bold = false): number =>
  [...s].reduce((w, ch) => { const c = ch.charCodeAt(0); return w + (c >= 32 && c <= 126 ? W[c - 32] : 556); }, 0) * size / 1000 * (bold ? 1.06 : 1);

export interface PdfCell {
  readonly text: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly size?: number;           // pt
  readonly color?: string;          // #rrggbb
  readonly fill?: string;
  readonly align?: "start" | "center" | "end";
  readonly borders?: { readonly top?: string; readonly right?: string; readonly bottom?: string; readonly left?: string };  // #rrggbb
}

export interface PdfSheet {
  readonly name: string;
  readonly cells: (row: number, col: number) => PdfCell | null;
  readonly rows: number;            // how many, from the range's top
  readonly cols: number;
  readonly firstRow: number;        // the sheet's row and column numbers of the range's top-left (0-based)
  readonly firstCol: number;
  readonly colWidth: (col: number) => number;    // pt
  readonly rowHeight: (row: number) => number;   // pt
  // Manual page breaks: the rows (columns), by their index here, a page
  // begins at.
  readonly rowBreaks?: readonly number[];
  readonly colBreaks?: readonly number[];
  // Rows printed again at the top of each page after the one they first appear
  // on (Page Setup > Rows to repeat at top): their heights and cells, and the
  // sheet row (0-based) of the last of them; `sheetRow` is the sheet row of a
  // printed row, to tell which pages come after.
  readonly titles?: { readonly heights: readonly number[]; readonly cell: (i: number, col: number) => PdfCell | null; readonly last: number };
  readonly sheetRow?: (row: number) => number;
  // Columns printed again at the left of each page after the one they first appear on (Page Setup >
  // Columns to repeat at left): their widths, their cells in a printed row (and in a repeated row, the
  // corner), the sheet column (0-based) of the last of them, and each one's sheet column; `sheetCol` is
  // the sheet column of a printed column, to tell which pages come after.
  readonly titleCols?: { readonly widths: readonly number[]; readonly cell: (row: number, i: number) => PdfCell | null; readonly corner: (titleRow: number, i: number) => PdfCell | null; readonly index: readonly number[]; readonly last: number };
  readonly sheetCol?: (col: number) => number;
  // This sheet's own setup, when it is not the one the whole file is written with.
  readonly setup?: PageSetup;
}

export interface PageSetup {
  readonly paper: "a4" | "letter";
  readonly landscape: boolean;
  readonly margin: number;          // pt
  readonly gridlines: boolean;
  readonly headings: boolean;
  // Percent of full size (100 when left out), and the header and footer text
  // (&P page, &N pages, &A sheet, &F file, &L &C &R sections), none when left out.
  readonly scale?: number;
  // Fitted to this many pages wide and tall (0: as many as it takes), the largest scale up to 100% that
  // does it, in place of `scale`, and the manual breaks left out (Page Setup > Fit to).
  readonly fit?: { readonly wide: number; readonly tall: number };
  readonly header?: string;
  readonly footer?: string;
  readonly fileName?: string;
}

// What a sheet keeps of its page setup (Page.layout): the PageSetup's parts,
// the pages to fit to, the print area ("" for the used range), the rows to repeat ("1:2", or "") and the
// columns to repeat ("A:B", or "").
export interface PageLayout {
  readonly paper: "a4" | "letter";
  readonly landscape: boolean;
  readonly margin: number;
  readonly scale: number;
  readonly fitWide: number;     // pages wide to fit the sheet to; 0: as many as it takes
  readonly fitTall: number;     // pages tall; both 0: not fitted, the scale applies
  readonly area: string;
  readonly titles: string;
  readonly titleCols: string;   // the columns to repeat ("A:B", or "")
  readonly header: string;
  readonly footer: string;
  readonly gridlines: boolean;
  readonly headings: boolean;
}
export const DEFAULT_LAYOUT: PageLayout = { paper: "a4", landscape: false, margin: 36, scale: 100, fitWide: 0, fitTall: 0, area: "", titles: "", titleCols: "", header: "", footer: "", gridlines: false, headings: false };
export const PAPER: Readonly<Record<string, [number, number]>> = { a4: [595, 842], letter: [612, 792] };

const colName = (c: number): string => { let s = ""; for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
// Text as a PDF string in WinAnsi: what Latin-1 has, else ?.
const pdfText = (s: string) => `(${[...s].map((ch) => { const c = ch.charCodeAt(0); return c > 255 ? "?" : ch === "(" || ch === ")" || ch === "\\" ? `\\${ch}` : c < 32 ? " " : c > 126 ? `\\${c.toString(8).padStart(3, "0")}` : ch; }).join("")})`;
const rgb = (c: string) => { const h = c.replace("#", ""); return [0, 2, 4].map((i) => (parseInt(h.slice(i, i + 2), 16) / 255).toFixed(3)).join(" "); };
const n = (x: number) => (Math.round(x * 100) / 100).toString();

// The pages: each a run of rows and of columns that fit, down then over;
// a manual break (an index in `breaks`) begins a page there whatever fits.
export function paginate(sizes: readonly number[], room: number | ((start: number) => number), breaks: readonly number[] = []): [number, number][] {
  const out: [number, number][] = [];
  let start = 0, used = 0;
  const forced = new Set(breaks);
  const roomAt = (at: number) => (typeof room === "number" ? room : room(at));
  sizes.forEach((s, i) => {
    if ((used + s > roomAt(start) || forced.has(i)) && i > start) { out.push([start, i - 1]); start = i; used = 0; }
    used += s;
  });
  if (sizes.length) out.push([start, sizes.length - 1]);
  return out;
}

// A page's width and height in points.
export const pageSize = (setup: PageSetup): [number, number] => { const [w, h] = PAPER[setup.paper]; return setup.landscape ? [h, w] : [w, h]; };

// How a sheet falls on pages under a setup: the scaled sizes of its columns and rows, and its pages
// (a run of rows, a run of columns), down then over. `repeats(start)`: a page beginning at printed
// row `start` has the sheet's title rows at its top; `repeatsCols(start)` likewise for the columns to
// repeat at its left. A setup that fits the sheet to pages takes the largest scale (to 100%) that does,
// with the manual breaks left out, and `scale` is the one used.
export function pagination(sh: PdfSheet, setup: PageSetup) {
  const fit = setup.fit && (setup.fit.wide > 0 || setup.fit.tall > 0) ? setup.fit : null;
  if (!fit) return pagesAt(sh, setup, setup.scale ?? 100, true);
  const ok = (s: number) => { const l = pagesAt(sh, setup, s, false); return (fit.wide === 0 || l.colPages.length <= fit.wide) && (fit.tall === 0 || l.rowPages.length <= fit.tall); };
  // The largest whole percent from 10 to 100 that fits (fewer pages as the scale falls, so a search).
  if (ok(100)) return pagesAt(sh, setup, 100, false);
  let lo = 10, hi = 99;
  if (!ok(lo)) return pagesAt(sh, setup, 10, false);
  while (lo < hi) { const mid = Math.ceil((lo + hi) / 2); if (ok(mid)) lo = mid; else hi = mid - 1; }
  return pagesAt(sh, setup, lo, false);
}

function pagesAt(sh: PdfSheet, setup: PageSetup, scale: number, manual: boolean) {
  const [pw, ph] = pageSize(setup), k = scale / 100;
  const head = setup.headings ? 24 : 0, foot = 18;
  const widths = Array.from({ length: sh.cols }, (_, c) => sh.colWidth(c) * k), heights = Array.from({ length: sh.rows }, (_, r) => sh.rowHeight(r) * k);
  const titleHeights = (sh.titles?.heights ?? []).map((h) => h * k), titleTotal = titleHeights.reduce((t, h) => t + h, 0);
  const titleWidths = (sh.titleCols?.widths ?? []).map((w) => w * k), titleWidthTotal = titleWidths.reduce((t, w) => t + w, 0);
  const sheetRow = sh.sheetRow ?? ((r: number) => sh.firstRow + r);
  const sheetCol = sh.sheetCol ?? ((c: number) => sh.firstCol + c);
  const full = ph - 2 * setup.margin - foot - (setup.headings ? 14 : 0);
  // A page after the rows to repeat first appear has them at its top, and that much less room.
  const repeats = (start: number) => !!sh.titles && titleTotal > 0 && sheetRow(start) > sh.titles.last;
  const repeatsCols = (start: number) => !!sh.titleCols && titleWidthTotal > 0 && sheetCol(start) > sh.titleCols.last;
  const across = pw - 2 * setup.margin - head;
  const rowPages = paginate(heights, (start) => full - (repeats(start) ? titleTotal : 0), manual ? sh.rowBreaks : []),
    colPages = paginate(widths, (start) => across - (repeatsCols(start) ? titleWidthTotal : 0), manual ? sh.colBreaks : []);
  return { widths, heights, titleHeights, titleWidths, sheetRow, sheetCol, repeats, repeatsCols, rowPages, colPages, scale, k };
}

// A header or footer's text with its codes worked out, in its left, centre and
// right parts (text before any &L, &C or &R is the centre's).
export function headerParts(text: string, page: number, pages: number, sheet: string, file: string): [string, string, string] {
  const parts: [string, string, string] = ["", "", ""];
  let at = 1;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch !== "&" || i + 1 >= text.length) { parts[at] += ch; continue; }
    const c = text[++i].toUpperCase();
    if (c === "L") at = 0; else if (c === "C") at = 1; else if (c === "R") at = 2;
    else if (c === "P") parts[at] += String(page); else if (c === "N") parts[at] += String(pages);
    else if (c === "A") parts[at] += sheet; else if (c === "F") parts[at] += file; else if (c === "&") parts[at] += "&";
  }
  return parts;
}

export function writePdf(sheets: readonly PdfSheet[], base: PageSetup): { bytes: Uint8Array; pages: number } {
  const streams: string[] = [];
  const pageInfo: { name: string; setup: PageSetup }[] = [];
  const sizes: [number, number][] = [];
  for (const sh of sheets) {
    const setup = sh.setup ?? base;
    const [pw, ph] = pageSize(setup), head = setup.headings ? 24 : 0;
    const { widths, heights, titleHeights, titleWidths, sheetRow, sheetCol, repeats, repeatsCols, rowPages, colPages, k } = pagination(sh, setup);
    for (const [c1, c2] of colPages) for (const [r1, r2] of rowPages) {
      const ops: string[] = [];
      const x0 = setup.margin + head, top = ph - setup.margin - (setup.headings ? 14 : 0);
      // The columns of the page: the repeated ones, then the page's own; and its lines: the repeated rows,
      // then the page's own. A part is a title (its index among them) or a printed one.
      type Part = { size: number; title: boolean; i: number };
      const slots: Part[] = [];
      if (repeatsCols(c1)) titleWidths.forEach((w, i) => slots.push({ size: w, title: true, i }));
      for (let c = c1; c <= c2; c++) slots.push({ size: widths[c], title: false, i: c });
      const lines: Part[] = [];
      if (repeats(r1)) titleHeights.forEach((h, i) => lines.push({ size: h, title: true, i }));
      for (let r = r1; r <= r2; r++) lines.push({ size: heights[r], title: false, i: r });
      const xs = [x0], ys = [top];
      for (const sl of slots) xs.push(xs[xs.length - 1] + sl.size);
      for (const l of lines) ys.push(ys[ys.length - 1] - l.size);
      const cellAt = (l: Part, sl: Part): PdfCell | null => l.title
        ? (sl.title ? sh.titleCols!.corner(l.i, sl.i) : sh.titles!.cell(l.i, sl.i))
        : (sl.title ? sh.titleCols!.cell(l.i, sl.i) : sh.cells(l.i, sl.i));
      const line = (a: number, b: number, c: number, d: number, colour: string, w = 0.5) => ops.push(`${rgb(colour)} RG ${w} w ${n(a)} ${n(b)} m ${n(c)} ${n(d)} l S`);
      // Fills first, then the gridlines, then the borders and the text.
      lines.forEach((l, li) => slots.forEach((sl, si) => {
        const cell = cellAt(l, sl);
        if (cell?.fill) ops.push(`${rgb(cell.fill)} rg ${n(xs[si])} ${n(ys[li + 1])} ${n(sl.size)} ${n(l.size)} re f`);
      }));
      if (setup.gridlines) {
        for (const x of xs) line(x, ys[0], x, ys[ys.length - 1], "#d0d0d0", 0.25);
        for (const y of ys) line(xs[0], y, xs[xs.length - 1], y, "#d0d0d0", 0.25);
      }
      if (setup.headings) {
        slots.forEach((sl, si) => { const s = colName(sl.title ? sh.titleCols!.index[sl.i] : sheetCol(sl.i)), x = xs[si] + (sl.size - textWidth(s, 8)) / 2; ops.push(`BT /F1 8 Tf 0.4 0.4 0.4 rg ${n(x)} ${n(top + 4)} Td ${pdfText(s)} Tj ET`); });
        lines.forEach((l, li) => { if (l.title) return; const s = String(sheetRow(l.i) + 1), y = ys[li + 1] + (l.size - 8) / 2 + 1; ops.push(`BT /F1 8 Tf 0.4 0.4 0.4 rg ${n(x0 - 4 - textWidth(s, 8))} ${n(y)} Td ${pdfText(s)} Tj ET`); });
      }
      lines.forEach((l, li) => slots.forEach((sl, si) => {
        const cell = cellAt(l, sl);
        if (!cell) return;
        const x = xs[si], y = ys[li + 1], w = sl.size, h = l.size, b = cell.borders ?? {};
        if (b.top) line(x, y + h, x + w, y + h, b.top);
        if (b.bottom) line(x, y, x + w, y, b.bottom);
        if (b.left) line(x, y, x, y + h, b.left);
        if (b.right) line(x + w, y, x + w, y + h, b.right);
        if (!cell.text) return;
        const size = (cell.size ?? 10) * k, font = cell.bold && cell.italic ? "F4" : cell.bold ? "F2" : cell.italic ? "F3" : "F1";
        const tw = textWidth(cell.text, size, cell.bold), tx = cell.align === "end" ? x + w - 2 - tw : cell.align === "center" ? x + (w - tw) / 2 : x + 2;
        // Clipped to its cell, as the grid shows it.
        ops.push(`q ${n(x)} ${n(y)} ${n(w)} ${n(h)} re W n BT /${font} ${n(size)} Tf ${rgb(cell.color ?? "#000000")} rg ${n(tx)} ${n(y + (h - size) / 2 + size * 0.22)} Td ${pdfText(cell.text.split("\n")[0])} Tj ET Q`);
      }));
      streams.push(ops.join("\n"));
      pageInfo.push({ name: sh.name, setup });
      sizes.push([pw, ph]);
    }
  }
  if (!streams.length) { streams.push(""); pageInfo.push({ name: "", setup: base }); sizes.push([PAPER[base.paper][base.landscape ? 1 : 0], PAPER[base.paper][base.landscape ? 0 : 1]]); }
  // The objects: catalog, pages, four fonts, then each page and its content.
  const objs: string[] = [];
  const total = streams.length, fontBase = 3;
  objs[1] = "<< /Type /Catalog /Pages 2 0 R >>";
  const kidsIds = streams.map((_, i) => 7 + i * 2);
  objs[2] = `<< /Type /Pages /Kids [${kidsIds.map((k) => `${k} 0 R`).join(" ")}] /Count ${total} >>`;
  ["Helvetica", "Helvetica-Bold", "Helvetica-Oblique", "Helvetica-BoldOblique"].forEach((f, i) => { objs[fontBase + i] = `<< /Type /Font /Subtype /Type1 /BaseFont /${f} /Encoding /WinAnsiEncoding >>`; });
  streams.forEach((body, i) => {
    const { name, setup } = pageInfo[i], [pw, ph] = sizes[i], file = setup.fileName ?? "";
    const put = (parts: [string, string, string], y: number) => parts.map((t, at) => {
      if (!t) return "";
      const w = textWidth(t, 8), x = at === 0 ? setup.margin : at === 1 ? (pw - w) / 2 : pw - setup.margin - w;
      return `\nBT /F1 8 Tf 0.4 0.4 0.4 rg ${n(x)} ${n(y)} Td ${pdfText(t)} Tj ET`;
    }).join("");
    // The foot is the sheet's name and the page number unless the setup has a footer of its own.
    const foot = setup.footer ? headerParts(setup.footer, i + 1, total, name, file) : ["", `${name ? `${name}  ·  ` : ""}Page ${i + 1} of ${total}`, ""] as [string, string, string];
    const content = `${body}${setup.header ? put(headerParts(setup.header, i + 1, total, name, file), ph - setup.margin / 2 - 4) : ""}${put(foot, setup.margin / 2)}`;
    objs[7 + i * 2] = `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${pw} ${ph}] /Resources << /Font << /F1 3 0 R /F2 4 0 R /F3 5 0 R /F4 6 0 R >> >> /Contents ${8 + i * 2} 0 R >>`;
    objs[8 + i * 2] = `<< /Length ${content.length} >>\nstream\n${content}\nendstream`;
  });
  let pdf = "%PDF-1.4\n%âãÏÓ\n";
  const offsets: number[] = [];
  for (let i = 1; i < objs.length; i++) {
    offsets[i] = pdf.length;
    pdf += `${i} 0 obj\n${objs[i]}\nendobj\n`;
  }
  const xref = pdf.length;
  pdf += `xref\n0 ${objs.length}\n0000000000 65535 f \n${offsets.slice(1).map((o) => `${String(o).padStart(10, "0")} 00000 n \n`).join("")}trailer\n<< /Size ${objs.length} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  // Every character is one byte (WinAnsi), so offsets above are byte offsets.
  const bytes = new Uint8Array(pdf.length);
  for (let i = 0; i < pdf.length; i++) bytes[i] = pdf.charCodeAt(i) & 255;
  return { bytes, pages: total };
}
