// Protection (Review > Protect Sheet, Protect Workbook, Allow Edit Ranges):
// a sheet's cells are locked unless unlocked (Format Cells > Protection),
// which only matters once the sheet is protected; then a locked cell
// refuses a change, as do what the protection does not allow (formatting,
// inserting or deleting rows and columns, sorting, filtering, objects ...),
// and an edit range lets its cells be changed, at once or after its own
// password. A protected workbook's structure refuses sheets added,
// deleted, renamed, moved, hidden or copied. Passwords are kept as Excel's
// legacy 16-bit hash (ECMA-376 Part 4, 14.7.1's password verifier), which
// is what a file's sheetProtection password="..." holds: protection is a
// guard against slips, not secrecy (that is an encrypted file's).
import { parseRange } from "./address.ts";
import type { Addr } from "./address.ts";

// What a protected sheet may still have done to it (Excel's Protect Sheet
// check boxes, by their attribute names in a file). Selecting cells is
// always allowed here.
export const ALLOWS = ["formatCells", "formatColumns", "formatRows", "insertColumns", "insertRows", "insertHyperlinks", "deleteColumns", "deleteRows", "sort", "autoFilter", "pivotTables", "objects", "scenarios"] as const;
export type Allow = typeof ALLOWS[number];

export interface EditRange {
  readonly title: string;
  readonly range: readonly [Addr, Addr];
  readonly hash?: string;
}

export interface BookProtection {
  readonly hash?: string;
}

export interface SheetProtection {
  readonly hash?: string;
  readonly allow: readonly Allow[];
}

// Excel's legacy password hash, as four hex digits ("CC1A"); "" for no
// password.
export function passwordHash(password: string): string {
  if (!password) return "";
  let h = 0;
  const cs = [...password].map((c) => c.charCodeAt(0) & 0xff).slice(0, 15);
  for (let i = cs.length - 1; i >= 0; i--) {
    h = ((h >> 14) & 1) | ((h << 1) & 0x7fff);
    h ^= cs[i];
  }
  h = ((h >> 14) & 1) | ((h << 1) & 0x7fff);
  h ^= cs.length;
  h ^= 0xce4b;
  return h.toString(16).toUpperCase().padStart(4, "0");
}

// Whether a password opens what a hash guards (no hash: any).
export const opens = (hash: string | undefined, password: string): boolean => !hash || passwordHash(password) === hash.toUpperCase();

export const within = (a: Addr, r: readonly [Addr, Addr]): boolean =>
  a.col >= Math.min(r[0].col, r[1].col) && a.col <= Math.max(r[0].col, r[1].col) && a.row >= Math.min(r[0].row, r[1].row) && a.row <= Math.max(r[0].row, r[1].row);

// The allow list from "formatCells,sort" (or "" for none), each checked.
export function allowsOf(text: string): Allow[] {
  const out: Allow[] = [];
  for (const w of text.split(/[,;\s]+/).filter((x) => x)) {
    const a = ALLOWS.find((x) => x.toLowerCase() === w.toLowerCase());
    if (!a) throw new RangeError(`${w} is not something a protected sheet allows (${ALLOWS.join(", ")})`);
    if (!out.includes(a)) out.push(a);
  }
  return out;
}

export const rangeText = (s: string): [Addr, Addr] | null => parseRange(s);

// What a command needs of a protected sheet: one of its allowances (then
// it may run), "free" (it changes nothing a protection guards: the view,
// the workbook's names, a file, undo), or "cells" (it may run, and is
// taken back if it changed a locked cell or anything else of the sheet).
const NEEDS: Readonly<Record<string, Allow | "free" | "never">> = {
  "rows.insert": "insertRows", "columns.insert": "insertColumns", "rows.delete": "deleteRows", "columns.delete": "deleteColumns",
  "data.sort": "sort", "data.sortAscending": "sort", "data.sortDescending": "sort", "table.sort": "sort",
  "data.filter": "autoFilter", "data.clearFilter": "autoFilter", "data.advancedFilter": "autoFilter", "slicer.select": "autoFilter", "slicer.clear": "autoFilter",
  "pivot.create": "pivotTables", "pivot.options": "pivotTables", "pivot.update": "pivotTables", "pivot.delete": "pivotTables", "chart.pivot": "pivotTables",
  "scenario.add": "scenarios", "scenario.show": "scenarios", "scenario.delete": "scenarios", "scenario.summary": "scenarios",
  "chart.insert": "objects", "chart.option": "objects", "chart.trendline": "objects", "chart.data": "objects", "chart.delete": "objects",
  "insert.picture": "objects", "insert.shape": "objects", "insert.textbox": "objects", "drawing.edit": "objects", "drawing.delete": "objects", "drawing.order": "objects",
  "insert.sparkline": "objects", "cell.comment": "objects", "range.clearComments": "objects",
  "comment.new": "objects", "comment.reply": "objects", "comment.resolve": "objects", "comment.delete": "objects", "comment.convertNotes": "objects",
  "cell.hyperlink": "insertHyperlinks", "cell.removeHyperlink": "insertHyperlinks",
  "rows.height": "formatRows", "rows.hide": "formatRows", "rows.unhide": "formatRows", "rows.autofit": "formatRows",
  "columns.width": "formatColumns", "columns.hide": "formatColumns", "columns.unhide": "formatColumns", "columns.autofit": "formatColumns",
  "edit.copy": "free", "edit.undo": "free", "edit.redo": "free", "names.define": "free", "names.delete": "free",
  "view.freezeRows": "free", "view.freezeCols": "free", "view.freezePanes": "free", "validation.circle": "free",
  "formulas.calculation": "free", "formulas.iteration": "free", "formulas.use": "free", "setup.set": "free", "setup.run": "free", "workbook.dateSystem": "free",
  "theme.set": "free", "theme.color": "free", "theme.fonts": "free", "style.define": "free", "style.delete": "free",
  "sheet.add": "free", "sheet.activate": "free", "sheet.rename": "free", "sheet.move": "free", "sheet.color": "free", "sheet.hide": "free", "sheet.copy": "free", "sheet.delete": "free",
  "cells.merge": "never", "cells.unmerge": "never", "table.define": "never", "table.delete": "never", "outline.group": "never", "outline.ungroup": "never",
  "data.subtotal": "never", "validation.set": "never", "validation.clear": "never",
};

// A formatting command: what Format Cells allows (formatCells), its
// changes to a cell's look and number format let through, not to its value.
export const formatting = (id: string): boolean =>
  id.startsWith("format.") || id === "cell.format" || id === "style.apply" || id === "range.clearFormats";

export const needOf = (id: string): Allow | "free" | "never" | "cells" =>
  id.startsWith("view.") || id.startsWith("file.") || id.startsWith("protect.") || id.startsWith("comment.author") ? "free" : NEEDS[id] ?? (formatting(id) ? "formatCells" : "cells");

export const PROTECTED = "The cell or chart you're trying to change is on a protected sheet. To make a change, unprotect the sheet (Review > Unprotect Sheet). You might be asked for a password.";
export const STRUCTURE = "The workbook's structure is protected: sheets cannot be added, deleted, moved, renamed, hidden or copied (Review > Protect Workbook to unprotect it).";
