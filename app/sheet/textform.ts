// The canonical text form of a workbook (blog-entry-todo.md B.1): the file's
// JSON (file.ts's writeSheetFile) as lines that diff well. One cell per
// line, sheet by sheet, rows then columns:
//
//   # rexxsheet text 1
//   @formulas<TAB>"excel"                 a workbook setting: JSON after the tab
//   @macros<TAB>Total<TAB>"..."           an object's entry, sorted by key
//   @sheet<TAB>Sheet1
//   A1<TAB>Item                           a cell: its input as typed
//   B2<TAB>=SUM(A1:A9)<TAB>{"format":"0.00"}   a formula as typed, no value
//   @names<TAB>QTY<TAB>"B2:B4"            a sheet's section, an entry per line
//   @rules[]<TAB>{...}                    a list's element, in the file's order
//
// A cell's second field is its input unquoted, or a JSON string where that
// would not read back (a tab, a newline, a leading quote or blank). A third
// field, when there is one, is JSON: what the cell has beyond its input (a
// comment, a number format, a look that is not the default, validation).
// A constant has no value beside its input; a formula's cached value is not
// kept (it is worked out again as the file is opened). JSON is written with
// its keys sorted, so the same workbook is the same bytes.
import { addrOf } from "./address.ts";

export const TEXT_HEADER = "# rexxsheet text 1";

// The parts of a sheet in the file's JSON; any other key belongs to the workbook.
const SHEET_KEYS = ["cells", "names", "rules", "charts", "view", "merges", "tables", "slicers", "scenarios", "drawings", "protection", "editRanges", "autoId", "pivots"];
// Lists whose order means nothing, written sorted so a move is not a change.
// Objects that mean something empty (protected, with no password).
const KEEP_EMPTY = ["bookProtection"];
const SORTED: Record<string, (x: any) => string> = { merges: (x) => String(x), tables: (t) => String(t?.name ?? "") };

function canon(v: unknown): unknown {
  if (Array.isArray(v)) return v.map(canon);
  if (v && typeof v === "object") return Object.fromEntries(Object.keys(v).sort().map((k) => [k, canon((v as Record<string, unknown>)[k])]));
  return v;
}
const json = (v: unknown): string => JSON.stringify(canon(v));

// Raw where that reads back as it is, else a JSON string.
const field = (s: string): string => (/[\u0000-\u001f]/.test(s) || s.startsWith('"') || s !== s.trim() ? JSON.stringify(s) : s);
const unfield = (s: string): string => (s.startsWith('"') ? JSON.parse(s) : s);

function sheetLines(sheet: Record<string, any>, out: string[]): void {
  const cells: Record<string, any> = sheet.cells ?? {};
  const keys = Object.keys(cells).sort((a, b) => { const x = addrOf(a)!, y = addrOf(b)!; return x.row - y.row || x.col - y.col; });
  for (const key of keys) {
    const c = cells[key];
    if (typeof c === "string") { out.push(`${key}\t${field(c)}`); continue; }
    const { value, expression, content, ...extras } = c as Record<string, unknown>;
    const input = expression ? `=${expression}` : String(content ?? value ?? "");
    const rest = Object.fromEntries(Object.entries(extras).filter(([, v]) => v !== null && v !== undefined));
    out.push(Object.keys(rest).length ? `${key}\t${field(input)}\t${json(rest)}` : `${key}\t${field(input)}`);
  }
  const parts = Object.fromEntries(Object.entries(sheet).filter(([k]) => SHEET_KEYS.includes(k) && k !== "cells"));
  // The view is written only where it differs from the default (no hidden
  // rows, nothing frozen, no sizes set); the reader supplies the rest.
  if (parts.view) parts.view = Object.fromEntries(Object.entries(parts.view).filter(([, v]) => !(v === 0 || (Array.isArray(v) && v.length === 0))));
  partsLines(parts, out);
}

// A workbook's or a sheet's parts: a scalar on a line, a list an element
// apiece, an object an entry apiece (sorted).
function partsLines(parts: Record<string, any>, out: string[]): void {
  for (const key of Object.keys(parts).sort()) {
    const v = parts[key];
    if (Array.isArray(v)) {
      const sort = SORTED[key];
      const items = sort ? [...v].sort((a, b) => (sort(a) < sort(b) ? -1 : sort(a) > sort(b) ? 1 : 0)) : v;
      for (const item of items) out.push(`@${key}[]\t${json(item)}`);
    } else if (v && typeof v === "object") {
      if (Object.keys(v).length === 0 && KEEP_EMPTY.includes(key)) out.push(`@${key}\t{}`);
      for (const k of Object.keys(v).sort()) out.push(`@${key}\t${field(k)}\t${json(v[k])}`);
    } else {
      out.push(`@${key}\t${json(v)}`);
    }
  }
}

// The text of a workbook's file (writeSheetFile's value).
export function textOfFile(file: Record<string, any>): string {
  const out: string[] = [TEXT_HEADER];
  const book = Object.fromEntries(Object.entries(file).filter(([k]) => !SHEET_KEYS.includes(k) && k !== "sheetName" && k !== "sheets"));
  partsLines(book, out);
  out.push(`@sheet\t${field(String(file.sheetName ?? "Sheet1"))}`);
  sheetLines(file, out);
  for (const s of (file.sheets ?? []) as Record<string, any>[]) {
    out.push(`@sheet\t${field(String(s.name))}`);
    sheetLines(s, out);
  }
  return `${out.join("\n")}\n`;
}

// The file (as readSheetFile takes it) a text is; a malformed one throws a
// RangeError naming its line.
export function fileOfText(text: string): Record<string, any> {
  const lines = text.split("\n").map((l) => (l.endsWith("\r") ? l.slice(0, -1) : l));
  if (lines[0] !== TEXT_HEADER) throw new RangeError(`not a workbook text: the first line is not "${TEXT_HEADER}"`);
  const book: Record<string, any> = {};
  const sheets: Record<string, any>[] = [];
  let target: Record<string, any> = book;
  const bad = (n: number, why: string): never => { throw new RangeError(`line ${n + 1}: ${why}`); };
  for (let n = 1; n < lines.length; n++) {
    const line = lines[n];
    if (line === "" || line.startsWith("#")) continue;
    const f = line.split("\t");
    try {
      if (f[0] === "@sheet") {
        if (f.length !== 2) bad(n, "@sheet takes one name");
        target = { cells: {}, name: unfield(f[1]) };
        sheets.push(target);
      } else if (f[0].startsWith("@")) {
        const key = f[0].slice(1);
        if (key.endsWith("[]")) {
          if (f.length !== 2) bad(n, "a list element has one value");
          const k = key.slice(0, -2);
          (target[k] ??= []).push(JSON.parse(f[1]));
        } else if (f.length === 3) {
          (target[key] ??= {})[unfield(f[1])] = JSON.parse(f[2]);
        } else if (f.length === 2) {
          target[key] = JSON.parse(f[1]);
        } else bad(n, "a setting has one value, an entry a key and a value");
      } else {
        if (f.length < 2 || f.length > 3 || !addrOf(f[0]) || target === book) bad(n, "not a cell line (a cell, a tab, its input) under a @sheet");
        const input = unfield(f[1]);
        target.cells[f[0]] = f.length === 3 ? { value: input, ...(input.startsWith("=") ? { expression: input.slice(1) } : {}), ...JSON.parse(f[2]) } : input;
      }
    } catch (e) {
      if (e instanceof SyntaxError) bad(n, e.message);
      throw e;
    }
  }
  if (!sheets.length) throw new RangeError("a workbook text has no @sheet");
  const [first, ...others] = sheets;
  const { name: firstName, ...firstParts } = first;
  return { ...book, sheetName: firstName, ...firstParts, sheets: others };
}
