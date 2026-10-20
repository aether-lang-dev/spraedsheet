// A workbook as a REXX program (blog-entry-todo.md B.2): the file's JSON
// (file.ts's writeSheetFile) as the calls that make it again, the way
// chaess's recorder writes a game as the script that replays it. Readable
// (a cell is `CALL SETCELL "B2", "=B1*21"`) and runnable: in a new workbook
// it gives the same one, whose text form (textform.ts) is the same.
//
//   /* Budget: ... */
//   ADDRESS SHEET "file.begin json=..."          the head: language, setup script, theme
//   ADDRESS SHEET "sheet.begin name=Sheet1 json=..."   the sheet there, its names defined
//   CALL SETCELL "A1", "Item"                    a cell, rows then columns
//   ADDRESS SHEET "cell.note cell=A1 json=..."   what it has besides its input
//   ADDRESS SHEET "sheet.parts name=Sheet1 json=..."   rules, charts, view, merges, tables ...
//   ADDRESS SHEET "file.end json=..."            the workbook's names, macros, protection
//
// These are file.ts's open() in phases; each json= is the file's own shape.
import { addrOf } from "./address.ts";
import { literal as rexxLiteral } from "./recorder.ts";

const quotedParam = (v: string) => `"${v.replace(/"/g, '""')}"`;
// A command line to ADDRESS SHEET: an id and its parameters (a JSON one
// quoted, a quote doubled). The line is a plain double-quoted string, which
// RexxJS passes on as it is (a CALL's argument has escapes; this has none).
const command = (id: string, params: Record<string, string>): string =>
  `ADDRESS SHEET "${`${id} ${Object.entries(params).map(([k, v]) => `${k}=${quotedParam(v)}`).join(" ")}`.replace(/"/g, '""')}"`;

const SHEET_KEYS = ["cells", "names", "rules", "charts", "view", "merges", "tables", "slicers", "scenarios", "drawings", "protection", "editRanges", "autoId", "pivots"];
const empty = (v: unknown) => v === 0 || (Array.isArray(v) && v.length === 0) || (v && typeof v === "object" && !Array.isArray(v) && Object.keys(v).length === 0);
const keep = (o: Record<string, any>) => Object.fromEntries(Object.entries(o).filter(([, v]) => !empty(v)));

export function bookScript(file: Record<string, any>): string {
  const sheets: Record<string, any>[] = [{ ...file, name: file.sheetName }, ...((file.sheets ?? []) as Record<string, any>[])];
  const out: string[] = [`/* ${String(file.name).replace(/\*\//g, "* /")}: this workbook as a REXX program that makes it again (WORKBOOKSCRIPT) */`];
  const head = Object.fromEntries(["formulas", "setupScript", "sheetName", "dateSystem", "cellStyles", "theme", "calculation"].filter((k) => file[k] !== undefined).map((k) => [k, file[k]]));
  out.push(command("file.begin", { json: JSON.stringify(head) }));
  for (const s of sheets) {
    out.push(command("sheet.begin", { name: s.name, json: JSON.stringify(keep({ names: s.names ?? {} })) }));
    const cells: Record<string, any> = s.cells ?? {};
    for (const key of Object.keys(cells).sort((a, b) => { const x = addrOf(a)!, y = addrOf(b)!; return x.row - y.row || x.col - y.col; })) {
      const c = cells[key];
      if (typeof c === "string") { out.push(`CALL SETCELL ${rexxLiteral(key)}, ${rexxLiteral(c)}`); continue; }
      const { value, expression, content, ...extras } = c as Record<string, unknown>;
      const input = expression ? `=${expression}` : String(content ?? value ?? "");
      if (input !== "") out.push(`CALL SETCELL ${rexxLiteral(key)}, ${rexxLiteral(input)}`);
      const rest = Object.fromEntries(Object.entries(extras).filter(([, v]) => v !== null && v !== undefined));
      if (Object.keys(rest).length) out.push(command("cell.note", { cell: key, json: JSON.stringify(rest) }));
    }
    // The parts but the filter (worked out in file.end, once every sheet's cells are).
    const { filter, ...view } = (s.view ?? {}) as Record<string, any>;
    const parts = keep({ ...Object.fromEntries(Object.entries(s).filter(([k]) => SHEET_KEYS.includes(k) && k !== "cells" && k !== "names" && k !== "view")), view: keep(view) });
    if (Object.keys(parts).length) out.push(command("sheet.parts", { name: s.name, json: JSON.stringify(parts) }));
  }
  const end: Record<string, any> = Object.fromEntries(Object.entries(file).filter(([k]) => !SHEET_KEYS.includes(k) && k !== "sheets"));
  if (file.view?.filter) end.view = { filter: file.view.filter };
  const filtered = ((file.sheets ?? []) as Record<string, any>[]).filter((s) => s.view?.filter).map((s) => ({ name: s.name, view: { filter: s.view.filter } }));
  if (filtered.length) end.sheets = filtered;
  out.push(command("file.end", { json: JSON.stringify(end) }));
  return `${out.join("\n")}\n`;
}
