// The Tauri spreadsheet's queries (its RangeQuery, in public/lib/
// spreadsheet-functions.js) with no sheet and no language: a table's rows,
// perhaps under a row of headers, filtered, grouped, plucked and
// aggregated. rexx.ts gives RexxJS formulas RANGE, TABLE, WHERE, PLUCK,
// GROUP_BY, SUM, AVG, COUNT and RESULT over these; WHERE's condition is
// whatever the caller works out for a row (RexxJS there), never eval.
import { colName } from "./address.ts";
import { keyOf } from "./groupkey.ts";
import { CellError } from "./values.ts";
import type { Value } from "./values.ts";

export interface Query {
  readonly kind: "query";
  readonly headers: readonly string[] | null;
  readonly rows: readonly (readonly Value[])[];
  readonly firstCol: number;       // the sheet column of the first column, for letters
  // The sheet row (from 1) each row of `rows` came from, so a result says where it came from.
  readonly rowAt: readonly number[];
  // Each group: its key (the first value seen), its rows, and the sheet rows of those.
  readonly groups: readonly (readonly [Value, readonly (readonly Value[])[], readonly number[]])[] | null;
}

const isText = (v: Value) => typeof v === "string" && v !== "";

// A grid as a query. `headers` undefined: found as the app finds them (a
// first row all text, numbers below it); null: none; else these names.
// `firstRow` is the sheet row (from 1) of the grid's first row.
export function queryOf(grid: readonly (readonly Value[])[], firstCol: number, headers?: readonly string[] | null, firstRow = 1): Query {
  let rows = grid, names: readonly string[] | null = headers ?? null, top = firstRow;
  if (headers === undefined && grid.length > 1 && grid[0].every(isText) && grid.slice(1).some((r) => r.some((v) => typeof v === "number"))) {
    names = grid[0].map(String);
    rows = grid.slice(1);
    top++;
  }
  return { kind: "query", headers: names, rows, firstCol, rowAt: rows.map((_, i) => top + i), groups: null };
}

// A column by its header (any case), its sheet letter, or its index.
export function columnIndex(q: Query, ref: string | number): number {
  if (typeof ref === "number") return ref;
  const r = String(ref).trim();
  const byName = q.headers?.findIndex((h) => h.toLowerCase() === r.toLowerCase()) ?? -1;
  if (byName >= 0) return byName;
  const width = q.rows[0]?.length ?? q.headers?.length ?? 0;
  for (let i = 0; i < width; i++) if (colName(q.firstCol + i) === r.toUpperCase()) return i;
  throw new Error(`no column ${r}`);
}

// A row's value by a name as a condition says it: a header, or column_C,
// col_C (a letter or a header); undefined for anything else.
export function cellOf(q: Query, row: readonly Value[]): (name: string) => Value | undefined {
  return (name) => {
    const m = /^col(?:umn)?_(\w+)$/i.exec(name);
    try {
      if (m) return row[columnIndex(q, m[1])] ?? null;
      const i = q.headers?.findIndex((h) => h.toLowerCase() === name.toLowerCase()) ?? -1;
      return i >= 0 ? row[i] ?? null : undefined;
    } catch {
      return undefined;
    }
  };
}

// The rows a test keeps; it is given each row's values by name.
export async function where(q: Query, keep: (col: (name: string) => Value | undefined, row: readonly Value[]) => Promise<boolean>): Promise<Query> {
  const rows: (readonly Value[])[] = [], rowAt: number[] = [];
  for (const [i, row] of q.rows.entries()) if (await keep(cellOf(q, row), row)) { rows.push(row); rowAt.push(q.rowAt[i]); }
  return { ...q, rows, rowAt, groups: null };
}

export const pluck = (q: Query, ref: string | number): Value[][] => {
  const i = columnIndex(q, ref);
  return q.rows.map((r) => [r[i] ?? null]);
};

// Groups by a column's value, in the order they first appear. A key is typed
// (groupkey.ts): 1 and "1" are two groups, text matches without regard to case,
// a blank is a group of its own; a group's label is the first value seen.
export function groupBy(q: Query, ref: string | number): Query {
  const i = columnIndex(q, ref), groups = new Map<string, [Value, (readonly Value[])[], number[]]>();
  for (const [n, row] of q.rows.entries()) {
    const key = row[i] ?? null, k = keyOf(key);
    let g = groups.get(k);
    if (!g) groups.set(k, (g = [key, [], []]));
    g[1].push(row);
    g[2].push(q.rowAt[n]);
  }
  return { ...q, groups: [...groups.values()] };
}

const numbersIn = (rows: readonly (readonly Value[])[], i: number) =>
  rows.map((r) => (typeof r[i] === "number" ? (r[i] as number) : parseFloat(String(r[i])))).filter((n) => !Number.isNaN(n));

function over(kind: "sum" | "avg" | "count", rows: readonly (readonly Value[])[], i: number): number {
  if (kind === "count") return rows.length;
  const ns = numbersIn(rows, i), total = ns.reduce((s, n) => s + n, 0);
  return kind === "sum" ? total : ns.length ? total / ns.length : 0;
}

// SUM, AVG or COUNT: over every row, or each group's key beside its own.
export function aggregate(q: Query, kind: "sum" | "avg" | "count", ref?: string | number): number | [Value, number][] {
  const i = kind === "count" ? 0 : columnIndex(q, ref ?? 0);
  return q.groups ? q.groups.map(([k, rows]) => [k, over(kind, rows, i)]) : over(kind, q.rows, i);
}

// The rows, group by group when grouped.
export const result = (q: Query): Value[][] => (q.groups ? q.groups.flatMap(([, rows]) => rows) : q.rows).map((r) => [...r]);

// Where a result came from: each group's key and the sheet rows of its
// members (an ungrouped query is one group of every row, with no key).
export interface Lineage { readonly key: Value; readonly rows: readonly number[] }
export const lineage = (q: Query): Lineage[] => (q.groups ? q.groups.map(([key, , rows]) => ({ key, rows })) : [{ key: null, rows: q.rowAt }]);

const NUMBER = (n: number) => String(Number(n.toPrecision(12)));
const SHOWN = 8;
const rowsText = (rows: readonly number[]) => `${rows.length === 1 ? "row" : "rows"} ${rows.slice(0, SHOWN).join(", ")}${rows.length > SHOWN ? `, and ${rows.length - SHOWN} more` : ""}`;
export const keyLabel = (k: Value): string => (k === null ? "(blank)" : k instanceof CellError ? k.code : typeof k === "boolean" ? (k ? "TRUE" : "FALSE") : typeof k === "number" ? NUMBER(k) : k === "" ? '""' : k);

// Why a group's aggregate is what it is: the rows it folds, the numbers taken
// from them (the others are skipped, and said so), and the sum, mean or count.
export interface Explained extends Lineage { readonly result: number; readonly text: string }
export function explain(q: Query, kind: "sum" | "avg" | "count", ref?: string | number): Explained[] {
  const i = kind === "count" ? 0 : columnIndex(q, ref ?? 0);
  const column = q.headers?.[i] ?? colName(q.firstCol + i);
  const source = q.groups ?? [[null, q.rows, q.rowAt] as const];
  return source.map(([key, rows, at]) => {
    const head = q.groups ? `${keyLabel(key)}: ` : "";
    const result = over(kind, rows, i);
    if (kind === "count") return { key, rows: at, result, text: `${head}COUNT of ${rowsText(at)}: ${NUMBER(result)}` };
    const ns = numbersIn(rows, i), skipped = rows.length - ns.length;
    const terms = ns.slice(0, SHOWN).map(NUMBER).join(" + ") + (ns.length > SHOWN ? ` + ${ns.length - SHOWN} more` : "");
    const how = !ns.length ? "no numbers" : kind === "sum" ? terms : `(${terms}) / ${ns.length}`;
    const note = skipped ? ` (${skipped} not ${skipped === 1 ? "a number" : "numbers"}, skipped)` : "";
    return { key, rows: at, result, text: `${head}${kind.toUpperCase()} of ${column} over ${rowsText(at)}: ${how} = ${NUMBER(result)}${note}` };
  });
}
