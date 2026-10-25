// What makes two rows one group (docs/comparisons/VISIDATA.md action 3): a
// key is a value with its type, not its text. Used by GROUP_BY (query.ts) and
// by GROUPBY and PIVOTBY (groupby.ts), so they agree.
//
//   number     by value, -0 as 0: 1 and "1" are two groups
//   text       without regard to case, as `=` compares it; a group is labelled
//              by the spelling first seen
//   boolean    TRUE and FALSE, apart from the text "TRUE"
//   blank      one group of its own: not 0, not "" (`=` equates a blank with
//              both, but 0 and "" are not equal, so as an equality it would
//              depend on which row came first)
//   error      by its code (#N/A is one group, whatever the reason)
//
// Rows of one key are in one group however far apart they are, and the groups
// come in the order their first row does.
import { CellError } from "./values.ts";
import type { Value } from "./values.ts";

export function keyOf(v: Value): string {
  if (v === null) return "b";
  if (v instanceof CellError) return `e${v.code}`;
  if (typeof v === "number") return `n${v === 0 ? 0 : v}`;
  if (typeof v === "boolean") return v ? "tT" : "tF";
  if (typeof v === "string") return `s${v.toUpperCase()}`;
  return `o${String(v)}`;
}

// A row of key fields as one key.
export const keyOfRow = (values: readonly Value[]): string => values.map(keyOf).join("\u0001");
