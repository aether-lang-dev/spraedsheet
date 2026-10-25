// Row lineage and grouping explanations (docs/comparisons/LIBREOFFICE.md action
// 4, on the typed keys of VISIDATA.md action 3): a grouped result says which
// source rows each of its groups came from, and why its aggregate is what it is.
// REXX: RANGE(...) |> GROUP_BY("region") |> LINEAGE() / EXPLAIN("sum", "amount");
// the command group.explain (GROUPEXPLAIN from a script) does the same over the
// active sheet's range. Rows are the sheet's own row numbers, so they can be
// read back and edited at the source: the aggregate is a result, not a cell.
import { aggregate, explain, groupBy, lineage, queryOf, where } from "./sheet/query.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import type { Value } from "./sheet/values.ts";
import { eq } from "./lib/check.ts";

const table: Value[][] = [["region", "amount"], ["North", 120], ["south", 80], ["north", 40], [null, 5], ["West", 300]];
// The table sits at sheet rows 10 to 15 (its header on 10).
const q = queryOf(table, 0, undefined, 10);
const g = groupBy(q, "region");
eq("a query knows the sheet row of each of its rows", q.rowAt, [11, 12, 13, 14, 15]);
eq("a group has the sheet rows of its members, in order", lineage(g).map((l) => [l.key, l.rows]), [["North", [11, 13]], ["south", [12]], [null, [14]], ["West", [15]]]);
eq("a filter keeps each row's origin", lineage(groupBy(await where(q, async (col) => (col("amount") as number) > 50), "region")).map((l) => l.rows), [[11], [12], [15]]);
eq("an explanation says what was folded, over which rows, to what", explain(g, "sum", "amount").map((e) => e.text),
  ["North: SUM of amount over rows 11, 13: 120 + 40 = 160", "south: SUM of amount over row 12: 80 = 80", "(blank): SUM of amount over row 14: 5 = 5", "West: SUM of amount over row 15: 300 = 300"]);
eq("COUNT and AVG say so", [explain(g, "count")[0].text, explain(g, "avg", "amount")[0].text], ["North: COUNT of rows 11, 13: 2", "North: AVG of amount over rows 11, 13: (120 + 40) / 2 = 80"]);
eq("the explanation's number is the aggregate's", explain(g, "sum", "amount").map((e) => e.result), (aggregate(g, "sum", "amount") as [Value, number][]).map((a) => a[1]));
eq("a query that is not grouped has one group of every row", lineage(q).map((l) => [l.key, l.rows]), [[null, [11, 12, 13, 14, 15]]]);

// From a sheet.
const wb = new Workbook(), cmds = new Commands(wb);
await cmds.run("formulas.use language=RexxJS");
const set = (cell: string, input: string) => cmds.run(`cell.set cell=${cell} input="${input.replace(/"/g, '""')}"`);
for (const [r, row] of table.entries()) for (const [c, v] of row.entries()) if (v !== null) await set(`${"AB"[c]}${r + 3}`, String(v));
// The header is on row 3, the data on rows 4 to 8.
await set("D1", `=RANGE("A3:B8") |> GROUP_BY("region") |> LINEAGE()`);
eq("LINEAGE spills each group's key, its count and its rows", [0, 1, 2, 3].map((i) => [show(wb.sheet.value(`D${i + 1}`)), show(wb.sheet.value(`E${i + 1}`)), show(wb.sheet.value(`F${i + 1}`))]),
  [["North", "2", "4, 6"], ["south", "1", "5"], ["", "1", "7"], ["West", "1", "8"]]);
await set("H1", `=RANGE("A3:B8") |> GROUP_BY("region") |> EXPLAIN("sum", "amount")`);
eq("EXPLAIN spills each group's key beside the sum it came to", [show(wb.sheet.value("H1")), show(wb.sheet.value("I1"))], ["North", "North: SUM of amount over rows 4, 6: 120 + 40 = 160"]);
await set("B5", "90");
eq("and follows the cells", show(wb.sheet.value("I2")), "south: SUM of amount over row 5: 90 = 90");

const out = await cmds.execute("group.explain", { range: "A3:B8", by: "region", agg: "sum", value: "amount" });
eq("group.explain over a sheet range", [out.ok, (out.output as { key: string; rows: number[]; result: number }[]).map((o) => [o.key, o.rows, o.result])],
  [true, [["North", [4, 6], 160], ["south", [5], 90], ["", [7], 5], ["West", [8], 300]]]);
const one = await cmds.execute("group.explain", { range: "A3:B8", by: "region", agg: "count", key: "west" });
eq("one group by its key (matched without case), and the cells of its rows", [one.message, (one.output as { cells: string[] }[])[0].cells], ["West: COUNT of row 8: 1", ["A8", "B8"]]);
eq("a key that is no group, and a column that is no column", [(await cmds.execute("group.explain", { range: "A3:B8", by: "region", key: "east" })).ok, (await cmds.execute("group.explain", { range: "A3:B8", by: "price" })).message],
  [false, "no column price"]);
const said: string[] = [];
await cmds.script('SAY GROUPEXPLAIN("A3:B8", "region", "sum", "amount", "West")', (l) => said.push(l));
eq("GROUPEXPLAIN from a script", said, ["West: SUM of amount over row 8: 300 = 300"]);

// expect: ok a query knows the sheet row of each of its rows
// expect: ok a group has the sheet rows of its members, in order
// expect: ok a filter keeps each row's origin
// expect: ok an explanation says what was folded, over which rows, to what
// expect: ok COUNT and AVG say so
// expect: ok the explanation's number is the aggregate's
// expect: ok a query that is not grouped has one group of every row
// expect: ok LINEAGE spills each group's key, its count and its rows
// expect: ok EXPLAIN spills each group's key beside the sum it came to
// expect: ok and follows the cells
// expect: ok group.explain over a sheet range
// expect: ok one group by its key (matched without case), and the cells of its rows
// expect: ok a key that is no group, and a column that is no column
// expect: ok GROUPEXPLAIN from a script
