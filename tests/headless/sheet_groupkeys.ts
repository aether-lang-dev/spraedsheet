// Typed group keys (docs/comparisons/VISIDATA.md action 3): what makes two
// rows one group. A key is its value with its type, not its text: the number 1
// and the text "1" are two groups, a blank cell is a group of its own (neither
// 0 nor ""), text matches without regard to case as `=` does (labelled by the
// spelling first seen), -0 is 0, an error is its code, and a group is the
// same whatever order the rows come in (the old GROUPBY compared a row with a
// group's first row only, and a blank equals 0 and "" but those two do not
// equal each other, so the answer depended on the order).
import { aggregate, groupBy, queryOf, result } from "./sheet/query.ts";
import { keyOf } from "./sheet/groupkey.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { CellError } from "./sheet/values.ts";
import type { Value } from "./sheet/values.ts";
import { eq } from "./lib/check.ts";

eq("a key keeps its type: 1, \"1\", TRUE, \"TRUE\", blank and \"\" are six", new Set([1, "1", true, "TRUE", null, ""].map(keyOf)).size, 6);
eq("text without regard to case, -0 as 0, an error by its code", [keyOf("North") === keyOf("NORTH"), keyOf(-0) === keyOf(0), keyOf(new CellError("#N/A", "a")) === keyOf(new CellError("#N/A", "b")), keyOf(new CellError("#N/A", "a")) === keyOf(new CellError("#REF!", "a"))], [true, true, true, false]);

const rows = (...keys: Value[]): Value[][] => keys.map((k, i) => [k, i + 1]);
const q = (...keys: Value[]) => queryOf(rows(...keys), 0, ["k", "v"]);
const sums = (...keys: Value[]) => aggregate(groupBy(q(...keys), "k"), "sum", "v");
eq("GROUP_BY: 1 and \"1\" are two groups, summed apart", sums(1, "1", 1), [[1, 4], ["1", 2]]);
eq("a group is labelled by the first spelling seen, matched without case", sums("North", "south", "NORTH", "South"), [["North", 4], ["south", 6]]);
eq("a blank is its own group, not 0 and not \"\"", sums(null, 0, "", null), [[null, 5], [0, 2], ["", 3]]);
eq("booleans are not their text", sums(true, "TRUE", true), [[true, 4], ["TRUE", 2]]);
eq("errors group by code", sums(new CellError("#N/A", "x"), new CellError("#N/A", "y"), new CellError("#REF!", "z")).map(([k, n]) => [(k as CellError).code, n]), [["#N/A", 3], ["#REF!", 3]]);
eq("-0 and 0 are one group", sums(0, -0), [[0, 3]]);
const g = groupBy(q("a", "b", "A"), "k");
eq("the rows of each group are in the groups, in order, and RESULT lists group by group", [g.groups!.map(([k, r]) => [k, r.length]), result(g).map((r) => r[1])], [[["a", 2], ["b", 1]], [1, 3, 2]]);

// Order makes no difference: the same keys in any order make the same groups.
{
  const keys: Value[] = [null, 0, "", false, "a", "A", 1, "1"];
  const counts = (ks: Value[]) => {
    const m = new Map<string, number>();
    for (const k of ks) m.set(keyOf(k), (m.get(keyOf(k)) ?? 0) + 1);
    return [...m.values()].sort();
  };
  const reversed = [...keys].reverse();
  eq("the same keys, reversed, make the same groups", [counts(keys), counts(reversed)], [counts(keys), counts(keys)]);
}

// The GROUPBY and PIVOTBY functions use the same keys.
const wb = new Workbook(), cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
const put = async (cell: string, input: string) => { await cmds.run(`cell.set cell=${cell} input="${input.replace(/"/g, '""')}"`); };
await put("A1", "1"); await put("A2", "'1"); await put("A3", "1"); await put("A4", "x"); await put("A5", "X");
for (let i = 1; i <= 5; i++) await put(`B${i}`, String(i));
await put("D1", "=GROUPBY(A1:A5, B1:B5, SUM, 0, 0)");
eq("GROUPBY: the number 1 and the text 1 apart, X with x", [0, 1, 2].map((r) => `${show(wb.sheet.value(`D${r + 1}`))}=${show(wb.sheet.value(`E${r + 1}`))}`), ["1=4", "1=2", "x=9"]);
await put("D6", "=ROWS(GROUPBY(A1:A5, B1:B5, SUM, 0, 0))");
eq("and three groups", show(wb.sheet.value("D6")), "3");
await put("G1", "=PIVOTBY(A1:A5, A1:A5, B1:B5, SUM, 0, 0, 0, 0)");
eq("PIVOTBY: the same three keys down and across", [show(wb.sheet.value("G2")), show(wb.sheet.value("G3")), show(wb.sheet.value("G4"))], ["1", "1", "x"]);

// expect: ok a key keeps its type: 1, "1", TRUE, "TRUE", blank and "" are six
// expect: ok text without regard to case, -0 as 0, an error by its code
// expect: ok GROUP_BY: 1 and "1" are two groups, summed apart
// expect: ok a group is labelled by the first spelling seen, matched without case
// expect: ok a blank is its own group, not 0 and not ""
// expect: ok booleans are not their text
// expect: ok errors group by code
// expect: ok -0 and 0 are one group
// expect: ok the rows of each group are in the groups, in order, and RESULT lists group by group
// expect: ok the same keys, reversed, make the same groups
// expect: ok GROUPBY: the number 1 and the text 1 apart, X with x
// expect: ok and three groups
// expect: ok PIVOTBY: the same three keys down and across
