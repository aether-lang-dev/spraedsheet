// Computed criteria for the database functions (DSUM, DCOUNT, DGET ...): a criteria cell that holds a
// formula, under a header that is no field of the database (a label, or empty), is worked out for each
// record, its relative references to the first record's cells moving down to that record's, and the
// record is kept when it is TRUE (or not 0), as Excel does (docs/comparisons/gridcraft_parity.md).
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook(), cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
const set = async (k: string, v: string) => { const o = await cmds.run(`cell.set cell=${k} input="${v.replace(/"/g, '""')}"`); if (!o.ok) print(`refused ${k}: ${o.message}`); };
const rows: [string, number, number][] = [["Apple", 18, 14], ["Pear", 12, 10], ["Cherry", 13, 9], ["Apple", 14, 10], ["Pear", 9, 8], ["Apple", 8, 6]];
for (const [c, h] of [["A", "Tree"], ["B", "Height"], ["C", "Yield"]]) await set(`${c}5`, h);
for (const [i, [t, h, y]] of rows.entries()) { await set(`A${6 + i}`, t); await set(`B${6 + i}`, String(h)); await set(`C${6 + i}`, String(y)); }
await set("J1", "12");
// The criteria: a label over a formula (taller than the average), and Tree beside it for the second case.
await set("E1", "Taller"); await set("E2", "=B6>AVERAGE($B$6:$B$11)");
await set("G1", "Tree"); await set("H1", "Check"); await set("G2", "Apple"); await set("H2", "=B6>$J$1");
await set("G3", "Pear");
await set("L1", "Tall"); await set("L2", "=B6>$J$1");
const at = (k: string) => show(wb.sheet.value(k));

await set("N1", '=DSUM(A5:C11, "Yield", E1:E2)');
await set("N2", "=DCOUNT(A5:C11, , E1:E2)");
await set("N3", '=DSUM(A5:C11, "Yield", G1:H3)');
await set("N4", '=DGET(A5:C11, "Yield", L1:L2)');
await set("N5", '=DAVERAGE(A5:C11, "Height", E1:E2)');
eq("a formula under a label: the records taller than the average (18, 13, 14)", [at("N1"), at("N2"), at("N5")], ["33", "3", "15"]);
eq("beside a text criterion in its row (and), with another row (or)", at("N3"), "42");
eq("DGET of a computed criterion that matches several records is #NUM!", at("N4"), "#NUM!");

// What it reads outside the database moves it too: the threshold in J1.
await set("J1", "17");
eq("the cell it reads changes: only the 18 is over 17", [at("N4"), at("N3")], ["14", "32"]);
await set("B7", "20");
eq("and a record changes", [at("N1"), at("N2")], ["34", "3"]);
// expect: ok a formula under a label: the records taller than the average (18, 13, 14)
// expect: ok beside a text criterion in its row (and), with another row (or)
// expect: ok DGET of a computed criterion that matches several records is #NUM!
// expect: ok the cell it reads changes: only the 18 is over 17
// expect: ok and a record changes
