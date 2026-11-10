// References as values of formulas: the range operator between a cell and a call that makes a
// reference (A1:OFFSET(A1, 2, 0), A1:INDEX(B:B, 3), CHOOSE(...):CHOOSE(...)), a LET name or
// LAMBDA parameter that stays a reference, IF and CHOOSE working out only the branch taken (a
// cell named in the other is no cycle), and the postfix % (IronCalc's range_operator.xlsx,
// CHOOSE.xlsx, ISREF.xlsx, percentage.xlsx, MINUS_MINUS.xlsx).
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook(), sh = wb.sheet;
const at = (k: string): string => show(sh.value(k));
for (const [k, v] of [["A1", "1"], ["A2", "2"], ["A3", "3"], ["A4", "4"], ["B1", "10"], ["B2", "20"], ["B3", "30"]]) await sh.set(k, v);

// The range operator, and what it reads when a cell between changes.
await sh.set("D1", "=SUM(A1:OFFSET(A1, 2, 0))");
await sh.set("D2", "=SUM(A1:INDEX(A1:A4, 3))");
await sh.set("D3", "=SUM(OFFSET(A1, 0, 0):OFFSET(A1, 1, 1))");
await sh.set("D4", "=SUM(CHOOSE(1, A1, A2):CHOOSE(2, A3, A4))");
await sh.set("D5", "=SUM(INDIRECT(\"A1\"):INDIRECT(\"B2\"))");
eq("the range operator between references", [at("D1"), at("D2"), at("D3"), at("D4"), at("D5")], ["6", "6", "33", "10", "33"]);
await sh.set("A2", "100");
eq("a cell between the corners changed", [at("D1"), at("D2"), at("D3"), at("D4"), at("D5")], ["104", "104", "131", "108", "131"]);
await sh.set("D6", "=A1:INDEX(A1:A4, 2)");
eq("a reference spills as its values", [at("D6"), at("D7")], ["1", "100"]);
await sh.set("M1", "=A1:OFFSET(A1, 1, 1)");
eq("a rectangle from corner to corner", [at("M1"), at("N1"), at("M2"), at("N2")], ["1", "10", "100", "20"]);

// IF and CHOOSE: the cell's own name in the branch not taken is no cycle.
await sh.set("F1", "=CHOOSE(2, F1, \"Not a cycle\")");
await sh.set("F2", "=IF(TRUE, \"fine\", F2)");
await sh.set("F3", "=IF(TRUE, F4, 0)");
await sh.set("F4", "=F3");
eq("a cell in the branch not taken", [at("F1"), at("F2")], ["Not a cycle", "fine"]);
eq("a cycle through the branch taken is one", [at("F3"), at("F4")], ["#CYCLE!", "#CYCLE!"]);

// A reference held by a name.
await sh.set("H1", "=LET(x, A2:A3, MIN(ROW(x)) + COLUMN(x))");
await sh.set("H2", "=LAMBDA(r, SUM(r) + ROWS(r))(A1:A3)");
await sh.set("H3", "=LET(x, 1 / 0, y, 5, y)");
eq("a name holding a reference, an error not raised until used", [at("H1"), at("H2"), at("H3")], ["3", "107", "5"]);

// The postfix %.
await sh.set("J1", "=50%");
await sh.set("J2", "=A1*50%");
await sh.set("J3", "=SUM(B1:B3%)");
await sh.set("J4", "=-A4%");
await sh.set("J5", "=2^50%");
await sh.set("J6", "=200*5%%");
await sh.set("J7", "=10%+10%");
await sh.set("K1", "=20%*SUM(B1:B3)");
eq("the postfix %", [at("J1"), at("J2"), at("J3"), at("J4"), at("J5"), at("J6"), at("J7"), at("K1")], ["0.5", "0.5", "0.6", "-0.04", "1.414213562", "0.1", "0.2", "12"]);

// SINGLE(x) is @x, as a file keeps it.
await sh.set("Q2", "=SINGLE(B1:B3)");
await sh.set("Q3", "=SINGLE(B:B) + 1");
eq("SINGLE is the implicit intersection", [at("Q2"), at("Q3")], ["20", "31"]);

// An error beside an array is that error in each cell, which IFERROR answers one by one.
await sh.set("R1", "=SUM(IFERROR(A1:A3 / NA(), 2))");
await sh.set("R2", "=SUM(IFNA(A1:A3 / NA(), 1))");
await sh.set("R3", "=SUM(IFERROR(A1:A3 / {1; 0; 1}, 100))");
eq("IFERROR and IFNA over an array", [at("R1"), at("R2"), at("R3")], ["6", "3", "104"]);

// An array's empty cells are 0 where it lands, and its errors sort in Excel's order.
await sh.set("S1", "5");
await sh.set("S3", "=NA()");
await sh.set("S4", "=1/0");
await sh.set("U1", "=SORT(S1:S4)");
await sh.set("V1", "=SORT(S1:S4, 1, -1)");
await sh.set("W1", "=S1:S4");
eq("SORT: errors in Excel's order, the empty cell last and 0", [at("U1"), at("U2"), at("U3"), at("U4")], ["5", "#DIV/0!", "#N/A", "0"]);
eq("SORT descending", [at("V1"), at("V2"), at("V3"), at("V4")], ["#N/A", "#DIV/0!", "5", "0"]);
eq("an empty cell in a spilled range is 0", [at("W1"), at("W2"), at("W3"), at("W4")], ["5", "0", "#N/A", "#DIV/0!"]);

// expect: ok the range operator between references
// expect: ok a cell between the corners changed
// expect: ok a reference spills as its values
// expect: ok a rectangle from corner to corner
// expect: ok a cell in the branch not taken
// expect: ok a cycle through the branch taken is one
// expect: ok a name holding a reference, an error not raised until used
// expect: ok the postfix %
// expect: ok SINGLE is the implicit intersection
// expect: ok IFERROR and IFNA over an array
// expect: ok SORT: errors in Excel's order, the empty cell last and 0
// expect: ok SORT descending
// expect: ok an empty cell in a spilled range is 0
