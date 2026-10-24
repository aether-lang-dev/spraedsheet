// Analysis that explains itself (docs/comparisons/GNUMERIC.md action 4): a goal seek, a
// solve or a regression that cannot be answered says so, and why, and leaves the
// sheet as it was; one that is answered says which method found it and what the
// answer is. Never a number that hides a failed fit or an infeasible solve.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { CellError } from "./sheet/values.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook(), cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
await wb.sheet.setMany([["A1", "2"], ["B1", "=A1^2+1"], ["C1", "1"], ["C2", "1"], ["D1", "=C1+C2"], ["E1", "=C1*C2"],
  ["P1", "2"], ["P2", "4.1"], ["P3", "5.9"], ["P4", "8"], ["Q1", "1"], ["Q2", "2"], ["Q3", "3"], ["Q4", "4"], ["R1", "1"], ["R2", "2"], ["R3", "3"], ["R4", "4"]]);

// Goal seek: B1 = A1^2 + 1 can never be 0.
const gs = await cmds.execute("data.goalSeek", { set: "B1", to: "0", by: "A1" });
eq("a goal seek with no answer says so and puts the cell back", [gs.ok, gs.message, wb.sheet.input("A1")], [false, "B1 cannot reach 0 by changing A1", "2"]);
eq("one with an answer says what it found", (await cmds.execute("data.goalSeek", { set: "B1", to: "10", by: "A1" })).message, "B1 = 10 with A1 = 3");

// Solver: infeasible, unbounded, answered (each method named).
const bad = await cmds.execute("solver.solve", { objective: "D1", changing: "C1:C2", goal: "max", constraints: "C1<=5;C2<=7;C1>=6" });
eq("an infeasible solve says so and leaves the cells", [bad.ok, bad.message, wb.sheet.input("C1"), wb.sheet.input("C2")], [false, "Solver could not find a feasible solution", "1", "1"]);
const unb = await cmds.execute("solver.solve", { objective: "D1", changing: "C1:C2", goal: "max" });
eq("an unbounded one says so", [unb.ok, unb.message, wb.sheet.input("C1")], [false, "the objective's cells do not converge: the problem is unbounded", "1"]);
const lp = await cmds.execute("solver.solve", { objective: "D1", changing: "C1:C2", goal: "max", constraints: "C1<=5;C2<=7" });
eq("an answered solve gives the objective's value and the cells", [lp.ok, lp.message, lp.output], [true, "Solver found a solution: D1 = 12 at C1=5, C2=7", [5, 7]]);
const grg = await cmds.execute("solver.solve", { objective: "E1", changing: "C1:C2", goal: "max", constraints: "C1<=5;C2<=7", method: "grg" });
eq("a nonlinear one names its method", [grg.ok, grg.message], [true, "Solver (GRG Nonlinear) found a solution: E1 = 35 at C1=5, C2=7"]);

// Regression: collinear predictors are refused, not fitted; one-variable functions say why.
const reg = await cmds.execute("data.analyze", { tool: "Regression", input: "P1:R4", output: "T1" });
eq("a regression on collinear predictors says so and writes nothing", [reg.ok, reg.message, wb.sheet.value("T1")], [false, "the x columns are collinear", null]);
await wb.sheet.setMany([["K1", "5"], ["K2", "5"], ["K3", "5"], ["K4", "5"]]);
const why = async (f: string) => { await wb.sheet.set("L1", `=${f}`); const v = wb.sheet.value("L1"); return v instanceof CellError ? `${v.code}: ${v.why}` : String(v); };
eq("a fit on x values that are all alike is #DIV/0! and says it", [await why("SLOPE(P1:P4,K1:K4)"), await why("FORECAST(2,P1:P4,K1:K4)")],
  ["#DIV/0!: SLOPE: the x values are all alike", "#DIV/0!: FORECAST.LINEAR: the x values are all alike"]);
eq("LINEST on collinear predictors is #NUM! and says so", await why("INDEX(LINEST(P1:P4,Q1:R4),1,1)"), "#NUM!: LINEST: the x columns are collinear");

// expect: ok a goal seek with no answer says so and puts the cell back
// expect: ok one with an answer says what it found
// expect: ok an infeasible solve says so and leaves the cells
// expect: ok an unbounded one says so
// expect: ok an answered solve gives the objective's value and the cells
// expect: ok a nonlinear one names its method
// expect: ok a regression on collinear predictors says so and writes nothing
// expect: ok a fit on x values that are all alike is #DIV/0! and says it
// expect: ok LINEST on collinear predictors is #NUM! and says so
