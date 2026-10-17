// Solver's GRG Nonlinear and Evolutionary (6.8b), on textbook problems:
// Rosenbrock's valley (min (1 - x)² + 100(y - x²)²: 1, 1, from -1.2, 1);
// a constrained maximum (max xy, x + y ≤ 10: 5, 5, 25); a circle's nearest
// point to (3, 4) (min (x-3)² + (y-4)², x² + y² = 1: 0.6, 0.8); and on the
// sheet: its formulas worked out at each point, Evolutionary on a step
// function no gradient helps (a knapsack of integers), its bounds needed,
// the Answer Report naming the engine, SOLVER from REXX.
import { evolutionary, grg } from "./sheet/nlp.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const near = (x: readonly number[], want: readonly number[], tol: number) => x.every((v, i) => Math.abs(v - want[i]) <= tol);
const free2 = { lower: [-Infinity, -Infinity], upper: [Infinity, Infinity] };
const rosen = await grg({ n: 2, maximise: false, rels: [], ...free2, start: [-1.2, 1], evaluate: async ([x, y]) => ({ obj: (1 - x) ** 2 + 100 * (y - x * x) ** 2, diff: [] }) });
eq("GRG: Rosenbrock's valley", [rosen.status, near(rosen.x, [1, 1], 1e-3)], ["optimal", true]);
const box = await grg({ n: 2, maximise: true, rels: ["<="], lower: [0, 0], upper: [Infinity, Infinity], start: [1, 1], evaluate: async ([x, y]) => ({ obj: x * y, diff: [x + y - 10] }) });
eq("GRG: max xy with x + y ≤ 10", [box.status, near(box.x, [5, 5], 1e-3), box.status !== "infeasible" && Math.abs(box.obj - 25) < 1e-2], ["optimal", true, true]);
const circle = await grg({ n: 2, maximise: false, rels: ["="], ...free2, start: [1, 0], evaluate: async ([x, y]) => ({ obj: (x - 3) ** 2 + (y - 4) ** 2, diff: [x * x + y * y - 1] }) });
eq("GRG: the circle's point nearest (3, 4), an equality", [circle.status, near(circle.x, [0.6, 0.8], 1e-3)], ["optimal", true]);
const rastrigin = await evolutionary({ n: 2, maximise: false, rels: [], lower: [-5, -5], upper: [5, 5], start: [4, 4], evaluate: async (x) => ({ obj: 20 + x.reduce((s, v) => s + v * v - 10 * Math.cos(2 * Math.PI * v), 0), diff: [] }) });
eq("Evolutionary: Rastrigin's many valleys, to the global one", [rastrigin.status, near(rastrigin.x, [0, 0], 0.05)], ["converged", true]);
const again = await evolutionary({ n: 2, maximise: false, rels: [], lower: [-5, -5], upper: [5, 5], start: [4, 4], evaluate: async (x) => ({ obj: 20 + x.reduce((s, v) => s + v * v - 10 * Math.cos(2 * Math.PI * v), 0), diff: [] }) });
eq("seeded: the same answer twice", again.x, rastrigin.x);

// On the sheet.
const wb = new Workbook();
const cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
await cmds.run("cell.set cell=B1 input=1");
await cmds.run("cell.set cell=B2 input=1");
await cmds.run('cell.set cell=B3 input="=B1*B2"');
await cmds.run('cell.set cell=B4 input="=B1+B2"');
const o = await cmds.run('solver.solve objective=B3 goal=max changing=B1:B2 constraints="B4<=10" method=grg report=1');
eq("GRG on the sheet: max B1*B2 with B1+B2 ≤ 10", [o.ok, o.message.startsWith("Solver (GRG Nonlinear) found a solution: B3 = "), Math.round(Number(wb.sheet.value("B3")) * 1000) / 1000, Math.round(Number(wb.sheet.value("B1")) * 1000) / 1000],
  [true, true, 25, 5]);
eq("its Answer Report names the engine", [wb.page("Answer Report 1")?.sheet.input("B2"), show(wb.page("Answer Report 1")!.sheet.value("A1"))], ["GRG Nonlinear", "Answer Report"]);
eq("Simplex LP still refuses it", (await cmds.run('solver.solve objective=B3 goal=max changing=B1:B2 constraints="B4<=10"')).message, "the problem is not linear in B1, B2: Simplex LP needs linear formulas");
// A knapsack: items of weight 3, 4, 5 and value 4, 5, 6, at most 2 of each, weight ≤ 10: 1 + 0 + ... best 13 (one of each weighing 12 is too heavy).
for (const [k, v] of [["D1", "0"], ["D2", "0"], ["D3", "0"], ["E1", '"=3*D1+4*D2+5*D3"'], ["E2", '"=4*D1+5*D2+6*D3"'], ["E3", '"=IF(E1<=10,E2,0)"']]) await cmds.run(`cell.set cell=${k} input=${v}`);
const ks = await cmds.run('solver.solve objective=E3 goal=max changing=D1:D3 constraints="D1:D3<=2;D1:D3=int;E1<=10" method=evolutionary');
eq("Evolutionary on a step function of integers", [ks.ok, show(wb.sheet.value("E3")), [1, 2, 3].map((i) => Number.isInteger(Number(wb.sheet.value(`D${i}`))))], [true, "13", [true, true, true]]);
eq("Evolutionary needs bounds; GRG takes no int", [(await cmds.run("solver.solve objective=B3 goal=max changing=B1:B2 method=evolutionary")).message, (await cmds.run('solver.solve objective=E3 changing=D1:D3 constraints="D1:D3=int" method=grg')).message, (await cmds.run("solver.solve objective=B3 changing=B1 method=newton")).message],
  ["Evolutionary needs bounds on every changing cell, as Excel's (B1, B2 has none above)", "GRG Nonlinear here takes no int or bin constraint: use Evolutionary, or Simplex LP for a linear problem", "newton is not a solving method (simplex, grg, evolutionary)"]);
const said: string[] = [];
await cmds.script('CALL SETCELL "G1", 0\nCALL SETCELL "G2", "=(G1-3)^2+1"\nSAY SOLVER("G2", "min", "G1", "", "method=grg")', (l) => said.push(l));
eq("from REXX", [said[0].startsWith("Solver (GRG Nonlinear) found a solution: G2 = "), Math.round(Number(wb.sheet.value("G2")) * 1000) / 1000, Math.round(Number(wb.sheet.value("G1")) * 1000) / 1000], [true, 1, 3]);

// expect: ok GRG: Rosenbrock's valley
// expect: ok GRG: max xy with x + y ≤ 10
// expect: ok GRG: the circle's point nearest (3, 4), an equality
// expect: ok Evolutionary: Rastrigin's many valleys, to the global one
// expect: ok seeded: the same answer twice
// expect: ok GRG on the sheet: max B1*B2 with B1+B2 ≤ 10
// expect: ok its Answer Report names the engine
// expect: ok Simplex LP still refuses it
// expect: ok Evolutionary on a step function of integers
// expect: ok Evolutionary needs bounds; GRG takes no int
// expect: ok from REXX
