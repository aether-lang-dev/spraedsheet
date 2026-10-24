// LINEST with several predictors, a constant switch and the statistics block
// (docs/comparisons/GNUMERIC.md action 4: a fit that says how good it is). Expected values:
// an exact plane; Anscombe's quartet I (published: slope 0.5001, intercept 3.0001,
// R squared 0.67, residual standard error 1.237, F 17.99 on 9 degrees of freedom,
// regression SS 27.51, residual SS 13.76).
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { CellError } from "./sheet/values.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook(), cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
const near = (a: unknown, b: number, tol = 1e-9) => typeof a === "number" && Math.abs(a - b) <= tol * Math.max(1, Math.abs(b));
const cell = (k: string) => { const v = wb.sheet.value(k); return v instanceof CellError ? v.code : v; };

// y = 1 + 2 x1 + 3 x2, exactly: coefficients come back last predictor first.
const x1 = [1, 2, 3, 4, 5, 6], x2 = [2, 1, 4, 3, 6, 5];
await wb.sheet.setMany(x1.flatMap((v, i): [string, string][] => [[`A${i + 1}`, String(v)], [`B${i + 1}`, String(x2[i])], [`C${i + 1}`, String(1 + 2 * v + 3 * x2[i])]]));
await wb.sheet.set("E1", "=LINEST(C1:C6, A1:B6)");
eq("two predictors: x2's coefficient, x1's, then the intercept", [near(cell("E1"), 3), near(cell("F1"), 2), near(cell("G1"), 1)], [true, true, true]);
await wb.sheet.set("E3", "=LINEST(C1:C6, A1:B6, FALSE)");
eq("through the origin (constant FALSE): no intercept, a 0", cell("G3"), 0);

// Anscombe's quartet I with the statistics block.
const xs = [10, 8, 13, 9, 11, 14, 6, 4, 12, 7, 5], ys = [8.04, 6.95, 7.58, 8.81, 8.33, 9.96, 7.24, 4.26, 10.84, 4.82, 5.68];
await wb.sheet.setMany(xs.flatMap((x, i): [string, string][] => [[`J${i + 1}`, String(x)], [`K${i + 1}`, String(ys[i])]]));
await wb.sheet.set("M1", "=LINEST(K1:K11, J1:J11, TRUE, TRUE)");
eq("Anscombe I: slope and intercept", [near(cell("M1"), 0.5000909, 1e-6), near(cell("N1"), 3.0000909, 1e-6)], [true, true]);
eq("their standard errors (0.1179, 1.1247)", [near(cell("M2"), 0.1179055, 1e-5), near(cell("N2"), 1.1247468, 1e-5)], [true, true]);
eq("R squared 0.667 and the standard error of y 1.237", [near(cell("M3"), 0.6665425, 1e-6), near(cell("N3"), 1.2366034, 1e-5)], [true, true]);
eq("F 17.99 on 9 degrees of freedom", [near(cell("M4"), 17.98994, 1e-5), cell("N4")], [true, 9]);
eq("regression SS 27.51 and residual SS 13.76", [near(cell("M5"), 27.51, 1e-4), near(cell("N5"), 13.76269, 1e-4)], [true, true]);

// What cannot be fitted says so.
await wb.sheet.setMany(x1.map((v, i): [string, string] => [`P${i + 1}`, String(v)]).concat(x1.map((v, i): [string, string] => [`Q${i + 1}`, String(v)])));
await wb.sheet.set("S1", "=LINEST(C1:C6, P1:Q6)");
const v = wb.sheet.value("S1");
eq("collinear predictors are #NUM! and say so, not a confident coefficient", [v instanceof CellError ? v.code : "no error", v instanceof CellError ? v.why : ""], ["#NUM!", "LINEST: the x columns are collinear"]);
await wb.sheet.set("S3", "=LINEST(C1:C3, A1:B6)");
eq("y and x of different lengths are #REF!", cell("S3"), "#REF!");

// expect: ok two predictors: x2's coefficient, x1's, then the intercept
// expect: ok through the origin (constant FALSE): no intercept, a 0
// expect: ok Anscombe I: slope and intercept
// expect: ok their standard errors (0.1179, 1.1247)
// expect: ok R squared 0.667 and the standard error of y 1.237
// expect: ok F 17.99 on 9 degrees of freedom
// expect: ok regression SS 27.51 and residual SS 13.76
// expect: ok collinear predictors are #NUM! and say so, not a confident coefficient
// expect: ok y and x of different lengths are #REF!
