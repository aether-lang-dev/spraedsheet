// LOGEST: the exponential fit y = b * m1^x1 * m2^x2 ..., LINEST on ln y with the
// coefficients exponentiated (the standard errors and the rest of the statistics
// stay those of the ln fit, as in Excel). Expected values are closed forms:
// y = 3 * 2^x is {2, 3}; y = 2 * 3^x1 * 5^x2 is {5, 3, 2}; the statistics block
// of an inexact fit is LINEST's on ln y (itself checked against Anscombe's
// quartet in sheet_linest).
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { CellError } from "./sheet/values.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook(), cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
const near = (a: unknown, b: number, tol = 1e-9) => typeof a === "number" && Math.abs(a - b) <= tol * Math.max(1, Math.abs(b));
const cell = (k: string) => { const v = wb.sheet.value(k); return v instanceof CellError ? v.code : v; };

// y = 3 * 2^x exactly: the base m = 2 first, then the multiplier b = 3.
const x = [1, 2, 3, 4, 5];
await wb.sheet.setMany(x.flatMap((v, i): [string, string][] => [[`A${i + 1}`, String(v)], [`B${i + 1}`, String(3 * 2 ** v)]]));
await wb.sheet.set("D1", "=LOGEST(B1:B5, A1:A5)");
eq("y = 3 * 2^x: m = 2, b = 3", [near(cell("D1"), 2), near(cell("E1"), 3)], [true, true]);
await wb.sheet.set("D3", "=LOGEST(B1:B5)");
eq("no known_x: x is 1, 2, 3 ...", [near(cell("D3"), 2), near(cell("E3"), 3)], [true, true]);

// Through the origin (const FALSE): y = m^x, b = 1.
await wb.sheet.setMany([["G1", "2"], ["G2", "4"], ["G3", "8"], ["G4", "16"], ["H1", "1"], ["H2", "2"], ["H3", "3"], ["H4", "4"]]);
await wb.sheet.set("J1", "=LOGEST(G1:G4, H1:H4, FALSE)");
eq("const FALSE: y = m^x, b is 1", [near(cell("J1"), 2), near(cell("K1"), 1)], [true, true]);

// Two predictors: y = 2 * 3^x1 * 5^x2, the last predictor's base first.
const p1 = [1, 2, 3, 4, 5, 6], p2 = [2, 1, 4, 3, 6, 5];
await wb.sheet.setMany(p1.flatMap((v, i): [string, string][] => [[`M${i + 1}`, String(v)], [`N${i + 1}`, String(p2[i])], [`O${i + 1}`, String(2 * 3 ** v * 5 ** p2[i])]]));
await wb.sheet.set("Q1", "=LOGEST(O1:O6, M1:N6)");
eq("two predictors: 5, 3, then 2", [near(cell("Q1"), 5), near(cell("R1"), 3), near(cell("S1"), 2)], [true, true, true]);

// An inexact fit with the statistics: row 1 is exp of LINEST's on ln y, the rest are LINEST's.
const ys = [2.9, 6.3, 11.8, 25.1, 47.6, 101];
await wb.sheet.setMany(ys.flatMap((y, i): [string, string][] => [[`U${i + 1}`, String(i + 1)], [`V${i + 1}`, String(y)], [`W${i + 1}`, `=LN(V${i + 1})`]]));
await wb.sheet.set("Y1", "=LOGEST(V1:V6, U1:U6, TRUE, TRUE)");
await wb.sheet.set("AB1", "=LINEST(W1:W6, U1:U6, TRUE, TRUE)");
const same = [] as boolean[];
for (const [r, c] of [[1, 0], [1, 1]]) same.push(near(cell(`${["Y", "Z"][c]}${r}`), Math.exp(cell(`${["AB", "AC"][c]}${r}`) as number), 1e-12));
eq("the coefficients are exp of the ln fit's", same, [true, true]);
const rest = [] as boolean[];
for (let r = 2; r <= 5; r++) for (let c = 0; c < 2; c++) {
  const a = cell(`${["Y", "Z"][c]}${r}`), b = cell(`${["AB", "AC"][c]}${r}`);
  rest.push(typeof b === "number" ? near(a, b, 1e-12) : a === b);
}
eq("standard errors, R squared, F, df and the sums of squares are the ln fit's", rest, rest.map(() => true));

// What cannot be fitted says so.
await wb.sheet.setMany([["AE1", "4"], ["AE2", "0"], ["AE3", "16"], ["AF1", "1"], ["AF2", "2"], ["AF3", "3"]]);
await wb.sheet.set("AH1", "=LOGEST(AE1:AE3, AF1:AF3)");
const v = wb.sheet.value("AH1");
eq("a y of zero or less is #NUM!, with its reason", [v instanceof CellError ? v.code : "no error", v instanceof CellError ? v.why : ""], ["#NUM!", "LOGEST takes y values above zero"]);
await wb.sheet.set("AH3", "=LOGEST(B1:B3, A1:A5)");
eq("y and x of different lengths are #REF!", cell("AH3"), "#REF!");

// expect: ok y = 3 * 2^x: m = 2, b = 3
// expect: ok no known_x: x is 1, 2, 3 ...
// expect: ok const FALSE: y = m^x, b is 1
// expect: ok two predictors: 5, 3, then 2
// expect: ok the coefficients are exp of the ln fit's
// expect: ok standard errors, R squared, F, df and the sums of squares are the ln fit's
// expect: ok a y of zero or less is #NUM!, with its reason
// expect: ok y and x of different lengths are #REF!
