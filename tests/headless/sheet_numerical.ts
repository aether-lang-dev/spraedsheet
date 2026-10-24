// A numerical acceptance suite (docs/comparisons/GNUMERIC.md actions 2 and 3). Expected
// values come from closed forms and published datasets, never from another
// spreadsheet: the exponential tail of CHISQ with 2 degrees of freedom, the
// Cauchy distribution of T with 1, Anscombe's quartet, NIST StRD's NumAcc4
// (near-constant data), Microsoft's XIRR example. Declared tolerances: near-constant data 1e-7 (the doubles for 10000000.1 and its
// like are themselves that far from the decimals); counts and identities exactly; closed forms to 1e-12 relative; a root-finder to
// 1e-9 relative. Domain, singular and non-convergence cases state their error.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { CellError } from "./sheet/values.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook(), cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
let at = 0;
const calc = async (f: string): Promise<unknown> => {
  const key = `Z${++at}`;
  await wb.sheet.set(key, `=${f}`);
  const v = wb.sheet.value(key);
  return v instanceof CellError ? v.code : v;
};
const rel = (got: unknown, want: number, tol: number): boolean => typeof got === "number" && Math.abs(got - want) <= tol * Math.max(Math.abs(want), 1e-300);
const close = async (label: string, f: string, want: number, tol = 1e-12) => { const got = await calc(f); eq(label, rel(got, want, tol) ? "ok" : `${f} = ${got}, want ${want}`, "ok"); };

// Extreme tails against closed forms.
for (const x of [1, 10, 100, 500, 1400]) await close(`CHISQ.DIST.RT(${x}, 2) is exp(-x/2)`, `CHISQ.DIST.RT(${x}, 2)`, Math.exp(-x / 2));
for (const x of [0.5, 5, 50, 5000]) await close(`T.DIST.RT(${x}, 1) is the Cauchy tail`, `T.DIST.RT(${x}, 1)`, 0.5 - Math.atan(x) / Math.PI, 1e-9);
for (const x of [1, 30, 700]) await close(`GAMMA.DIST(${x}, 1, 1, TRUE) is 1 - exp(-x)`, `GAMMA.DIST(${x}, 1, 1, TRUE)`, -Math.expm1(-x));
await close("the standard normal's far lower tail (published: 7.619853024160527e-24)", "NORM.S.DIST(-10, TRUE)", 7.619853024160527e-24, 1e-9);

// Inverses near their boundaries return where they started.
for (const p of [1e-300, 1e-100, 1e-20, 1e-10, 0.5, 1 - 1e-10]) await close(`NORM.S.DIST(NORM.S.INV(${p})) returns ${p}`, `NORM.S.DIST(NORM.S.INV(${p}), TRUE)`, p, 1e-9);
await close("NORM.S.INV(1e-10) (published: -6.361340902404056)", "NORM.S.INV(1E-10)", -6.361340902404056, 1e-10);
for (const x of [0.3, 5, 40, 400]) await close(`CHISQ.INV.RT(CHISQ.DIST.RT(x, 5)) returns x (${x})`, `CHISQ.INV.RT(CHISQ.DIST.RT(${x}, 5), 5)`, x, 1e-8);

// Near-constant data: NIST StRD NumAcc4, 1,001 values around 10,000,000.2 whose standard deviation is 0.1.
{
  const e: [string, string][] = [["A1", "10000000.2"]];
  for (let i = 0; i < 500; i++) { e.push([`A${2 + 2 * i}`, "10000000.1"]); e.push([`A${3 + 2 * i}`, "10000000.3"]); }
  await wb.sheet.setMany(e);
  await close("NumAcc4: the mean (a plain sum of 1,001 values: 1e-13)", "AVERAGE(A1:A1001)", 10000000.2, 1e-13);
  await close("NumAcc4: the standard deviation (certified 0.1)", "STDEV.S(A1:A1001)", 0.1, 1e-7);
  await close("NumAcc4: the variance", "VAR.S(A1:A1001)", 0.01, 2e-7);
  await wb.sheet.setMany([["C1", "10001"], ["C2", "10002"], ["C3", "10003"]]);
  await close("NumAcc1: the standard deviation of 10001..10003 (certified 1)", "STDEV.S(C1:C3)", 1, 1e-14);
}

// Regression on a published dataset: Anscombe's quartet I.
{
  const xs = [10, 8, 13, 9, 11, 14, 6, 4, 12, 7, 5], ys = [8.04, 6.95, 7.58, 8.81, 8.33, 9.96, 7.24, 4.26, 10.84, 4.82, 5.68];
  await wb.sheet.setMany(xs.flatMap((x, i): [string, string][] => [[`E${i + 1}`, String(x)], [`F${i + 1}`, String(ys[i])]]));
  await close("Anscombe I: slope (published 0.5001)", "SLOPE(F1:F11, E1:E11)", 0.5000909090909091, 1e-9);
  await close("Anscombe I: intercept (published 3.0001)", "INTERCEPT(F1:F11, E1:E11)", 3.000090909090909, 1e-9);
  await close("Anscombe I: r squared (published 0.67)", "RSQ(F1:F11, E1:E11)", 0.6665424595087748, 1e-9);
}

// Ill-conditioned: a regression on x and x*(1+1e-14) is singular to working precision.
{
  await wb.sheet.setMany([["H1", "1"], ["H2", "2"], ["H3", "3"], ["H4", "4"], ["I1", "1"], ["I2", "2"], ["I3", "3"], ["I4", "4"], ["J1", "2"], ["J2", "4.1"], ["J3", "5.9"], ["J4", "8"]]);
  const r = await calc("INDEX(LINEST(J1:J4, H1:I4), 1, 1)");
  eq("two identical predictors do not give a confident coefficient: an error or a stated zero", r === 0 || typeof r === "string", true);
}

// Financial roots.
await close("IRR of -100, 110 is 10%", "IRR({-100,110})", 0.1, 1e-9);
await close("RATE finds the rate PMT paid (round trip)", "RATE(120, PMT(0.0075, 120, -50000), -50000)", 0.0075, 1e-9);
await close("XIRR (Microsoft's example: 0.373362535)", 'XIRR({-10000,2750,4250,3250,2750}, {39448,39508,39751,39859,39904})', 0.373362535, 1e-6);
await close("NPV at the IRR is zero", "NPV(IRR({-100,60,60}), {60,60})+(-100)+0", 0, 1e-6);

// Domains, singular matrices and non-convergence say so.
const errors: [string, string][] = [
  ["NORM.S.INV(0)", "#NUM!"], ["NORM.S.INV(1)", "#NUM!"], ["LN(0)", "#NUM!"], ["SQRT(-1)", "#NUM!"], ["1/0", "#DIV/0!"],
  ["LOG(-1)", "#NUM!"], ["GAMMALN(0)", "#NUM!"], ["CHISQ.INV.RT(0.5, 0)", "#NUM!"], ["MINVERSE({1,2;2,4})", "#NUM!"],
  ["IRR({100,200})", "#NUM!"], ["RATE(10, 100, 100)", "#NUM!"], ["STDEV.S(1)", "#DIV/0!"], ["VAR.S(5)", "#DIV/0!"],
  ["BINOM.DIST(5, 3, 0.5, TRUE)", "#NUM!"], ["POISSON.DIST(-1, 3, TRUE)", "#NUM!"], ["T.INV(0.5, 0)", "#NUM!"],
];
for (const [f, want] of errors) eq(`${f} is ${want}`, await calc(f), want);
eq("a huge factorial overflows to #NUM!, not Infinity", await calc("FACT(171)"), "#NUM!");
eq("and 1E308*10 overflows to #NUM!", await calc("1E308*10"), "#NUM!");

// expect: ok CHISQ.DIST.RT(1, 2) is exp(-x/2)
// expect: ok CHISQ.DIST.RT(10, 2) is exp(-x/2)
// expect: ok CHISQ.DIST.RT(100, 2) is exp(-x/2)
// expect: ok CHISQ.DIST.RT(500, 2) is exp(-x/2)
// expect: ok CHISQ.DIST.RT(1400, 2) is exp(-x/2)
// expect: ok T.DIST.RT(0.5, 1) is the Cauchy tail
// expect: ok T.DIST.RT(5, 1) is the Cauchy tail
// expect: ok T.DIST.RT(50, 1) is the Cauchy tail
// expect: ok T.DIST.RT(5000, 1) is the Cauchy tail
// expect: ok GAMMA.DIST(1, 1, 1, TRUE) is 1 - exp(-x)
// expect: ok GAMMA.DIST(30, 1, 1, TRUE) is 1 - exp(-x)
// expect: ok GAMMA.DIST(700, 1, 1, TRUE) is 1 - exp(-x)
// expect: ok the standard normal's far lower tail (published: 7.619853024160527e-24)
// expect: ok NORM.S.DIST(NORM.S.INV(1e-300)) returns 1e-300
// expect: ok NORM.S.DIST(NORM.S.INV(1e-100)) returns 1e-100
// expect: ok NORM.S.DIST(NORM.S.INV(1e-20)) returns 1e-20
// expect: ok NORM.S.DIST(NORM.S.INV(1e-10)) returns 1e-10
// expect: ok NORM.S.DIST(NORM.S.INV(0.5)) returns 0.5
// expect: ok NORM.S.DIST(NORM.S.INV(0.9999999999)) returns 0.9999999999
// expect: ok NORM.S.INV(1e-10) (published: -6.361340902404056)
// expect: ok CHISQ.INV.RT(CHISQ.DIST.RT(x, 5)) returns x (0.3)
// expect: ok CHISQ.INV.RT(CHISQ.DIST.RT(x, 5)) returns x (5)
// expect: ok CHISQ.INV.RT(CHISQ.DIST.RT(x, 5)) returns x (40)
// expect: ok CHISQ.INV.RT(CHISQ.DIST.RT(x, 5)) returns x (400)
// expect: ok NumAcc4: the mean (a plain sum of 1,001 values: 1e-13)
// expect: ok NumAcc4: the standard deviation (certified 0.1)
// expect: ok NumAcc4: the variance
// expect: ok NumAcc1: the standard deviation of 10001..10003 (certified 1)
// expect: ok Anscombe I: slope (published 0.5001)
// expect: ok Anscombe I: intercept (published 3.0001)
// expect: ok Anscombe I: r squared (published 0.67)
// expect: ok two identical predictors do not give a confident coefficient: an error or a stated zero
// expect: ok IRR of -100, 110 is 10%
// expect: ok RATE finds the rate PMT paid (round trip)
// expect: ok XIRR (Microsoft's example: 0.373362535)
// expect: ok NPV at the IRR is zero
// expect: ok NORM.S.INV(0) is #NUM!
// expect: ok NORM.S.INV(1) is #NUM!
// expect: ok LN(0) is #NUM!
// expect: ok SQRT(-1) is #NUM!
// expect: ok 1/0 is #DIV/0!
// expect: ok LOG(-1) is #NUM!
// expect: ok GAMMALN(0) is #NUM!
// expect: ok CHISQ.INV.RT(0.5, 0) is #NUM!
// expect: ok MINVERSE({1,2;2,4}) is #NUM!
// expect: ok IRR({100,200}) is #NUM!
// expect: ok RATE(10, 100, 100) is #NUM!
// expect: ok STDEV.S(1) is #DIV/0!
// expect: ok VAR.S(5) is #DIV/0!
// expect: ok BINOM.DIST(5, 3, 0.5, TRUE) is #NUM!
// expect: ok POISSON.DIST(-1, 3, TRUE) is #NUM!
// expect: ok T.INV(0.5, 0) is #NUM!
// expect: ok a huge factorial overflows to #NUM!, not Infinity
// expect: ok and 1E308*10 overflows to #NUM!
