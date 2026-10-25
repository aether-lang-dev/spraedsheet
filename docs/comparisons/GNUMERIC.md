# Gnumeric: actions

**Licence and reuse:** GPL. Ideas only: no numerical routine, test workbook or
website text copied. Check expected values against independent references, not
Gnumeric.

Done, so dropped: distributions, financial, regression, analysis tools, Goal
Seek, and Solver including GRG nonlinear and evolutionary (TODO 6.8b); the
`sheet` CLI for conversion and inspection (7.3).

1. ~~**Function maturity metadata.** For each function: implemented, known
   deviations, reference examples, boundary tests. Generate status docs from
   `described.ts` and the function tables. Shares work with HyperFormula 4.~~ Done
   in part 2026-10-24 with HyperFormula 4: each function's record (implemented,
   arity, volatile, known deviations as a count of cells that differ from Excel's
   stored values) from `formulas.functionStatus`. Not done: reference examples and
   boundary tests per function, and a generated status document.
2. ~~**Numerical acceptance suite.** Extreme tails, inverses near boundaries,
   near-constant data, ill-conditioned regression, financial roots, convergence
   failure. Original fixtures from published datasets or high-precision values;
   today's tests are mostly spot values (e.g. `sheet_functions.ts`). Compare
   independently, since two wrong functions can round-trip.~~ Done 2026-10-24
   (headless `sheet_numerical`): expected values are closed forms (the chi-square
   tail with 2 degrees of freedom is exp(-x/2), T with 1 is the Cauchy tail, GAMMA
   with shape 1 is 1 - exp(-x)), published datasets (NIST StRD NumAcc1 and NumAcc4,
   Anscombe's quartet) and Microsoft's XIRR example, never another spreadsheet's
   output. It found and fixed: the far tails of the normal and chi-square came out
   0 (`1 - cdf` cancelling: `NORM.S.DIST(-10)` is 7.6e-24, `CHISQ.DIST.RT(100, 2)`
   1.9e-22; the incomplete gamma now works out the small side directly), and
   `NORM.S.INV(1e-10)` and everything that round-trips through the normal was good
   to only 5e-9 (now to the cdf's precision). Four cells of IronCalc's workbooks
   came right with it. Found, not a bug: near-constant data is limited to about
   1e-8 by the doubles for 10000000.1 and its like, so the sd is tested to 1e-7.
3. ~~**Declared tolerances and error policy.** Exact for counts and identities,
   justified bounds for approximations; stated NaN, infinity, domain, singular and
   non-convergence results.~~ Done 2026-10-24 as the header and table of
   `sheet_numerical`: identities exact; closed forms 1e-12 relative; a root-finder
   1e-9; near-constant data 1e-7; and each domain, singular and non-convergence
   case with its error (`NORM.S.INV(0)` #NUM!, `1/0` #DIV/0!, a singular
   `MINVERSE` #NUM!, `IRR` with no sign change #NUM!, overflow #NUM! and never
   Infinity).
4. ~~**Self-explaining analysis reports:** method, assumptions, convergence status;
   never a number that hides a failed fit or infeasible solve.~~ Done 2026-10-24
   (headless `sheet_reports`, `sheet_linest`): goal seek, Solver (infeasible,
   unbounded, answered, the method named) and the regression tool already refused
   with a reason and left the sheet alone, now stated by tests. The gap it found:
   `LINEST` fitted one predictor only and ignored the constant and statistics
   arguments; it now takes several predictors, `const`, and the 5-row statistics
   block, and says "the x columns are collinear" (#NUM!) where Excel would give a
   coefficient made of rounding. `LOGEST` is in (2026-10-25: LINEST on ln y, coefficients exponentiated; headless `sheet_logest`, and the four IronCalc/Excel `LINEST_LOGEST.xlsx` cells that differed now agree).
5. ~~**Machine-readable failures and loss reports from the CLI** (`sheet convert`
   exists; add the structured errors).~~ Done 2026-10-24: `sheet convert|run --json`:
   convert prints one JSON object with `read`, `write` and `report` (what the file
   left out and the writer could not keep), run prints each line and what it gave,
   and a failure is one JSON object on stderr (`ok`, `command`, `line`, `error`,
   `refused`) with the same exit status as before (1 refused, 2 no app). Shell test
   `tests/cli/test_json.sh` (failed first).
