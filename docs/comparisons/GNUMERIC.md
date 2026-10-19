# Gnumeric: actions

**Licence and reuse:** GPL. Ideas only: no numerical routine, test workbook or
website text copied. Check expected values against independent references, not
Gnumeric.

Done, so dropped: distributions, financial, regression, analysis tools, Goal
Seek, and Solver including GRG nonlinear and evolutionary (TODO 6.8b); the
`sheet` CLI for conversion and inspection (7.3).

1. **Function maturity metadata.** For each function: implemented, known
   deviations, reference examples, boundary tests. Generate status docs from
   `described.ts` and the function tables. Shares work with HyperFormula 4.
2. **Numerical acceptance suite.** Extreme tails, inverses near boundaries,
   near-constant data, ill-conditioned regression, financial roots, convergence
   failure. Original fixtures from published datasets or high-precision values;
   today's tests are mostly spot values (e.g. `sheet_functions.ts`). Compare
   independently, since two wrong functions can round-trip.
3. **Declared tolerances and error policy.** Exact for counts and identities,
   justified bounds for approximations; stated NaN, infinity, domain, singular and
   non-convergence results.
4. **Self-explaining analysis reports:** method, assumptions, convergence status;
   never a number that hides a failed fit or infeasible solve.
5. **Machine-readable failures and loss reports from the CLI** (`sheet convert`
   exists; add the structured errors).
