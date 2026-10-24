# HyperFormula: actions

**Licence and reuse:** GPLv3 or commercial. Ideas only: no code, tests or text
copied, and no integration. Not a recommendation to replace the native engine.

Done, so dropped: the full-size sheet (TODO 1.4); a large-file timing test
(7.8, `tests/headless/sheet_large.ts`: 100,000 cells set, rippled, recalculated,
saved, reopened, drawn, all asserted); fuzz and headless-vs-interactive parity
(7.6, 7.7); script grouping with whole-workbook rollback.

1. ~~**Calculation work counters, not just elapsed time.** Count evaluated
   formulas, range visits, settlement rounds and notifications; add running
   totals, diamonds, dynamic references (`OFFSET`/`INDIRECT`) and volatile
   functions to the `sheet_large` workload. Check an unaffected branch is not
   evaluated. Extends 7.8; no step yet.~~ Done 2026-10-24: `Workbook.work()` /
   `resetWork()` count formulas evaluated, cells considered, passes, cells reported
   changed, cells read (a range's every cell) and settle rounds. Headless
   `sheet_work` states the cost of each shape: a chain of 200 (each formula once, an
   unrelated branch of 500 none), a diamond (3), a running total (an edit at row
   600 of 1,000 evaluates 401), a volatile OFFSET (evaluated on any edit), SUM over
   300 cells (once), `SUM($A$1:An)` down 200 rows (200 evaluations, each reading its
   whole range). It found a bug: dependents were collected by recursion, so a
   running total of 1,000 rows overran the engine's call stack ("Maximum call stack
   size exceeded"); it is iterative now. Note `sheet_large` itself is unchanged.
2. ~~**Say what a script reads mid-batch.** Fresh staged results or the committed
   snapshot, when recalculation runs, what a failure restores. Keep today's
   read-after-write on purpose; do not copy HyperFormula's suspended-read rule.
   Document it where scripts are described (`described.ts`).~~ Done 2026-10-21:
   said at the top of `described.ts` and held by headless `sheet_scriptreads`:
   a read after a write sees the written state with formulas worked out; under
   manual calculation a dependent formula keeps its old value until `CALCULATE()`;
   a failing script from the panel keeps what it did (one undo step), a failing
   bus request is taken back whole.
3. **Range-aware aggregates, only if 1 shows a cost.** Floating-point regrouping
   changes results, so give it its own tolerance. Measured 2026-10-24 (the counters
   of 1): a cumulative `SUM($A$1:An)` column re-reads n²/2 cells on an edit at its
   head: 500 rows 0.2 s (125,250 reads), 1,000 rows 0.9 s (500,500), 2,000 rows
   past the 4 s a page entry gets. The cost is real; whether to change sums
   (prefix sums give a differently rounded result) is a decision for Paul, in TODO.md.
4. ~~**One function metadata contract.** Arity, coercion, array lifting,
   volatility, help and locale names agreeing across `described.ts` and the
   function tables, rather than parallel lists. (Shared with Gnumeric's
   maturity labels.)~~ Done 2026-10-24 for arity, kind, volatility, the TRUE/FALSE
   refusal, the Spanish name and where it differs from Excel: `funcinfo.ts`,
   `formulas.functionInfo`, `formulas.functionStatus`, `FUNCTIONINFO()`; headless
   `sheet_funcinfo` checks that the Spanish names, volatile forms, refusals and
   special forms all name functions that exist and that the declared deviations
   are the ledger recounted. Not done: coercion and array lifting are still in
   the function bodies (`lifted()`), not declared; the REXX control functions
   (`described.ts`) were already checked against `control.ts` (`sheet_described`).
5. ~~**A compatibility ledger:** intended Excel/RexxJS differences vs bugs vs
   unsupported vs unverified. IronCalc's PROVENANCE.md "Deliberate differences"
   is the seed.~~ Done 2026-10-24: `COMPATIBILITY.md` here (the four classes, the
   ledger's 388 differing cells counted by kind) and the per-function count in
   `formulas.functionStatus`.
