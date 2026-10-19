# HyperFormula: actions

**Licence and reuse:** GPLv3 or commercial. Ideas only: no code, tests or text
copied, and no integration. Not a recommendation to replace the native engine.

Done, so dropped: the full-size sheet (TODO 1.4); a large-file timing test
(7.8, `tests/headless/sheet_large.ts`: 100,000 cells set, rippled, recalculated,
saved, reopened, drawn, all asserted); fuzz and headless-vs-interactive parity
(7.6, 7.7); script grouping with whole-workbook rollback.

1. **Calculation work counters, not just elapsed time.** Count evaluated
   formulas, range visits, settlement rounds and notifications; add running
   totals, diamonds, dynamic references (`OFFSET`/`INDIRECT`) and volatile
   functions to the `sheet_large` workload. Check an unaffected branch is not
   evaluated. Extends 7.8; no step yet.
2. **Say what a script reads mid-batch.** Fresh staged results or the committed
   snapshot, when recalculation runs, what a failure restores. Keep today's
   read-after-write on purpose; do not copy HyperFormula's suspended-read rule.
   Document it where scripts are described (`described.ts`).
3. **Range-aware aggregates, only if 1 shows a cost.** Floating-point regrouping
   changes results, so give it its own tolerance.
4. **One function metadata contract.** Arity, coercion, array lifting,
   volatility, help and locale names agreeing across `described.ts` and the
   function tables, rather than parallel lists. (Shared with Gnumeric's
   maturity labels.)
5. **A compatibility ledger:** intended Excel/RexxJS differences vs bugs vs
   unsupported vs unverified. IronCalc's PROVENANCE.md "Deliberate differences"
   is the seed.
