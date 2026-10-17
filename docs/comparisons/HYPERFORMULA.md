# HyperFormula compared with Spraedsheet

**Most useful lesson: make calculation cost and batch behaviour explicit parts
of the engine contract.** Spraedsheet already has a headless model and dependency
calculation; the opportunity is to strengthen and measure them.

Review: `/home/paul/scm/hyperformula`, revision `187dfa1`. Baseline and limits:
[comparison index](../../OPEN_SOURCE_COMPARISONS.md).

## What is present

HyperFormula is a TypeScript calculation engine with no required grid UI. It
provides formula parsing/evaluation, workbook edits, named expressions,
custom functions, undo/redo, and batches. The checkout contains a dependency
graph with cell and range vertices, topological/SCC processing, address mapping,
formula transformations, and change-result objects.

Its dependency-graph guide explains why ranges deserve their own representation:
running totals over progressively larger ranges can otherwise create quadratic
dependency storage. Shared range structure and cached aggregates can reduce
repeated work. That is an optimisation opportunity to measure, not evidence that
Spraedsheet is currently slower.

Its batch API defers evaluation and returns consolidated affected-cell changes.
The guide also declares which reads are unavailable while evaluation is suspended.
This does **not** establish an all-or-nothing transaction guarantee; grouping
recalculation, grouping undo, and rolling back a failed operation are separate
contracts.

## Where we stand comparatively

Spraedsheet's `language.ts`, `sheet.ts`, and `workbook.ts` already compile formulas,
await language evaluation, track reads, handle cycles, and settle cross-sheet
changes. `commands.ts` returns changes and supports grouped script work; the bus
can restore workbook state on failure. Adding another engine would duplicate
these responsibilities and complicate RexxJS semantics.

The current 26×100 bound limits what performance claims mean. Raising it requires
a workload suite before sophisticated optimisation.

## Things to build independently

- **A calculation cost harness.** Measure compilation, invalidation, evaluated
  formula count, range visits, settlement rounds, memory, and notification count.
  Include chains, diamonds, running totals, cross-sheet graphs, volatile
  functions, dynamic references, and arrays. Pin time/random inputs where needed.
- **Explicit batch read semantics.** Specify whether scripts read fresh staged
  results or a committed snapshot, when recalculation occurs, and what a failure
  restores. Preserve Spraedsheet's existing read-after-write behaviour deliberately.
  Do not copy HyperFormula's suspended-read restriction by accident.
- **Range-aware optimisation after measurement.** Reuse range dependencies or
  aggregates only where they materially reduce cost. Floating-point regrouping
  can change results; numerical equivalence needs its own acceptance tolerance.
- **One function metadata contract.** Arity, coercion, array lifting, volatility,
  help, and locale names should agree. Extend `described.ts` and current worksheet
  metadata rather than maintaining disconnected documentation lists.
- **A declared compatibility ledger.** Distinguish intended Excel/RexxJS
  differences from bugs, unsupported features, and unverified edge cases.

## First proof and priority

Run an original workload directly through commands and through the control bus;
compare values, errors, changed addresses, and undo state. Track calculation work
as well as elapsed time. Verify an unaffected branch is not evaluated, a dynamic
reference changes its dependencies, and a failing script restores workbook state.

**Priority: high for hardening; optimisation depends on evidence.** This is an
architecture reference, not a recommendation to replace the native engine or to
bring a Node/browser runtime into sae.

## Local evidence and limits

Under `/home/paul/scm/hyperformula/`: `DEV_DOCS.md`,
`docs/guide/dependency-graph.md`, `docs/guide/batch-operations.md`,
`docs/guide/list-of-differences.md`, `src/DependencyGraph/`, `src/Evaluator.ts`,
`src/HyperFormula.ts`, `src/LazilyTransformingAstService.ts`, and
`src/interpreter/functionMetadata/`.

`test/README.md` says the public tests are smoke tests and the comprehensive suite
is separately held. This review did not obtain that private suite. The project
offers GPLv3 and commercial licensing; comparison does not imply code reuse or
an integration licence decision.
