# Grist compared with Spraedsheet

**Most useful lesson: promote table rows into identifiable records, then let
several views operate on those same records.** This is the strongest comparison
for the object↔spreadsheet direction.

Review: `/home/paul/scm/grist-core`, revision `edab14f`. Baseline and review limits:
[comparison index](../../OPEN_SOURCE_COMPARISONS.md).

## What is present

Grist's README describes typed columns, Python formula columns, reference and
reference-list fields, linked widgets, summaries, forms, incremental imports,
REST access, collaboration, and granular access rules. Its document format uses
SQLite. Those are related capabilities: a record remains meaningful when a
widget sorts or filters it, and another widget can show its related records.

The implementation makes that distinction concrete. `sandbox/grist/records.py`
represents records by row identity and table association. `depend.py` describes
dependencies between table columns, with relations mapping which rows are
affected; property reads participate in dependency tracking. `table.py` exposes
record lookups. `ActionBundle.ts` distinguishes stored changes, calculated
changes, and inverse actions. This is richer than treating a record as whichever
cells currently occupy row 12.

The README also distinguishes community capabilities from features outside
`grist-core`, including its newer built-in MCP and automation products. Do not
attribute the entire hosted product to this checkout. In particular,
Spraedsheet already has its own MCP/control-bus path.

## Where we stand comparatively

| Concern | Current Spraedsheet | Lesson from Grist |
|---|---|---|
| Programmable formulas | Excel or RexxJS per workbook | General-purpose formulas are useful elsewhere too; make RexxJS workflows excellent |
| Tables | Range-based descriptors, types/descriptions, totals and calculated columns | Add record identity and relation semantics to selected tables |
| Objects | JSON row import and RexxJS queries | Preserve an object's identity and types across reconciliation and edits |
| Data presentation | Grid, filters, slicers, pivots and charts | Persist distinct views and link selection between them |
| History | Grouped workbook snapshots | Measure snapshot cost; keep direct edits distinguishable from derived results |

## Things to build independently

1. **A record contract for tables.** Give a table and each field stable identities;
   give each record an immutable identity independent of visible row position.
   An optional business key, such as an account code, is separately editable and
   validated. Existing visible auto IDs can remain useful without being the only
   internal identity mechanism.
2. **Reconcile object imports by key.** Preview additions, changes, conflicts,
   and optional deletions. Explicitly distinguish missing properties, blank
   values, nulls, and deletion. Commit the selected reconciliation as one command.
3. **Reference fields and linked views.** Start with Orders → Customer and a
   customer detail pane showing its orders. Editing a display name must preserve
   the relationship. Define restrict/clear/cascade policy before record deletion.
4. **Record-aware invalidation.** Extend the existing cell dependency machinery
   through relation reads. A lookup that currently matches nothing must update
   when a matching record is inserted; dependencies cannot cover only existing
   matches. Measure before introducing column-level optimisations.

Keep free-form worksheets available. A calculated column should have a clear
override policy; automatically making every worksheet column homogeneous would
change the product substantially.

## First proof and priority

Import Customers and Orders from independent JSON fixtures. Re-import changed
objects by business key without duplicates. Rename a customer, sort Orders,
change an order in a filtered view, undo it, and reopen the file. The same
customer/order links and totals must remain. Exercise grid, REXX and MCP entry
points against the same expected workbook state.

**Priority: high**, following the existing recording/hardening work. Avoid taking
on a collaboration server or full permissions system just to get linked records.

## Local evidence

Under `/home/paul/scm/grist-core/`: `README.md`,
`sandbox/grist/{records,table,depend,engine,summary}.py`,
`app/common/ActionBundle.ts`, `app/common/DocActions.ts`,
`app/client/components/LinkingState.ts`,
`app/client/components/Forms/FormView.ts`,
`app/server/lib/ActiveDocImport.ts`, `app/server/lib/GranularAccess.ts`, and
`sandbox/grist/test_lookup_relations.py`.

Spraedsheet counterparts: `app/sheet/{language,query,tabular,workbook,commands}.ts`.
The proposal is behavioural; it does not require Python, SQLite, or Grist's data
structures to be adopted.
