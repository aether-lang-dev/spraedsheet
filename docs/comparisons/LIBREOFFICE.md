# LibreOffice Calc: actions

**Licence and reuse:** MPL-2.0. Ideas only: no source, fixtures, assets or
numerical routines copied. Native Aether/AetherUI/sae stays; no DOM, Tauri or
Pyodide work.

Done, so dropped: cross-sheet settlement, names, tables, pivots, grouped undo and
script rollback (TODO 1-4, 7.1); the full-size sheet (1.4); large-file limits
(7.8); malformed-command and seeded-sequence fuzz (7.6); headless vs interactive
(7.7); Excel-style tracing, watches, evaluate steps and error checking (2.16);
XLSX/XLS/ODS paths with loss reports (5.x); accessibility and localization
(7.9, 7.10).

1. ~~**Cross-feature structural-edit workbook.** Insert a row into a table read by
   another sheet, rename that sheet, undo and redo, save, reopen; compare
   formulas, values, names, metadata. Include validation, conditional formats,
   merges, chart sources. Copy (`$` anchors) tested separately from move. Same
   workload by commands, REXX/bus and UI. Shared with Univer 3.~~ Done 2026-10-21
   for commands, the file, the text form and the REXX program (headless
   `sheet_structural`): a table, workbook and sheet-own names, validation, a
   conditional format, a merge, a chart source, a comment and a second sheet's
   formulas, a row inserted above them, the other sheet renamed, undo and redo,
   reopened three ways, and a column inserted after. It found one bug, fixed: a
   name that is a range (the workbook's or a sheet's own) did not move when rows
   or columns were inserted or deleted on its sheet, so `=TOTAL` read the wrong
   cell (a name whose cells are deleted is removed; Excel keeps it as `#REF!`).
   Left: copy against move with `$` anchors, and the same workload by REXX/bus
   and UI. The gap it showed is closed (2026-10-21): a structured reference to a
   table on another sheet (`=SUM(Sales[Amount])` away from the table's sheet) was
   `#REF!`; tables are the workbook's, as in Excel, and the reference follows the
   table through inserted rows, a renamed column, a deleted table, undo and a
   reopened file (headless `sheet_crosstable`).
2. ~~**Calculation measurement** (dirty formulas, range visits, settlement rounds,
   history memory; chains, diamonds, running totals, dynamic refs, spills,
   volatiles). See HyperFormula 1. No parallelism or GPU work without evidence.~~
   Done 2026-10-24 with HyperFormula 1 (counters, headless `sheet_work`, which also
   has a spill) and IronCalc 4 (history memory, `sheet_historycost`). No parallelism
   or GPU work: the one cost found that the counters show (a cumulative `SUM` down a
   column) is a rounding decision, not a compute-bound one.
3. ~~**Validation policy per entry path:** manual edit, paste, import, script,
   record form; make any bypass explicit. Unique keys validate the whole proposed
   batch, including conflicts among new rows. Fold into Grist 2.~~ Done 2026-10-24
   for the paths that exist (headless `sheet_validationpolicy`): typing and a
   script's SETCELL are refused when they break a rule; a paste, a fill and an
   import bring cells in (as Excel's do) and now say after which cells break their
   rule and that `validation.circle` shows them; `validation.policy` lists the
   paths and which are checked. Left with the record model (a decision for Paul):
   the record form, and unique keys validated across a whole batch.
4. **Inspection beyond Excel's:** calculation timings, query row lineage,
   grouping explanations; stepwise RexxJS evaluation belongs in canonical RexxJS
   first, then vendored. Lineage: see VisiData 3. Done in part 2026-10-24:
   calculation timings, `formulas.profile` (`PROFILE()`): on=1 times each formula's
   evaluations, none lists the slowest cells with their milliseconds and runs
   (headless `sheet_profile`; the clock is `Date.now`, so a page's coarse clock
   limits it). Not done: row lineage and grouping explanations wait on VisiData 3
   (typed group keys first); stepwise REXX is RexxJS's.
5. ~~**Interchange honesty:** publish supported semantics and known losses; an
   exported RexxJS workbook needs a formula/results policy for apps that cannot
   run RexxJS. ODS stays cached-value import.~~ Done 2026-10-24: `file.interchange`
   lists each format (xlsx, xls, ods, csv, json, html, pdf) as read or written
   with what it keeps and what it loses, and states the RexxJS policy: written to
   XLSX as its values, not its formulas, said in the file report. Headless
   `sheet_interchange` checks the stated losses against the reports the writer
   gives. ODS stays cached-value import (read only).
6. **Native usability gaps:** improve datagrid and sae services in aether-ui/sae
   (tested there), then consume: keyboard-only and screen-reader passes over every
   panel, print/layout integration, larger-grid interaction.
7. **Stable record identity, relations, saved views** (the object/table
   contract): see Grist.md and Airtable.md.
