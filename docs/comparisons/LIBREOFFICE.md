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
   limits it). Row lineage and grouping
   explanations done 2026-10-25 on VisiData 3's typed keys: `RANGE(...) |> GROUP_BY(..)
   |> LINEAGE()` spills each group's key, row count and sheet rows; `|> EXPLAIN("sum",
   "amount")` says why each aggregate is what it is ("North: SUM of amount over rows 4,
   6: 120 + 40 = 160", with any non-numbers skipped said); `group.explain range= by=
   [agg= value= key=]` (`GROUPEXPLAIN` from a script) does the same over the active
   sheet and gives the cells of each group's rows. Headless `sheet_lineage`. Not
   done: stepwise REXX is RexxJS's.
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
   panel, print/layout integration, larger-grid interaction. Done in part
   2026-10-24: a screen-reader pass over the toolbar and all fifteen panels (spec_sheet
   "every control ... has an accessible name": no visible button, field, slider or
   toggle lacks an accessible name, and the audit sees at least five controls in each
   panel); it found nothing to fix. Done 2026-10-25, with no new aether-ui capability
   needed (so nothing to add on the four backends):
   - **Keyboard-only:** spec_sheet "keyboard only: every control in the toolbar and in
     each of the fifteen panels takes the keyboard focus": each shown control is given the
     keyboard focus by the driver and `/focus` must name it. F6 through the panels, the
     focus in each, was already tested. Not done: the *order* Tab visits them in.
     aether-ui's driver has no "move focus as Tab does" route (`/focus` reads, a widget's
     `/focus` sets); a real-Tab walk needs one on each backend (GTK `gtk_widget_child_focus`,
     AppKit `selectNextKeyView`, Win32 `GetNextDlgTabItem`, UIKit's focus system, Android
     `focusSearch`) and is the next step if Paul wants the order itself asserted.
     `real_keys_x11` already sends real Tab and Shift+Tab to the grid on the Linux lane.
   - **Print/layout:** Page Setup is the sheet's own (`page.setup` / `page.layout`,
     PAGESETUP / GETPAGESETUP, and a row in the File panel): paper, orientation, margin,
     scale (10 to 400 percent), print area, rows to repeat at the top of each page after
     theirs, header and footer (&P &N &A &F, &L &C &R sections), gridlines and headings.
     It is undoable (a journal entry), saved in the file and in XLSX (`pageSetup`,
     `pageMargins`, `printOptions`, `headerFooter`, `_xlnm.Print_Area`, `_xlnm.Print_Titles`),
     and Print, Export PDF and Page Breaks lay the sheet out by it unless told otherwise.
     Headless `sheet_pagesetup`. Left: columns to repeat; fit to N pages wide; the File
     panel's Export PDF and Print buttons still pass `gridlines` themselves (a decision
     for Paul, TODO.md).
   - **Larger grid:** `spec_sheet_large` (a page of its own: spec_sheet's is full of sixty sheets by
     its end): over 20,000 rows Ctrl+Down, Ctrl+Up,
     Ctrl+Shift+Down (counted in the status bar), Ctrl+End and Ctrl+Home land where
     Excel's do.
7. **Stable record identity, relations, saved views** (the object/table
   contract): see Grist.md and Airtable.md.
