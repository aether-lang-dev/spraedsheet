# LibreOffice Calc: ideas for Spraedsheet

Prepared for Paul and Nic, 18 October 2026.

**Recommendation:** use Calc as a reference for dependable spreadsheet behaviour,
while developing Spraedsheet as a programmable workbook for objects, tables, and
automation. The best next investment is a shared foundation for reference handling,
recalculation, and transactional edits. Build the query and object features in
[ENHANCEMENT_PLAN.md](ENHANCEMENT_PLAN.md) on that foundation.

This is a study of the local LibreOffice checkout at
`/home/paul/scm/libre_office_core`, revision `1fd35eecd`, compared with Spraedsheet
revision `9617649`. It describes what is present in those sources, not a claim
about a particular released LibreOffice version. Spraedsheet is also called
RexxSheet/RexxJS Spreadsheet in the existing documents.

The implementation suggestions below are original design proposals. No LibreOffice
source, tests, or assets are reproduced. This is a source-informed comparison,
not a claim that a formal clean-room development process has taken place.

## What LibreOffice brings

LibreOffice is an integrated office suite; **Calc** is the relevant spreadsheet
application. Its capabilities extend beyond a grid and formula interpreter:

| Area | Calc in this checkout | Opportunity for Spraedsheet |
|---|---|---|
| Workbook model | Documents, sheets, cells, ranges, names, and structural reference updates | Consistent identities and reference semantics across every operation |
| Calculation | Formula compilation, dirty-cell tracking, range listeners, formula groups, threaded/OpenCL calculation paths | Predictable incremental calculation, then measured optimisation |
| Data work | Database ranges, sorting/query settings, imports and transformations, DataPilot pivot tables | Named tables shared by RexxJS queries, pivots, charts, and row editing |
| Editing | Dedicated undo machinery, copying/pasting, validation, conditional formats | A common edit pipeline for UI and Control Bus |
| Interchange | Separate format filters, including ODF and Excel-related formats | Explicit import/export contracts and fidelity reports |
| Automation | UNO interfaces for ranges, bulk data, recalculation, pivots, and goal seeking | A small, explicit Control Bus with comparable operational completeness |
| Presentation | Number formats, charts, printing, accessibility infrastructure | Polished workflows selected for Spraedsheet's users |
| Quality | Extensive calculation, import/export, UI, and regression test collections | Independent behavioural fixtures for the edge cases we support |

Calc's value as a reference is how these features interact. A row insertion can
affect formulas, named ranges, validation, chart sources, and undo history. Adding
the menu item is only part of implementing the behaviour.

## Where Spraedsheet actually stands

[COMPARISON.md](COMPARISON.md) is useful for positioning, but its missing-feature
list is stale. Inspection of the active `src/` implementation finds:

- Multiple sheets, merging, cell editors, pivots, sorting, filtering, named ranges,
  validation, freeze settings, fill operations, and undo/redo in
  [spreadsheet-model.js](src/spreadsheet-model.js).
- CSV, JSON, TOML, and YAML import utilities in
  [spreadsheet-import-export.js](src/spreadsheet-import-export.js).
- Save handling, sheet tabs, and a grid in
  [SpreadsheetApp.jsx](src/SpreadsheetApp.jsx).
- RexxJS integration and a shared set of script/control functions in
  [spreadsheet-rexx-adapter.js](src/spreadsheet-rexx-adapter.js) and
  [spreadsheet-control-functions.js](src/spreadsheet-control-functions.js).

These are implementation observations, not a certification that every feature is
complete, exposed in the UI, or correct across all combinations. This review did
not run either application or benchmark them. Root-level legacy implementations
also exist; this comparison uses `src/`, which the Vite entry point imports.

The important gap is therefore depth and consistency as much as feature breadth.
For example, the model uses regular expressions for dependency extraction and
reference rewriting, starts asynchronous evaluation from a synchronous `setCell`,
and records undo snapshots using shallow copies of several maps. Those choices
deserve attention before larger queries and batches make their limitations more
visible.

## Highest-value ideas to build independently

### 1. Make references and recalculation a workbook service

Calc separates formula compilation, reference updates, dependency notification,
and calculation. Spraedsheet can adopt those responsibilities with a much smaller
JavaScript design while retaining RexxJS as its language.

Build a reference representation that distinguishes a cell, range, named range,
table column, and sheet identity. Integrate it with RexxJS parsing or its variable
resolution hooks so a string containing `"A1"` is not mistaken for a dependency.
Track actual reads where expressions compute references dynamically; use an
explicit conservative invalidation policy when a dependency cannot be known in
advance. A range dependency must notice an edit to any interior cell.

Separate **copying a formula** from **moving its referenced cells**. For example,
copying `=$A$2` should preserve the address; inserting a row before row 2 should
move the reference to `$A$3`. The current insertion logic treats absolute markers
as a reason not to shift, so this distinction is a concrete improvement.

Use a workbook-wide dependency graph keyed by stable sheet identity. Queue dirty
formulas in dependency order, deduplicate work, and report cycles. Give each
evaluation a revision so a slow PyOdide result cannot overwrite a newer edit.
Expose an awaitable completion point to both the UI and Control Bus; merely
awaiting the current synchronous `setCell` does not await its internal work.

**First acceptance examples:** changing an interior range cell updates its sum;
renaming a sheet preserves references; a diamond-shaped dependency calculates
correctly; a stale asynchronous result is discarded; strings survive reference
rewriting unchanged.

### 2. Turn batching into a transaction boundary

The enhancement plan rightly proposes fewer Control Bus round-trips. Calc's bulk
range and calculation-control APIs suggest an additional benefit: defer repeated
calculation and display work until an operation is complete.

Implement one model-level edit operation used by paste, import, scripts, and UI
actions. For a batch, stage changes, validate the proposed result, commit once,
update dependencies, calculate, and publish a consolidated change event. Capture
one undo entry. Make all-or-nothing behaviour the initial batch contract; offer
best-effort processing later only through an explicit option with per-item results.

Replace or strengthen shallow snapshots with immutable before/after data or edit
patches. Include sheet creation/deletion, names, dimensions, table metadata, and
other changed workbook state in the history contract.

**First acceptance examples:** an invalid item leaves an atomic batch unchanged;
undo reverses an entire import; a batch returns only after its documented
calculation boundary. Measure transport time, calculation time, UI notifications,
and history memory separately. A speed-up percentage needs an actual baseline.

Transaction scope should initially cover workbook state. RexxJS functions can
perform effects, including cell writes; classify those functions and define when
they may execute. Rolling back cells cannot automatically reverse an external
side effect.

### 3. Make named tables the bridge between objects and cells

Calc has database-range metadata, query/sort settings, import transformations,
and pivot infrastructure. These support the direction in priorities 1 and 5 of
the enhancement plan, without requiring Spraedsheet to reproduce their APIs.

Define one table descriptor: stable identity, source range, header policy, column
names and types, optional key, and expansion rules. Let queries, filters, pivots,
charts, and object conversion consume it. Avoid implementing a separate notion of
columns for each feature.

Keep ordinary named ranges lightweight; a table adds semantics rather than
silently changing every named range into a database. Specify blank/null values,
duplicate headers, type conversion failures, ordering, and key uniqueness.

The plan's `TABLE(...) |> WHERE(...) |> GROUP_BY(...)` direction is a strong fit.
Prefer parsed RexxJS predicates or callable predicates over an unrelated second
expression language. Decide whether a query is a live view or a materialised
result. For results written to cells, define output ownership, refresh, and
collision behaviour so refresh cannot overwrite unrelated work.

**Object round trips should be explicit:** map object properties to named columns,
preserve types, use keys to reconcile edits, and distinguish missing, blank, and
deleted values. Start with flat records; define nesting and one-to-many relations
later. This would make Spraedsheet particularly useful as an editable view of
application data.

### 4. Enforce validation in the shared write path

Calc provides typed and custom validation, input guidance, and different error
responses. Spraedsheet already has validation rules and editors; extend those
through one write path so scripts, imports, paste, and manual edits behave alike.

Implement the plan's create/update distinction using explicit operation context
and stable record identity. An empty cell does not necessarily mean a new record.
Provide old value, proposed value, operation, and proposed row to a pure validator;
make table uniqueness inspect the staged batch, including other new rows.

If auto IDs are introduced, keep them stable through sorting and insertion.
Specify allocation after failed writes, undo, import, and merge. A local counter
can serve a single workbook; independently edited copies need a separate identity
and reconciliation policy. Row number is unsuitable as record identity.

**First acceptance examples:** two rows with the same key in one batch are
rejected; an invalid imported value reports its row and column; an update can
compare with the previous committed value.

### 5. Offer an explanation of a calculated result

Calc's Detective facilities provide precedent/dependent tracing and error
inspection. Spraedsheet's programmable formulas make an equivalent especially
valuable.

Add an inspector showing formula source, dependencies, evaluation state, error
details, timing, and loaded-library requirements. Extend it to table pipelines:
show row counts after filtering, grouping keys, and the source of output rows.
Expose the same information through the Control Bus for automated diagnosis.

Start with dependency lists and an error chain. Query provenance and a graphical
view can follow once the underlying metadata exists.

## Useful follow-on work

| Idea from Calc | Independently implementable Spraedsheet version | Priority |
|---|---|---|
| Rich import controls | Preview, delimiter/encoding options, per-column types, preserved leading zeros, and an explicit policy for formula-looking text | High after the edit pipeline |
| Dedicated format filters | Versioned native workbook format; round-trip fixtures; report unsupported fields for each interchange format | High |
| Pivot infrastructure | Extend existing pivots with multiple measures, grouping, refresh/invalidation, and named-table sources | Medium |
| Conditional formats and number formatting | Rule editor over RexxJS expressions; documented locale/date/number semantics; separate raw values from display | Medium |
| Goal Seek and Solver interfaces | Start with one numeric input and one target formula, bounded iterations, cancellation, and restored state on failure | Medium, after reliable calculation |
| Accessibility infrastructure | Keyboard-complete grid workflows, focus management, screen-reader labels, and announced validation errors | High for wider adoption |
| Sparse column storage and grouped formulas | Keep the existing sparse maps initially; profile memory and repeated formulas before changing representation | Measure first |
| Extensive regression suites | Original small workbooks covering edits, formulas, undo, refresh, and save/load together | Start now |
| Charts and printing | Build on current chart support with table-bound sources; add print/PDF layouts if users need shareable reports | Demand-led |

For ODS/XLSX, choose a limited supported subset and make losses visible. Preserving
formulas, cached results, styles, names, dates, and charts is a much larger task
than importing cell values. RexxJS formulas also need a declared policy when the
destination application cannot evaluate them. Keep native storage authoritative
until fidelity is demonstrated.

Calc's `LAMBDA` documentation in this checkout also qualifies the existing
positioning: programmable formulas and user-defined functions are not exclusive
to Spraedsheet. Our stronger distinction is **RexxJS throughout the workbook,
object/table pipelines, library integration, and accessible external control**.

## Suggested sequence for Paul and Nic

1. **Establish the baseline.** Refresh the stale feature matrix, distinguish model
   support from UI support, and write down reference, type, and calculation rules.
   Create small independent behavioural workbooks for those rules.
2. **Strengthen the edit engine.** Implement parsed references, workbook-wide
   dependencies, awaited calculation, atomic batches, and complete undo records.
   These directly support priority 2 and make priorities 1 and 3 safer to deliver.
3. **Deliver one object/table workflow.** Combine named-table metadata, keys,
   validation, and a narrow query pipeline. Demonstrate importing customer objects,
   editing rows, filtering/grouping them, and exporting changes by key.
4. **Connect existing analysis features.** Bind pivots and charts to those tables,
   implement refresh and output ownership, and add the inspector.
5. **Expand from measured demand.** Improve interchange, larger-data rendering,
   analysis functions, and reporting according to actual workbooks.

These are dependency-ordered milestones, not time estimates. The day-level
estimates in the existing plan do not account for the correctness work above.
Calc's broad office framework, UNO runtime, GPU infrastructure, and full legacy
format compatibility would each introduce substantial scope; the immediate
product can remain React/Tauri, a DOM-free model, RexxJS, and the Control Bus.

## Local evidence map

Paths below are relative to `/home/paul/scm/libre_office_core/`. They identify the
reference material behind this comparison, not implementation templates.

| Topic | Relevant checkout locations |
|---|---|
| Suite structure and integration | `README.md`, `sc/README.md` |
| Workbook, storage, and calculation | `sc/inc/document.hxx`, `sc/inc/column.hxx`, `sc/source/core/inc/bcaslot.hxx`, `sc/inc/formulagroup.hxx` |
| Structural reference changes | `sc/inc/refupdatecontext.hxx` |
| Tables, transformations, and pivots | `sc/inc/dbdata.hxx`, `sc/inc/dpobject.hxx`, `sc/source/ui/dataprovider/`, `sc/source/ui/inc/datatransformation.hxx` |
| Validation and conditional formatting | `sc/inc/validat.hxx`, `sc/inc/conditio.hxx` |
| Undo and tracing | `sc/source/ui/inc/undobase.hxx`, `sc/inc/detfunc.hxx` |
| Bulk data and calculation APIs | `offapi/com/sun/star/sheet/XCellRangeData.idl`, `offapi/com/sun/star/sheet/XCalculatable.idl` |
| What-if APIs | `offapi/com/sun/star/sheet/XGoalSeek.idl`, `offapi/com/sun/star/sheet/XSolver.idl` |
| Import/export and CSV controls | `sc/source/filter/`, `sc/source/ui/dbgui/scuiasciiopt.cxx` |
| Formats and accessibility | `svl/source/numbers/`, `sc/source/ui/Accessibility/` |
| Programmable formulas | `sc/README.Lambda.md` |
| Behavioural coverage | `sc/qa/unit/ucalc_formula.cxx`, `ucalc_sort.cxx`, `ucalc_copypaste.cxx`, `ucalc_pivottable.cxx`, `ucalc_spilled_range.cxx`, and `subsequent_*` files in the same directory |

When implementing these proposals, write requirements and original tests around
the behaviours we choose, then implement them in Spraedsheet's architecture.
Consult format specifications for interchange semantics. Do not translate Calc
classes or transplant its source, fixtures, or UI assets.
