# LibreOffice Calc: ideas for the Aether/AetherUI/sae Spraedsheet

Updated for Paul and Nic, 18 October 2026.

**Recommendation:** keep the native Aether/AetherUI/sae implementation and learn
from Calc's depth of behaviour, correctness work, and integration between
features. Spraedsheet already has much of the previously proposed functionality;
the next work is stronger guarantees, larger-workbook evidence, and record-aware
object/table workflows.

This replaces the comparison of the old React/Tauri application. That version is
abandoned. Browser DOM work, a Tauri backend, and Pyodide integration are not
recommendations for the current app.

Review: `/home/paul/scm/libre_office_core` at `1fd35eecd`, compared with Spraedsheet
`cfc906f`. These are local source/documentation observations, not runtime tests
or a claim about a particular LibreOffice release. See the
[comparison index](OPEN_SOURCE_COMPARISONS.md) for the other studies and common
review limits. No source, fixtures, assets, or numerical routines are copied;
this is source-informed analysis, not a formal clean-room process.

## What LibreOffice brings

LibreOffice is an integrated office suite; **Calc** is its spreadsheet application.
The useful comparison is how its features interact inside a mature workbook.

| Area | Calc in this checkout | Current Spraedsheet | Useful next work |
|---|---|---|---|
| Workbook model | Sheets, ranges, names, reference-update contexts | Sheets, scoped names, structural edits, both formula languages | Cross-feature reference and persistence regression coverage |
| Calculation | Compiled formulas, dirty-cell tracking, range listeners, formula groups, threaded/OpenCL paths | Compiled language adapters, tracked reads, awaited calculation, cross-sheet settling/cycles, manual/iterative modes | Measure invalidation, range work, and settlement at scale |
| Data | Database ranges, transformations, DataPilot pivots | Tables, structured references, RexxJS queries, pivots, slicers, auto IDs | Stable record identity, typed query contracts, relations, keyed object reconciliation |
| Editing | Dedicated undo infrastructure and rich edit operations | Command-based edits, grouped workbook snapshots, script rollback | Malformed-command and random-sequence testing; measure history cost |
| Inspection | Detective tracing and calculation diagnostics | Precedents/dependents, watches, error checks, Excel evaluation steps | RexxJS expression steps, query lineage, timing explanations |
| Interchange | Dedicated ODF, Excel, and other filters | Native JSON, CSV/TSV, XLSX read/write, XLS import, ODS cached-value import, PDF output | Semantic round trips and accurate unsupported-feature reports |
| Presentation | Formatting, charts, printing, accessibility infrastructure | Native grid/menus/panels, formatting, charts and drawings | Native layout, keyboard accessibility, and remaining print/platform work |
| Automation | UNO interfaces for cells, ranges, calculation, pivots and analysis | Shared command registry, REXX, control bus, MCP | REXX recording, headless/interactive parity, operational limits |

Presence is not proof of equal completeness or correctness. In particular, neither
file support nor function counts establishes full Calc/Excel fidelity.

## Where Spraedsheet now stands

[README.md](README.md) identifies `app/` as the implementation. Its page and view
modules use sae and AetherUI; the workbook and commands have no DOM dependency.
[The tracker](docs/gridcraft_parity.md) and [TODO.md](TODO.md) describe the current
feature coverage and remaining work. The earlier `ENHANCEMENT_PLAN.md` and
`COMPARISON.md` are historical documents, available through git history.

Reading the active model changes the assessment:

- `language.ts`, `parser.ts`, `shift.ts`, and `sheet.ts` already separate formula
  compilation, reference changes, reads, and evaluation. The old criticism of
  regex-only model dependency extraction does not describe this implementation.
- `workbook.ts` provides cross-sheet settlement, names, metadata, grouped history,
  and persisted workbook state. `commands.ts` routes edits and exposes script
  grouping/rollback. These foundations already exist.
- `query.ts` and `rexx.ts` provide TABLE/RANGE, filtering, grouping, selection and
  aggregation; `language.ts` describes range-based tables and column metadata.
  Adding the original plan's pipeline functions again would duplicate work.
- `address.ts` still bounds each sheet to 26×100. Larger-file requirements need
  measured model and native-grid work together.

The rest of this document proposes extensions, rather than marking existing
features as new work.

## Ideas to develop independently

### 1. Test feature interactions as workbook behaviour

Calc has explicit reference-update context and broad calculation/import/export
regression collections. We should maintain our own behavioural fixtures for
structural edits through formulas, names, tables, validation, conditional formats,
merges, chart sources, and history together.

For example, insert a row into a table read by another sheet, rename that sheet,
undo/redo, save and reopen. Compare formulas, raw values, names and metadata.
Test copying separately from moving: `$` anchors affect copying, while a
reference follows its target through structural insertion. The current app
already models this distinction; regression coverage should preserve it.

Run matching workloads through commands, REXX/control bus, and interactive actions.
Use original fixtures and documented semantics, not Calc's source or test corpus.

### 2. Measure calculation before changing its architecture

Calc's range listeners, sparse column storage and grouped formulas show why
large-sheet work requires more than raising a row constant. For our engine,
measure dependency count, dirty formulas, range visits, cross-sheet rounds,
notification count, history memory, and time by phase.

Include long chains, diamond graphs, running totals, dynamic references, spills,
and controlled volatile functions in both supported languages where applicable.
Check unchanged branches and cycles as well as final numeric answers. Then decide
whether range dependencies, cached aggregates, shared compiled formulas, or
different settlement scheduling address the observed costs. Native parallelism
and GPU execution would require separate evidence and runtime design.

### 3. Finish the object/table contract

Calc's database ranges, query/sort settings and pivot infrastructure illustrate
the value of shared data metadata. Spraedsheet can go further toward its own
object↔sheet goal without turning all worksheets into databases.

Extend selected tables with stable table/field/record identities, optional business
keys, declared types, null/blank rules, and keyed import reconciliation. Keep
ordinary named ranges lightweight. Make filters, queries, pivots, charts and record
forms consume the same table contract.

Current queries materialise rows and group by textual keys. Define type-sensitive
equality and source lineage before adding editable joins or live views. Specify
output ownership and refresh collisions for derived results. The important new
capability is a safe object round trip, not another spelling of GROUP_BY.

### 4. State validation and transaction policies precisely

Calc supports typed/custom validation, prompts and different error responses.
Spraedsheet already supports contextual validation, several rule types, auto IDs,
and grouped scripted edits. Extend and test that behaviour rather than re-add it.

Declare validation policy for manual edits, paste, imports, scripts and record
forms. There may deliberately be different policies; make bypass/report modes
explicit. For unique keys, validate the whole proposed batch, including conflicts
between new rows. A blank field is not sufficient to establish record creation.

Keep recalculation batching, one-step undo, and atomic workbook rollback distinct.
Document read-after-write semantics in scripts. Workbook rollback cannot reverse
external effects; future triggered scripts need effect and retry contracts.

### 5. Extend inspection where RexxJS differs

Calc's Detective features are useful precedent for tracing and error diagnosis,
which our app already offers. Build on those with calculation timings, library
requirements, query row lineage and grouping explanations. Stepwise RexxJS
evaluation needs RexxJS's own expression representation; changes to the language
belong in canonical RexxJS first, then vendoring here.

Calc's `README.Lambda.md` also makes clear that user-defined programmable formulas
are not unique to Spraedsheet. The distinction to develop is REXX across formulas,
recordable operations, object/table workflows and external control.

### 6. Make compatibility and native usability assessable

For XLSX/ODS and native storage, publish supported semantics and known losses.
Test original workbooks through open/edit/save/reopen, including styles, names,
dates, errors, and unsupported objects. ODS cached-value import is not ODS formula
execution. An exported RexxJS workbook needs an explicit formula/results policy
for applications unable to execute RexxJS.

For UI work, improve AetherUI's datagrid and sae page services where needed, test
there, and consume them here. Complete keyboard-only and screen-reader workflows,
layout/print integration, and larger-grid interaction. Browser/DOM components
from another spreadsheet would not advance the selected native architecture.

## Suggested order

1. Harden current behaviour with cross-feature workbooks, malformed commands,
   seeded edit sequences, and headless/interactive comparisons.
2. Deliver REXX session recording from the existing journal (TODO 7.1).
3. Define stable record identities and keyed object reconciliation, then saved
   views and a detail form; use existing tables and queries.
4. Connect relations and derived-view lineage; improve refresh/error diagnostics.
5. Expand size, interchange and native presentation from measured requirements.

This is a dependency-aware proposal, not a time estimate or a replacement for the
existing tracker. Calc's office-wide framework, UNO runtime and acceleration paths
do not need to be reproduced to deliver these outcomes.

## Local evidence map

Paths relative to `/home/paul/scm/libre_office_core/`:

| Topic | Locations consulted |
|---|---|
| Suite and Calc organisation | `README.md`, `sc/README.md` |
| Workbook, storage, calculation | `sc/inc/document.hxx`, `sc/inc/column.hxx`, `sc/source/core/inc/bcaslot.hxx`, `sc/inc/formulagroup.hxx` |
| Reference updates | `sc/inc/refupdatecontext.hxx` |
| Data and transformations | `sc/inc/dbdata.hxx`, `sc/inc/dpobject.hxx`, `sc/source/ui/dataprovider/`, `sc/source/ui/inc/datatransformation.hxx` |
| Rules, history, tracing | `sc/inc/validat.hxx`, `sc/inc/conditio.hxx`, `sc/source/ui/inc/undobase.hxx`, `sc/inc/detfunc.hxx` |
| Automation contracts | `offapi/com/sun/star/sheet/XCellRangeData.idl`, `XCalculatable.idl`, `XGoalSeek.idl`, `XSolver.idl` in the same directory |
| Interchange, display, access | `sc/source/filter/`, `sc/source/ui/dbgui/scuiasciiopt.cxx`, `svl/source/numbers/`, `sc/source/ui/Accessibility/` |
| Programmability and coverage | `sc/README.Lambda.md`, `sc/qa/unit/ucalc_*.cxx`, `sc/qa/unit/subsequent_*` |

Spraedsheet evidence: `app/sheet/{language,sheet,workbook,commands,query,rexx,shift,audit,file,xlsx,xlsxwrite}.ts`,
`app/sheet/view/`, `tests/headless/`, `tests/specs/`, and the current tracker.
