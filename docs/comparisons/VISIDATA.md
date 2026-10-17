# VisiData compared with Spraedsheet

**Most useful lesson: an exploratory session can become a repeatable program.**
This is a particularly good reference for TODO 7.1, recording UI actions as REXX.

Review: `/home/paul/scm/visidata`, revision `f9f3dbe`. Baseline and limits:
[comparison index](../../OPEN_SOURCE_COMPARISONS.md).

## What is present

VisiData is a Python terminal tool for exploring tabular data. Its sheets and
columns can expose underlying objects through getters/setters; loaders adapt
different sources. Joining, grouping, pivoting, expressions, and derived sheets
are normal parts of the workflow.

`cmdlog.py` records named commands with their inputs and target context. It
supports replay and distinguishes replayable commands from navigation and other
operations that should not normally be logged. Row targeting can use keys rather
than only row positions. Functional tests replay command scripts and compare
outputs. These make automation a continuation of interactive use.

`column.py` distinguishes raw values, converted values, display text, and errors;
columns can cache or calculate asynchronously. `modify.py` supports deferred
additions, changes, and deletions for applicable sheet types. Join/pivot code
retains source associations, useful context for understanding derived results.
That does not mean every arbitrary join is safely editable.

## Where we stand comparatively

Spraedsheet already has a command registry/journal, RexxJS queries, imports,
grouping, pivots, and external program execution. It has raw/formatted values and
errors, and native keyboard interaction through AetherUI. The missing link is
turning a session into a durable, understandable REXX script and preserving enough
identity to replay it against changed data.

VisiData is a data-workflow reference, not an Excel layout/formatting benchmark.
We can keep a familiar native grid and apply its automation lessons.

## Things to build independently

- **Record resolved commands as editable REXX.** Capture table/record/field IDs
  where available and explicit ranges elsewhere. Serialize strings, multiline
  text, and Unicode losslessly. Provide Record, Stop, inspect, save, and replay.
  Define what file operations, undo, view changes, and external effects record.
- **Parameterise useful inputs.** Let an imported filename or filter criterion
  become a script argument. Preview the replay's targets and failures. State
  whether a whole replay is one undo group and how failures restore state.
- **Make source lineage available.** A grouped result should allow inspection of
  contributing records. Permit editing through a derived view only when the
  target record/field is unambiguous. Aggregates should be read-only by default.
- **Preview outbound object changes.** Use deferred patches for inserts, updates,
  and deletions, with an explicit apply/discard boundary. This extends object
  round trips beyond exporting whichever cells happen to be visible.
- **Use replay as a regression tool.** One original workflow should run via UI,
  recorded REXX, and the bus, producing the same semantic workbook state.

The existing `query.ts` uses materialised rows and textual grouping keys. Before
adding lineage, define type-sensitive equality, null handling, duplicate keys,
and whether query results remain live or are snapshots.

## First proof and priority

Record: import customer objects, filter a region, edit one customer's status,
group totals, and export changes. Replay from a clean workbook, then against
the same records in a different order. The intended customer must change in both
cases. A missing/duplicate key should produce a clear failure rather than editing
the record occupying the former row number.

**Priority: very high for recording.** This builds directly on our architecture
and differentiates the product without introducing another formula language.
Object writeback and joins follow identity and reconciliation contracts.

## Local evidence

Under `/home/paul/scm/visidata/`: `README.md`, `CLAUDE.md`,
`visidata/{cmdlog,column,modify,pivot,aggregators}.py`,
`visidata/features/join.py`, `visidata/features/replay_bulk.py`, `tests/`,
and `tests/golden/`.

Spraedsheet counterparts: `app/sheet/{commands,script,control,query,tabular,workbook}.ts`,
`rexx/`, and TODO 7.1–7.2. No upstream replay scripts or Python implementations
are copied.
