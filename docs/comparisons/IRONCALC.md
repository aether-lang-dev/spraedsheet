# IronCalc compared with Spraedsheet

**Most useful lesson: represent reversible edits carefully and test compatibility
with complete workbooks.** Rust/WASM is interesting context, but the transferable
ideas fit our existing Aether/AetherUI/sae implementation.

Review: `/home/paul/scm/IronCalc`, revision `297553d`. Baseline and limits:
[comparison index](../../OPEN_SOURCE_COMPARISONS.md).

## What is present

IronCalc separates its Rust engine (`base`), XLSX reader/writer (`xlsx`), bindings,
and web applications. The engine includes formula parsing/evaluation, locale and
timezone support, formatting, and a user-facing model for edits.

`base/src/user_model/history.rs` represents edits as typed differences with the
data needed to reverse them. Undo/redo applies lists of those differences. The
test tree includes cut/paste, sheet operations, formatting, arrays, batch
row/column edits, and diff queues. Those files establish specific engineering
concerns; a diff queue alone does not establish a complete collaborative editing
protocol.

`xlsx/build.rs` generates tests from workbook folders, with different policies for
calculation-only fixtures, export checks, templates, and documentation examples.
It also makes exclusions visible for volatile or unstable examples. Compatibility
is exercised as a workbook workflow rather than just a list of function names.

## Where we stand comparatively

Spraedsheet already separates calculation, commands, view modules, and file
handling. Its worksheet semantics and REXX control are native to its model; moving
that model to Rust would introduce a second integration task without itself
answering a user need.

The workbook uses snapshots for history, with grouped operations and a retained
history limit. That is a reasonable correctness baseline for the current small
grid. The question is how much memory and copying it costs at larger sizes, and
whether every new piece of metadata is included.

XLSX support already reads/writes substantial workbook content and reports losses.
IronCalc is useful for comparison of coverage and test organisation, not grounds
for assuming either engine has perfect Excel fidelity.

## Things to build independently

- **Measure history cost.** Compare one-cell edits, a large paste, style changes,
  structural edits, and scripted groups. Include redo memory. Only then choose
  immutable sharing, reversible patches, periodic checkpoints, or continued
  snapshots.
- **If patches are warranted, describe semantic changes.** Include names,
  references, styles, merges, tables, chart sources, and record identities.
  Reconstruct derived values rather than accidentally persisting stale spill
  output. A patch format also needs versioning before it becomes a file format.
- **An original workbook corpus.** For each supported interchange feature, test
  open → calculate → edit → save → reopen. Compare semantic state, not ZIP byte
  equality. Test unsupported-feature reports and source metadata preservation.
- **Locale/time controls.** Separate formula storage syntax, number/date display,
  import parsing, and timezone-dependent functions. Inject the clock for tests;
  avoid making saved results depend silently on the host locale.

## First proof and priority

Use independently authored workbooks covering mixed absolute references,
1900/1904 dates, shared formulas, strings that resemble numbers, validations,
tables, and a deliberately unsupported feature. After edit/save/open, compare
formulas, raw values, styles, names, and warnings. Undo/redo the same workflow and
assert semantic equality.

**Priority: high for the corpus; measure before changing history storage.** A
native app and headless tests can share these contracts without adopting a
browser frontend or an alternate engine language.

## Local evidence

Under `/home/paul/scm/IronCalc/`: `README.md`, `base/README.md`,
`base/src/{model,evaluation,workbook}.rs`, `base/src/user_model/`,
`base/src/test/user_model/`, `xlsx/README.md`, `xlsx/build.rs`,
`bindings/wasm/README.md` and `bindings/python/README.md`.

Spraedsheet counterparts: `app/sheet/{workbook,file,xlsx,xlsxwrite,dates,language}.ts`
and `tests/headless/`. This comparison does not reproduce upstream test fixtures
or implement its diff structures.
