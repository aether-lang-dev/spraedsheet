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

## Candidates for direct reuse with credit

IronCalc's README offers **MIT OR Apache-2.0**, at the recipient's option.
For selected reused material, choose MIT and retain its notices. This extends
the original independent-implementation proposal: attributed copies of selected
fixtures and adapted tests are worthwhile, followed by small production-code
ports where evidence demonstrates a benefit. Nothing is copied by this document.

### Copy fixtures and adapt their assertions first

| Candidate | Material to reuse | Value for Spraedsheet |
|---|---|---|
| Volatile shared-formula regression | `xlsx/tests/shared_formula_volatile.xlsx` and assertions from `xlsx/tests/test_import_shared_formula_volatile.rs` | Check that an empty shared-formula daughter element carrying `ca="1"` retains its translated formula, rather than being mistaken for a spill placeholder |
| Optional XLSX style attribute | `xlsx/tests/optional_xf_id.xlsx` and assertions from `xlsx/tests/optional_xf_id_import.rs` | Accept styles whose `xfId` attribute is absent while preserving their supported alignment/wrapping properties |
| Number-format cases | Selected inputs and expected outputs from `base/src/test/test_number_format.rs` and `base/src/formatter/test/` | Exercise rounding, sections, dates, fractions, scientific notation, and text preservation against our formatter |
| Editing regressions | Selected scenarios from `base/src/test/user_model/`, including autofill, cut/paste, merged cells, arrays, and undo | Check interactions between features, not only isolated edit commands |
| Formula compatibility workbooks | A curated subset of `xlsx/tests/calc_tests/` | Broaden workbook-based formula coverage without authoring every fixture ourselves |

Copy approved workbook fixtures unchanged where practical, and translate the Rust
assertions into our headless harness. Record which assertions were adapted and
which behaviours we intentionally handle differently. These cases supplement
our original corpus; neither IronCalc's expected values nor cached workbook
results should automatically become the oracle for every disputed behaviour.

### Consider small production-code ports after the tests

- **Series detection:** `base/src/user_model/sequence_detector.rs` is a candidate
  for an attributed port into the existing `app/sheet/series.ts` functionality.
  Useful areas include locale-aware number sequences, month/day names, case
  preservation, and decimal progression handling. Demonstrate gaps with the
  imported/adapted tests first; retain existing behaviour such as padded suffixes
  where it is intentional.
- **Specific formatter routines:** inspect `base/src/formatter/` for a small
  routine that solves a demonstrated gap in `app/sheet/numfmt.ts`. Port the
  smallest useful unit together with its relevant cases, accounting for runtime
  and numeric differences between Rust and sae.

The Rust workbook engine and history machinery would require substantial
adaptation. Keep our current implementations unless measurements identify a
concrete benefit that justifies replacing them.

### Attribution and provenance

For each copied fixture, adapted test, or port:

1. Check its individual provenance and any third-party notices, particularly for
   workbooks and generated data. Do not assume every bundled asset has identical
   provenance merely because the repository offers MIT.
2. Preserve IronCalc's applicable copyright and full MIT permission/warranty
   notice, including `LICENSE-MIT`, with the reused material.
3. Record the upstream repository, exact commit, original paths, selected licence,
   and local destinations in a provenance file. The reviewed short revision is
   `297553d`; resolve and store the full commit when material is imported.
4. Mark adaptations and ports, describe local modifications, and keep upstream
   fixtures separate from our original cases so later updates are reviewable.

**First reuse batch:** the volatile shared-formula and optional-`xfId` XLSX
regressions, plus selected number-format cases. Run those against the current app
before deciding which production routines warrant a port.

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
