# IronCalc material in this repository

Upstream: <https://github.com/ironcalc/IronCalc>
Commit: `297553dc3a37ec6b0f4821055220e3d26c4ac6a1`
Licence: IronCalc is MIT OR Apache-2.0 at the recipient's option; we take it
under **MIT**. Its `LICENSE-MIT` (copyright EqualTo GmbH and Nicolas Hatcher,
with the full permission and warranty notice) is [LICENSE-MIT](LICENSE-MIT)
here, unmodified, and applies to everything listed below.

Why it is here: Paul approved copying and porting selected IronCalc material
into Spraedsheet under that licence (an exception to the usual not-ours rule,
for IronCalc only). The plan is `docs/comparisons/IRONCALC.md`.

Rules this file keeps:

- Every upstream path copied, its local destination, and what was adapted is
  recorded here, in the commit that brings it in.
- Upstream workbooks are copied byte for byte into `third_party/ironcalc/`
  (never mixed with our own cases); everything we translated lives in our own
  test files, each case naming the upstream test it came from.
- Upstream expected values are not an oracle. Where ours differ, Excel's
  documented behaviour decides, and an intentional difference is recorded
  under "Deliberate differences" below.
- No workbook is assumed to have the provenance of the repository as a whole;
  each fixture's own notes are under its entry.

## Copied and adapted material

| Upstream path | Local destination | What was adapted |
|---|---|---|
| `LICENSE-MIT` | `third_party/ironcalc/LICENSE-MIT` | nothing |
| `xlsx/tests/shared_formula_volatile.xlsx` | `third_party/ironcalc/fixtures/xlsx/shared_formula_volatile.xlsx` | nothing (byte for byte) |
| `xlsx/tests/optional_xf_id.xlsx` | `third_party/ironcalc/fixtures/xlsx/optional_xf_id.xlsx` | nothing (byte for byte) |
| `xlsx/tests/test_import_shared_formula_volatile.rs` | `tests/headless/sheet_ironcalc_xlsx.ts` | assertions translated to our headless harness: the master's and daughter's formulas and the daughter's value, plus that nothing is reported left out |
| `xlsx/tests/optional_xf_id_import.rs` | `tests/headless/sheet_ironcalc_xlsx.ts` | assertions translated: the file loads and A2's wrap survives; the vertical alignment is not asserted (see below); we added the merge and frozen-pane checks |

The two workbooks are read by the headless tests as base64 in
`tests/headless/lib/ironcalc_fixtures.ts`, which `tests/headless/lib/mkironcalc.py`
generates from the copies above (the harness loads modules, not files).

Fixture notes: both were authored in IronCalc's repository (git history shows
contributors' commits, no third-party notice in the repository). Each is an
ordinary workbook; `optional_xf_id.xlsx` carries a stock Office `theme1.xml` and
`shared_formula_volatile.xlsx` a smaller one, which are Office's default
theme data rather than IronCalc's creative content. No cell content in either
is anything but placeholder numbers and a title string.

## Deliberate differences

- **Vertical alignment (`optional_xf_id`).** IronCalc asserts A2 reads
  `vertical="top"` as well as wrapped. Spraedsheet has no vertical-alignment
  look (cells are drawn top-aligned by default and the model keeps no
  alignment on that axis), so the importer drops `vertical=`; the test
  asserts the wrap only. This is a known unmodelled feature, not a bug in the
  `xfId` handling the fixture exists for (which already passed).
