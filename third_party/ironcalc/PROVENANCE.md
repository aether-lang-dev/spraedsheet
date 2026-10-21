# IronCalc material in this repository

Upstream: <https://github.com/ironcalc/IronCalc>
Commit: `297553dc3a37ec6b0f4821055220e3d26c4ac6a1`
Licence: IronCalc is MIT OR Apache-2.0 at the recipient's option; we take it
under **MIT**. Its `LICENSE-MIT` (copyright EqualTo GmbH and Nicolas Hatcher,
with the full permission and warranty notice) is [LICENSE-MIT](LICENSE-MIT)
here, unmodified, and applies to everything listed below.

Why it is here: Paul approved copying and porting selected IronCalc material
into Spraedsheet under that licence (an exception to the usual not-ours rule,
for IronCalc only). The remaining candidates are `docs/comparisons/IRONCALC.md`.

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
| `xlsx/tests/optional_xf_id_import.rs` | `tests/headless/sheet_ironcalc_xlsx.ts` | assertions translated: the file loads and A2's wrap survives; A2's vertical top is asserted too (TODO 4.2 added vertical alignment); we added the merge and frozen-pane checks |

The two workbooks are read by the headless tests as base64 in
`tests/headless/lib/ironcalc_fixtures.ts`, which `tests/headless/lib/mkironcalc.py`
generates from the copies above (the harness loads modules, not files).

Fixture notes: both were authored in IronCalc's repository (git history shows
contributors' commits, no third-party notice in the repository). Each is an
ordinary workbook; `optional_xf_id.xlsx` carries a stock Office `theme1.xml` and
`shared_formula_volatile.xlsx` a smaller one, which are Office's default
theme data rather than IronCalc's creative content. No cell content in either
is anything but placeholder numbers and a title string.

Number formats (no code copied; IronCalc's inputs and expected outputs only):

| Upstream path | Local destination | What was adapted |
|---|---|---|
| `base/src/test/test_number_format.rs` | `tests/headless/sheet_ironcalc_numfmt.ts` | `test_simple_format`, `test_maximum_zeros`, `test_leading_comma_text` as data rows (`test_wrong_locale` is IronCalc's locale API, not ours) |
| `base/src/formatter/test/test_en_examples.rs` | same | every case, including the upstream `#[ignore]`d fraction, rounding and leading-hash ones |
| `base/src/formatter/test/test_general.rs` | same | the numeric, percent, rounding, sections, currency, scientific and date cases (colour, condition-syntax error and locale cases left out: they test IronCalc's error type and colour palette indices) |
| `base/src/formatter/test/test_scientific.rs`, `test_time.rs`, `test_dates.rs` | same | every case (the dates one is `m.d.yy` of 46015) |

Our format code changed because of them, in `app/sheet/numfmt.ts`, written
here from Excel's behaviour (nothing was ported from IronCalc's formatter):

- a number now shows its first 15 significant digits and zeros after them
  (`fixed()`), so `#,##0.0000000000000000000` of 1/3 is `0.3333333333333330000`,
  and a format with many decimal places no longer printed `NaN`;
- commas before the decimal point and after the last integer digit scale
  (`#,##0,,.0#`), as they already did at the end of a format;
- `[hh]`, `[mm]` and `[ss]` pad their elapsed count;
- `mmmmmm` and longer month runs give the month's full name (only `mmmmm` is
  the first letter).

Series detection (no code ported; inputs and expected outputs only):

| Upstream path | Local destination | What was adapted |
|---|---|---|
| `base/src/user_model/sequence_detector.rs` (its `#[cfg(test)]` module, English cases) | `tests/headless/sheet_ironcalc_series.ts` | the numeric, float, suffixed and day/month name cases as data rows against `continued()`; the locale (de/fr/es), thousands-grouped (`1,000`) and `Jan.` cases left out |

Two gaps found in `app/sheet/series.ts`, fixed there (our own two-line changes,
not IronCalc's detector): a series of numbers keeps 15 significant digits, not
12 (`10000.4000000007, 10000.4000000008` went on as `10000.4`); and a text
number is padded only when written with a leading zero (`Q10, Q9` went on as
`Q08`; `Q-07, Q-09` still pads). Padded suffixes are kept as they were.

Deliberately different: IronCalc reads every selected cell and gives no series
when a later one breaks the step (and for `A1, B2`); `edit.fillSeries` reads the
first one or two cells only and copies when those two are not a series, so the
rows are cut to two cells and the no-series rows expect a copy.

Editing scenarios (no code ported; the scenarios and what they expect, translated):

| Upstream path | Local destination | What was adapted |
|---|---|---|
| `base/src/test/user_model/test_cut_n_paste.rs` | `tests/headless/sheet_ironcalc_editing.ts` | cut moves a cell's style and takes the target's on undo; absolute references and defined names (a cell's, a range's, a partly-overlapped range's) follow a cut; the dynamic-array cases left out (no spill cells here) |
| `base/src/test/user_model/test_paste_merged_cells.rs` | same | a merged block onto an identical merge, onto itself, a single cell into a merge, a paste overlapping its own source, cut onto a merge; undo and redo throughout; the CSV, external-replay and containment-replacement cases not translated (see below) |
| `base/src/test/user_model/test_autofill_rows.rs` (`styles`) | same | fill carries each source cell's look, undo and redo |

Formula compatibility workbooks (docs/comparisons/IRONCALC.md action 2):

| Upstream path | Local destination | What was adapted |
|---|---|---|
| `xlsx/tests/calc_tests/**/*.xlsx` (173 of the 221; the list is `tests/headless/lib/ironcalc_calc_<group>.ts`, one group per folder) | `third_party/ironcalc/fixtures/calc_tests/` (same relative paths) | nothing (byte for byte); read by the headless tests as base64 in `tests/headless/lib/ironcalc_calc_<group>.ts`, which `tests/headless/lib/mkironcalc_calc.py` generates from the copies, with the `sheet_ironcalc_calc_<group>.ts` test of each |

Each workbook was saved by Excel (`docProps/app.xml` says Microsoft Excel) with
the value it worked out stored beside each formula; the assertions are ours:
`lib/ironcalc_cached.ts` reads those stored values and compares what this app
works out, cell by cell. They are placeholder numbers, short strings and
function names; no third-party notice is in the repository for them. The 48 not
copied differ from Excel in more than twelve cells each (the largest are the
date, financial and statistical edge cases); they come in as their functions
are fixed. The 389 cells that still differ in the 173 are listed in
`lib/ironcalc_calc_known.ts`, each with how it differs (a listed cell that comes
right fails the test, so the list only shrinks). `NOW_TODAY`, `REPT` (it
exhausts memory) and the `WORKDAY` pair (it does not finish) are left out until
those are fixed.

What the workbooks found, fixed here (our own code, written from how Excel
behaves): the `_xlfn.`/`_xlpm.`/ANCHORARRAY storage of newer functions and cells
a dynamic array spilled into (headless `sheet_xlfn`); whole-column and whole-row
references (`sheet_wholecols`); TRUE and FALSE typed into a cell were text;
overflow was `#VALUE!` where Excel gives `#NUM!`; `0^0` was 1; and the
functions ROMAN (all five forms, checked against Excel over 2,000 numbers),
ARABIC, ISO.CEILING, SEC CSC COT SECH CSCH COTH ACOT ACOTH, COMBINA,
PERMUTATIONA, FACTDOUBLE, SQRTPI, MULTINOMIAL, SERIESSUM, SUMX2MY2, SUMX2PY2,
SUMXMY2, VARA VARPA STDEVA STDEVPA, SKEW.P, MMULT, MDETERM, MINVERSE and MUNIT
(`sheet_mathmore`); the binomial and Poisson distributions looped for a huge
count.

## Deliberate differences

Merged cells under paste and cut (`sheet_ironcalc_editing.ts`): IronCalc removes
a merge that lies wholly inside the paste target and replaces it with what is
pasted, and refuses only a merge that sticks out. Excel refuses a paste that
meets any merge it does not match ("To do this, all the merged cells need to be
the same size"); so do we, with nothing changed.

Number formats (the rows in `sheet_ironcalc_numfmt.ts` that carry a second
expected value and a reason):

- **A negative that rounds to zero** (`-0.00001` as `000.0`, `-1/3` as `0`):
  IronCalc shows no sign; Excel shows `-000.0` and `-0`. We keep showing no
  sign. Excel snaps near-zero results of subtraction to zero; our arithmetic
  does not, so a sign here would show `-0.00` for `=0.3-0.1-0.2`, where Excel
  shows `0.00`. Revisit with a snap in the calculation.
- **Doubled quotes in a format** (`0 """Millions"""`): IronCalc reads `""` as
  an escaped quote, as a formula's string does. In a number format Excel takes
  a backslash for a literal quote, and `""` is an empty string, so we show
  `42 Millions`.
- **`# ?/??` of 1.333**: IronCalc (ignored test, "not implemented") expects
  `1 1/3`. A `?` holds its place with a space, so Excel pads the unused
  denominator place: `1 1/3 `, as ours does.

XLSX: none. The former vertical-alignment difference (`optional_xf_id`) is gone: TODO 4.2
added vertical alignment, and the test now asserts A2 reads vertical top.
