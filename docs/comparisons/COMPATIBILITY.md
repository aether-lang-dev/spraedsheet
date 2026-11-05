# Compatibility ledger: where this app and Excel differ

One place for what differs from Excel, in four classes, so a difference is
never just "a failing test": **intended** (we chose it), **bug** (we are wrong,
cause known), **unsupported** (not built), **unverified** (differs from the value
Excel stored, cause not examined). Seeded from `third_party/ironcalc/PROVENANCE.md`
"Deliberate differences" and the cells that IronCalc's Excel-authored workbooks
still show different (`tests/headless/lib/ironcalc_calc_known.ts`, which shrinks
as they are fixed; the test says when a listed cell comes right).

`formulas.functionStatus` (and `FUNCTIONINFO(name)`) gives each function's record
with the count of its still-different cells (`app/sheet/funcinfo.ts`; headless
`sheet_funcinfo` recounts it from the ledger, so the two cannot drift).

## Intended

- A formula Excel stored without an array flag is intersected implicitly there;
  this app spills it (38 cells, kind `legacy` in the ledger). Whether to insert
  `@` on import as Excel's loader does is Paul's decision (TODO.md).
- The deliberate differences listed in `third_party/ironcalc/PROVENANCE.md`.
- RexxJS formulas are not Excel's: a file written from a RexxJS workbook carries
  their values, not the formulas (the file report says so).

## Bugs

- None known and unfixed that are not in the next class: each cell that differs
  is first "unverified" until looked at; a cause found moves it here with a
  test.

## Unsupported

- Whatever the file report (`FILEREPORT`) lists as left out on import or export.

## Unverified (the ledger's cells, by how they differ)

| Kind | Cells | Means |
|---|---|---|
| `we-error` | 104 | this app gives an error where Excel stored a value |
| `excel-error` | 97 | Excel stored an error (`#NUM!`, `#VALUE!`) where this app gives a value |
| `value` | 57 | both give a value, and they differ |

Counted 2026-11-05 (296 cells with the 38 `legacy`, in 73 workbooks; 379 on 2026-10-24); the
list by workbook is the ledger file itself. `docs/function_status.md` gives the count per function.

## Taken off the ledger (fixed, each with the case that failed first)

2026-11-05, 25 cells: `YEARFRAC` basis 1 within a year counts a year of 366 days only when a
29 February falls in the span (it averaged the two years), which `ACCRINTM` and the bond
functions use too; `YEARFRAC` and `DATEDIF` of a negative serial are #NUM!; `DATEDIF` counts
whole days (a start later in the same day was "after" the end); `PMT`, `FV`, `PV`, `NPER` and
`RATE` read any type but 0 as 1 (payments at the start); `PMT` of a rate of -1 or less is #NUM!;
`FV` of a rate of -1 over negative periods is #DIV/0!; `IPMT` and `PPMT` take a fractional period
up to the one after the last; `DB` of a cost of 0 is 0 (it was NaN). Cases in
`tests/headless/lib/examples.ts`.

2026-11-05, 18 cells: a date is a serial from 0 to 2,958,465 (31 December 9999): `YEAR`, `MONTH`,
`DAY`, `EDATE` and `EOMONTH` of one outside it are #NUM!, and so is an `EDATE` or `EOMONTH` that
comes to one outside it; a number written as text is that serial to the date functions
(`EOMONTH("44987", 1)`). The bound is the 1900 system's (a 1904 workbook's last date is 1,462
days earlier; not distinguished).

2026-11-05, 22 cells: text where a number goes is read as Excel reads it (`textNumber`,
`values.ts`): thousands separators, "8%", "$100", "€100", "100€", " €100 ", "(5)", a date or a
time as it stands; and no longer JavaScript's spellings ("0x10", "Infinity"). A date is checked
to be one ("30/2/2010" is not). EXP, SIGN, VALUE, `--"50%"`, CONFIDENCE and an IF over an array
came right with it.

2026-11-05, 13 cells: a value typed into the call is counted by SUM, PRODUCT, AVERAGE, MIN, MAX,
COUNT and the other functions over a list of numbers, where the same value in a cell is passed
over: `SUM("7", TRUE)` is 8 and `SUM("seven")` #VALUE! (COUNT passes the text over), as Excel
documents. An argument is typed when it is a number, text, TRUE or FALSE, a negation or an
operation (`Arg.typed`, set by the evaluator).

## Intended (found while looking)

- `T_VALUE_VALUETOTEXT.xlsx` C15 (`VALUE("2/3/2023")`): the workbook was saved where dates are
  written day first (44,987 is 2 March 2023); this app reads m/d/y, as en-US does, so it answers
  3 February (it was #VALUE! before text dates were read). Kept, kind `value` in the ledger.

## Examined, not changed (still on the ledger, with the reason)

- `XNPV.xlsx` R2, R5 (`XIRR`): the two agree to 1e-8 of the rate; Excel stops its iteration
  sooner than this app does. Matching would mean copying Excel's stopping rule, which is not
  documented; the difference is under the displayed precision.
- `ACCRINT.xlsx` (9 cells): Excel's accrual by quasi-coupon periods counts the days of each
  period on the basis differently from this app's (30/360 from 1 March 2007 to 1 May 2008 is
  420 days here, 421 in Excel's answer); not yet worked out from the documentation.
- `RATE.xlsx` F17, G17: Excel answers #NUM! where this app's iteration finds a root below -1
  (-1.87) or at -1; Excel's guess-driven Newton search does not reach it. Not changed until
  Excel's rule for refusing such roots is known.
