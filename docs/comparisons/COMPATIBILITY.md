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
| `we-error` | 83 | this app gives an error where Excel stored a value |
| `excel-error` | 41 | Excel stored an error (`#NUM!`, `#VALUE!`) where this app gives a value |
| `value` | 43 | both give a value, and they differ |

Counted 2026-11-06 (205 cells with the 38 `legacy`, in 61 workbooks; 379 on 2026-10-24); the
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

2026-11-05, 38 cells: AND, OR and XOR pass over text in a cell ("TRUE" there is no TRUE, so
`AND(A7)` with nothing else is #VALUE!); an IFS condition that is text is #VALUE!; LOG to base 1
is #DIV/0!; SEQUENCE of fewer than one row or column is #VALUE!, and its rows left out are 1;
SUMSQ that overflows is #NUM!; BIN2HEX, BIN2OCT and the other conversions refuse places outside 1
to 10 even for a negative, which ignores them. The earlier cases `LOG(8, 1)` = #NUM! and
`SEQUENCE(0)` = #CALC! were this app's guesses; Excel's stored values say otherwise, and they are
corrected.

2026-11-05, 17 cells: MATCH, VLOOKUP and HLOOKUP looking for an exact match take `*`, `?` and
`~` in the text looked for as wildcards (they compared it as it is); DROP and TAKE that leave
nothing are #VALUE! (the earlier cases and `sheet_more` said #CALC!, this app's guess).

2026-11-05, 9 cells: the typed-argument rule reaches SMALL, LARGE and RANK (`SMALL("3.4", 1)` is
3.4, `SMALL("Hello", 1)` #VALUE!) and the A forms (`MINA(4, "-1", TRUE)` is -1,
`MINA(1, "Hola")` #VALUE!).

2026-11-05, 17 cells: an error value typed into a cell is that error (`#N/A` typed is NA()'s, in any
case; `#NULL!` and `#GETTING_DATA` are Excel's too, ERROR.TYPE 1 and 8), and the XLSX importer keeps
text that reads as TRUE, FALSE or an error as text (a cell holding the text "true" was made TRUE);
a TRUE/FALSE argument of a function (cumulative ...) reads the text "TRUE" and "FALSE" as those.
Headless `sheet_typederrors`. ERROR.TYPE.xlsx B8 and B14 now differ as values: Excel stored #VALUE! in
the file for errors that are #CALC! and #SPILL! (rich values), which the file does not say.

2026-11-06, 8 cells: arguments are read left to right and the first failure answers. A function
that refuses TRUE and FALSE (WEEKDAY, EDATE, DAYS, YEARFRAC ...) looked through every argument for an
error value first, so `WEEKDAY("1/1/10000", 1/0)` gave the second argument's #DIV/0! where Excel
stores the first argument's #VALUE!. The look now stops at the first error value, which the function
meets itself when it reads that argument (after the failures of those before it); a TRUE or FALSE
before it is still #VALUE!. (`call`, `functions.ts`; the case is in `examples.ts`.) The ledger's
WEEKDAY_WEEKNUM_ISOWEEKNUM_issue_1502.xlsx (8 cells) is off it; no other workbook's cells depended on it.

2026-11-06, 2 cells: ACCRINT counts a quasi-coupon period lived through whole as 1 (A = NL),
where it divided that period's 30/360 days by its length: 31 August 2007 to 29 February 2008 is 179
days of 180 on the US count (a start on the 31st, an end on the last of February), which lost a
day in `ACCRINT(1 Mar 2007, 31 Aug 2008, 1 May 2008, 0.1, 1000, 2, 0)` (116.667; Excel 116.944).
(`accrued`, `bonds.ts`; the case is in `examples.ts`.) ACCRINT.xlsx B4 and B23.

## Intended (found while looking)

- `T_VALUE_VALUETOTEXT.xlsx` C15 (`VALUE("2/3/2023")`): the workbook was saved where dates are
  written day first (44,987 is 2 March 2023); this app reads m/d/y, as en-US does, so it answers
  3 February (it was #VALUE! before text dates were read). Kept, kind `value` in the ledger.

## Examined, not changed (still on the ledger, with the reason)

- `XNPV.xlsx` R2, R5 (`XIRR`): the two agree to 1e-8 of the rate; Excel stops its iteration
  sooner than this app does. Matching would mean copying Excel's stopping rule, which is not
  documented; the difference is under the displayed precision.
- `ACCRINT.xlsx` (7 cells left of 9; B4 and B23 were the day count, see the entry above): what Excel
  stored for the rest does not follow from the documented formula (par * rate / frequency * the sum
  over the quasi-coupon periods of A / NL), on any day count tried, so they stay.
  Basis 2: B33 (1 January to 1 April 2017, quarterly) is 86 days over 360 in Excel's answer, the span
  being 90; B36 (1 June to 1 September 2016) is 87 days, the span 92. Basis 3: B34 is 88.5 days over
  365 for the 90-day span. Basis 1: B38 (the same 2016 span) is 93 over 366 for 92 days, B39 (30 days,
  quarterly) 29 over 364, B32 59/90 + 29/91 where the days are 59/90 + 31/92; B37, basis 3 with the
  dates of B36, is right at 92/365, and B24 to B26 (2020) are right on bases 1 to 3, so a rule that
  moved A or NL by a day or more for these would have to be found from the cases themselves.
  B5 (calc_method FALSE) is 241 days of 180, from 31 August 2007, the first quasi-coupon date after
  the issue; B2 and B31 (FALSE too, and right) start at the issue because the settlement is in the
  issue's own period. Microsoft's text (FALSE: from first_interest to settlement) does not give
  that either, as the settlement is before the first interest. Not changed: three cases make a rule
  to guess at, not one to copy.
- `RATE.xlsx` F17, G17: Excel answers #NUM! where this app's iteration finds a root below -1
  (-1.87) or at -1; Excel's guess-driven Newton search does not reach it. Not changed until
  Excel's rule for refusing such roots is known.
