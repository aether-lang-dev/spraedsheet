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
| `we-error` | 7 | this app gives an error where Excel stored a value |
| `excel-error` | 2 | Excel stored an error (`#NUM!`, `#VALUE!`) where this app gives a value |
| `value` | 20 | both give a value, and they differ |

Counted 2026-11-10 (67 cells with the 38 `legacy`, in 23 workbooks; 199 on 2026-11-06, 379 on 2026-10-24); the
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

2026-11-06, 6 cells: TREND and GROWTH take their fourth argument, const (FALSE fits through the
origin, y = m x, and for GROWTH b = 1; a fourth argument was #VALUE!, as one too many),
and check their arguments as Excel does where they paired the cells off and passed over the odd one
(as SLOPE and the like do, rightly): a cell of known_y or known_x that is not a number (text, a blank,
TRUE) is #VALUE!, known_y read first; known_x of another size than known_y is #REF! (it was #N/A); a
GROWTH of a y that is not above 0 stays #NUM!. `line`, `functions.ts`; cases in `examples.ts`.
TREND_GROWTH.xlsx E2, J2 (const), D10, E10 (text, so a cell passed over made a value), G10, I10
(sizes). (Several known_x columns: 2026-11-10, above.)

2026-11-10, 39 cells (maths, statistics, finance; each case failed first):
QUOTIENT and MULTINOMIAL refuse TRUE and FALSE (typed or in a cell) as #VALUE!, and so does SERIESSUM's x, n and m;
`POWER(-8, 1/3)` and `(-8)^(1/3)` are -2: a negative base to 1/3, 1/5, -1/7 ... (1/y an odd integer to within 2^-48)
is the real odd root, where any other fraction stays #NUM! (the rule was checked against the stored values and
IronCalc's `power`, which does the same, MIT); COUNT passes over an error value, in a range, a cell or typed
(`COUNT(#N/A)` is 0), where SUM stops at it; FREQUENCY gives its counts in the order the bins are written, each
from the bin next below it (it sorted them); POISSON of a negative x is #NUM! (-0.000001 was cut to 0); PROB passes
over text and blanks in its ranges and is #DIV/0! with no numbers left; DDB takes a fractional period and life
(book value before less after, from the definition); the XLSX importer reads a defined name's `_xlfn.LAMBDA(_xlpm.x, ...)`
as a LAMBDA (`gelu(A1:C1)`, 5 cells of LAMBDA.xlsx and defined_names.xlsx B5, YEARFRAC_DATEDIF C30 and C31 and N.xlsx B20
came right with it), reads `SINGLE(x)` as `@x`, and a cell of an array formula that carries an empty `<f ca="1"/>`
as the array's, not a value of its own (POWER.xlsx, 3 cells). XNPV takes only numbers among its values (a blank,
text or TRUE is #NUM!, where XIRR's text is #VALUE!) and a date below 0 is #NUM! in both (XNPV.xlsx Q7, Q8, Q9, Q11, R11).
TREND and GROWTH take several known_x columns (a column of y) or rows (a row of y), by the least squares LINEST
does, with new_x of as many predictors (#REF! if not); no cell of the ledger had one, the cases are in `examples.ts`.

2026-11-10, 23 cells (text functions and text read as numbers or dates): FIND from a start under 1
is #VALUE!; SEARCH reads a lone `~` at the end of its pattern as nothing, so it finds the start;
ARRAYTOTEXT's concise form runs the rows on with ", " (the semicolon is the strict form's, in braces);
VALUE of TRUE or FALSE is #VALUE!; a cell of empty text in a file (an empty shared string) is kept as
text, not dropped as blank, so VALUE of it is #VALUE! (a blank cell's is 0) and MATCH finds the "" in
it; a lookup for an empty cell looks for 0 (`MATCH(blank, ...)`); a number made text (`&`, CONCAT,
VALUETOTEXT) has 15 significant digits and is plain decimal while that takes 20 characters or fewer
(`0.000000007123456`; it had 12 digits and JavaScript's switch to an exponent from 1e-7), else
`1E+20` or `1E-20`; text read as a date: a month named (`2-Feb-2026`, `15 January 2026`,
`Jan 15, 2026`), month first with dashes (`10-12-1999`), spaces round `/` and `-`
(`2026 - 01 - 01`), a two-digit year 00 to 29 in the 2000s and 30 to 99 in the 1900s; a time after a
date is read with it (`VALUE("2026-01-01 12:00")` is 46023.5, it was the time alone) and a time of 24
hours or more carries whole days (`DATEVALUE("2026-01-01 25:00:00")` is 2 January); DATEVALUE of a
number is #VALUE!; TIMEVALUE reads `4 pm` and `4 : 35`, and 24:00, 23:60 and 23:59:60 come round to
0; a number too big to hold (`1e400`) is text, not infinity (VALUE, a COUNTIF criterion, a sum's text);
LEFT of an error with a negative count is that error. `dates.ts` (`stamp`, `secondsOf`) and
`values.ts` (`numberText`) read the rules from IronCalc's `date_and_time.rs` and `number_format.rs`
(third_party/ironcalc/PROVENANCE.md); cases in `examples.ts`. The cells: FIND_SEARCH.xlsx (3),
ARRAYTOTEXT.xlsx, STRING_HANDLING.xlsx H7, T_VALUE_VALUETOTEXT.xlsx (5), DATEVALUE_issue_1379.xlsx (4),
DATEVALUE_issue_1380.xlsx, TIMEVALUE.xlsx (3), TIMEVALUE_issue_1402.xlsx, MATCH.xlsx C18, and
arithmetic.xlsx (3).

2026-11-10, 6 cells: FORMULATEXT and ISFORMULA read what their cell holds as typed, not its value, so
that cell is no cycle through them: `=FORMULATEXT(A1)` in A1 is its own text, `E1 =CONCATENATE(FORMULATEXT(F1), "...")`
with `F1 =E1` is no cycle, and what reads such a cell reads the text (`=E1-6+3` is #VALUE!). They are
worked out again when the cell is retyped, in no order with it (`Sheet`'s `#inputsOf`, `excel.ts`'s
`INPUT_READERS`); a cycle of values is still one. FORMULATEXT.xlsx, all six; headless `sheet_inputreads`.

2026-11-10, 7 cells: a file's LAMBDA is read whole: an optional parameter is stored as `_xlop.b` (read as
`[b]`, so ISOMITTED sees it), and a defined name that holds a LAMBDA (`CountWords`, `Temperature`) is
stored with its `_xlfn.` and `_xlpm.` prefixes, which are taken off as a cell's formula's are, so calling it
is no longer #NAME?. ISOMITTED.xlsx A6, A7; LAMBDA.xlsx F2, B5 to B8; headless `sheet_xlfn`.

2026-11-10, 27 cells (references, INDEX, the postfix %): the range operator takes a call that makes a
reference as a corner, `A1:OFFSET(A1, 4, 0)`, `A1:INDEX(B:B, 5)`, `OFFSET(...):OFFSET(...)`,
`CHOOSE(...):CHOOSE(...)`, `INDIRECT("A1"):INDIRECT("A3")` (the Expr `span`: the rectangle of the two,
its cells told the Env as OFFSET's are); INDEX, CHOOSE and IF of a reference are references (`ISREF(INDEX(K1:K2, 1))`
is TRUE, so is `ISREF(IF(TRUE, A1, 5))`); a LET name or LAMBDA parameter given a reference stays one (ISREF,
ROW, OFFSET see it) and one given an error holds it unraised (`LET(x, 1/0, 5)` is 5, `ISREF(x)` of an
error FALSE); INDEX takes the area argument, a row or column of 0 or left out is the whole of it
(`INDEX(array, , )` is the array, `INDEX(array, 0, 1)` a column), a negative number is #VALUE! and an area
past the first #REF!; the formula's own cell in the branch of IF or CHOOSE not taken is no cycle
(`=CHOOSE(2, B40, "Not a cycle")` in B40); and the postfix `%` (it was a parse error: `=A3*50%`,
`B1:B4%`, `{10,20}%`) divides by 100 and binds tighter than ^. CHOOSE.xlsx (2), INDEX.xlsx (6),
ISREF.xlsx (7), ISREF_ISFORMULA_ISODD_ISEVEN.xlsx F8, error_broadcast.xlsx E76, MINUS_MINUS.xlsx (2),
percentage.xlsx, range_operator.xlsx (7); examples.ts and headless `sheet_references`. Not done: ROW,
OFFSET and the like of a LET (a name that comes to a reference by LET's body, `ROW(LET(...))`).

2026-11-10, 4 cells (names): a name of more than three letters before its digits is a name, not a
cell's address (`named1`, `Named2`, as a file's defined names have them: the lexer made them cells off the
sheet, #REF!, and `names.define` refused them); a quoted sheet name may hold a doubled quote
(`'Third ''Sheet'' (3)'!B3`). defined_names_for_unit_test.xlsx (3), quotes.xlsx; headless `sheet_names`.

2026-11-10, 31 cells (lookups, criteria, arrays): `SINGLE(x)`, which a file keeps for `@x`, is the implicit
intersection (N.xlsx, YEARFRAC_DATEDIF.xlsx C30 and C31, POWER.xlsx `implicit_intersection!C2`); an error beside an
array is that error in each of its cells (an operand that raised an error took the whole array with it), so IFERROR
and IFNA, which work cell by cell over an array, answer `SUM(IFERROR(A1:C1 / NA(), 2))` with 6
(error_broadcast.xlsx E68); a VLOOKUP or HLOOKUP table typed into
the call is text, #VALUE!, or a number or TRUE, #N/A (HVLOOKUP.xlsx, 4); COUNTIF, SUMIF and AVERAGEIF read `*`,
`?` and `~` in a text criterion for = and <> as wildcards, and an error as a criterion is the cells that are that
error (`COUNTIF(B2:B10, NA())` counts the #N/A, SUMIF's sum of them is #N/A) (SUMIF_AVERAGE_IF.xlsx, 9); an
array's empty cells are 0 where it lands (`=A1:A3` over a blank, SORT's blank rows), and SORT orders errors by
Excel's number (#DIV/0!, #VALUE!, #REF!, #NAME?, #NUM!, #N/A; descending, the reverse) (SORT.xlsx, 3); XMATCH's
match_mode 3 is a regular expression (a text looked for is the pattern, any other value finds nothing) and its
lookup array is one row or one column (XMATCH.xlsx, 5, and XMATCH_arrays.xlsx). MIN_MAX_extras.xlsx (3) and
POWER.xlsx's `inside_a_lambda!A3` had come right with a LAMBDA's parameter keeping its reference, and come off here
(POWER's A4 is now `legacy`). Cases in `examples.ts` and headless `sheet_references`.

2026-11-10, 1 cell: the value of a call is typed as a number, text or TRUE typed into the call is (a
function's own text counts where a cell's is passed over): `MAX(FILTER(L1:L2, M1:M2, ""))` is #VALUE!, as
`MAX("")` is, and `SUM(LEFT("7", 1))` is 7; but not when the call gives a reference (`IF`, `CHOOSE`, `INDEX`,
`OFFSET`, `INDIRECT` of one: `SUM(IF(TRUE, A1))` passes over a text A1) nor `IFERROR`, `IFNA`, `IFS`, `SWITCH`
and `XLOOKUP`, which may pass one through (`Arg.typed` by `#arg`, `evaluate.ts`). ISREF.xlsx D2.

2026-11-10, 4 cells and a correction (errors a file keeps as rich values): the file format has no spelling for
#CALC! or #SPILL!: such a cell is `t="e"` with `<v>#VALUE!</v>` and a `vm`, and its rich value (`xl/richData`,
read by `richErrorsOf` in `xlsxparts.ts`) holds the real error, Excel's ERROR.TYPE less one (13, #CALC!; 8,
#SPILL!). The importer now reads an error constant that way (ERROR.TYPE.xlsx A8 and A14 are #CALC! and #SPILL!, so
B8 and B14 are 14 and 9), and the ledger's reader takes the stored error from it. That also showed the cells
this ledger had read as #VALUE! to be #CALC!: FILTER that finds nothing (FILTER_ERRORS.xlsx E1, ISREF.xlsx D7),
and, which the entries of 2026-11-05 above said were #VALUE!, DROP and TAKE that leave nothing and SEQUENCE of
0 rows or columns (`SEQUENCE(0.3)` too): #CALC! now; SEQUENCE of a negative number stays #VALUE! (SEQUENCE.xlsx
C18). (MAKEARRAY of no rows stays #VALUE!, as Microsoft's documentation of it says.) Written files still carry
`#CALC!` and `#SPILL!` as plain error text, which the format does not have.

## Intended (found while looking)

- `T_VALUE_VALUETOTEXT.xlsx` C15 (`VALUE("2/3/2023")`) and I5 (`VALUE("10-12-1999")`): the workbook
  was saved where dates are written day first (44,987 is 2 March 2023, 36,504 is 10 December 1999);
  this app reads m/d/y, as en-US does, so it answers 3 February and 12 October. Kept, kind `value` in
  the ledger.

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
- `spill_operator.xlsx` G2 (`SUM(C2#)` where C2 holds a legacy array formula `{=123}` over C2:C5): this
  app's array formula of a scalar fills one cell, so `C2#` is #REF!; Excel's fills the four. I2
  (`ANCHORARRAY(OFFSET(A1, 1, 1))`): the spill reference of a call has no spelling in this app's formulas
  (`A1#` takes a cell), and a function for it would be one more name to list everywhere; not added.
