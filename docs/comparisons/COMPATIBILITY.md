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
| `we-error` | 143 | this app gives an error where Excel stored a value |
| `excel-error` | 132 | Excel stored an error (`#NUM!`, `#VALUE!`) where this app gives a value |
| `value` | 75 | both give a value, and they differ |

Counted 2026-10-24; the list by workbook is the ledger file itself.
