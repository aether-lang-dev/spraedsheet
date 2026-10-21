# IronCalc: actions

**Licence and reuse:** MIT OR Apache-2.0, at our option. Code, fixtures and
tests may be copied or ported with attribution; we take it under MIT. Provenance
is in `third_party/ironcalc/PROVENANCE.md`, with `LICENSE-MIT` beside it. This is
the only not-ours project we copy from.

Done (78aa9fb..43df75b), recorded in PROVENANCE.md: the two XLSX regressions,
the number-format cases, the series-detection cases.

## Remaining reuse candidates

1. ~~**Editing regressions:** selected scenarios from `base/src/test/user_model/`
   (autofill, cut/paste, merged cells, arrays, undo), translated to the headless
   harness. Tests feature interactions.~~ Done 2026-10-21 for cut and paste,
   merged cells and fill (headless `sheet_ironcalc_editing`, recorded in
   PROVENANCE.md). It found three gaps, fixed: **Paste** carried only the
   cells' inputs, not their look, number format, validation, hyperlink or
   comment (Excel's Ctrl+V carries them; a blank cell pasted clears them); a
   paste, cut or fill into **merged cells** wrote values under the merge, or
   made a second identical merge; and **Fill Down/Right/Up/Left and Series**
   did not carry the source cell's look. Merged cells follow Excel: a paste or
   cut that meets a merge it does not match is refused ("all the merged cells
   need to be the same size"); one cell pasted into a merge goes to its anchor;
   a merged block pasted makes its merge. Left, not yet scenarios here: a fill
   that replicates merges, CSE and dynamic-array cut/paste (spill cells), the
   column-width and row-height undo cases of `test_delete_row_column_formatting`.
2. **Formula compatibility workbooks:** a curated subset of
   `xlsx/tests/calc_tests/`, copied byte for byte, assertions ours. Done in
   part 2026-10-21 (headless `sheet_ironcalc_calc_*`): Excel stored its own
   value beside every formula in those workbooks, so they are an oracle. Of 221
   workbooks (71,000 formula cells), 173 are in, with a ledger of the 389 cells
   that still differ (`lib/ironcalc_calc_known.ts`); at the start 48,000 of the
   71,000 cells differed, now 4,033, in 100 workbooks. What they found is in
   PROVENANCE.md: file interchange (`_xlfn.`), whole columns and rows, typed
   TRUE/FALSE, overflow and `0^0`, and twenty-five missing functions. Left:
   the 48 workbooks with more differences (date edge cases where Excel gives
   `#NUM!`, financial and statistical corners, TEXT and SUBSTITUTE options,
   WEEKDAY's return types 11 to 17), and one cause the ledger calls *legacy*:
   a formula Excel stored with no array flag is intersected implicitly
   (`=H34:H37` in row 36 is H36), where this app spills it; an importer that
   writes `@` where Excel's loader would is not done.
3. **Small ports, only where a test shows a gap:** a routine from
   `base/src/formatter/` for `numfmt.ts` (the cases are in; no code ported yet).
   Port the smallest unit with its cases.

## Not reuse

4. **Measure history cost** before choosing patches over snapshots (the workbook
   keeps whole snapshots, `HISTORY` deep): one-cell edit, large paste, styles,
   structural edits, scripted group, redo memory. Patches only if it pays.
5. **Original workbook corpus:** open, calculate, edit, save, reopen, comparing
   semantics (not ZIP bytes) and unsupported-feature reports.
6. ~~**Locale/time controls:** `dates.ts` `now()` reads the host clock and zone;
   make the clock injectable for tests and keep saved results independent of host
   locale.~~ Done 2026-10-21: `workbook.clock at=` (`SETCLOCK("2026-10-21 09:30")`,
   a date or a serial) pins what `TODAY()` and `NOW()` read, none gives the host's
   clock back, and the volatile cells are worked out at once; `GETCLOCK()`. A
   session's setting, never saved, so a saved file is independent of both clock
   and locale (a formula's cached value is not kept). Headless `sheet_clock`.

## Attribution rules

- Check each fixture's own provenance; do not assume the repo's MIT covers it.
- Keep IronCalc's copyright and full MIT notice with the reused material.
- Record upstream repo, full commit, paths, licence and local destinations in
  PROVENANCE.md, in the commit that brings it in.
- Upstream workbooks stay byte-for-byte under `third_party/ironcalc/`, apart from
  our own cases; mark adaptations. Upstream expected values are not an oracle:
  Excel's documented behaviour decides, differences go under "Deliberate
  differences".
