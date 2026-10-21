# IronCalc: actions

**Licence and reuse:** MIT OR Apache-2.0, at our option. Code, fixtures and
tests may be copied or ported with attribution; we take it under MIT. Provenance
is in `third_party/ironcalc/PROVENANCE.md`, with `LICENSE-MIT` beside it. This is
the only not-ours project we copy from.

Done (78aa9fb..43df75b), recorded in PROVENANCE.md: the two XLSX regressions,
the number-format cases, the series-detection cases.

## Remaining reuse candidates

1. **Editing regressions:** selected scenarios from `base/src/test/user_model/`
   (autofill, cut/paste, merged cells, arrays, undo), translated to the headless
   harness. Tests feature interactions.
2. **Formula compatibility workbooks:** a curated subset of
   `xlsx/tests/calc_tests/`, copied byte for byte, assertions ours.
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
