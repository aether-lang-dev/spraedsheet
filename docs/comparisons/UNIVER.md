# Univer: actions

**Licence and reuse:** Apache-2.0, but Paul decided ideas only: no source
copied. No canvas renderer, plugin/DI framework or JS UI is wanted.

Done, so dropped: one command registry reached from panels, REXX, bus and MCP;
labels, params, journal and grouped undo; `journaled` and `unstepped` flags
(`commands.ts`); recording as REXX and replay with arguments (TODO 7.1, 7.2).

1. **Classify commands by effect.** Durable edit, view change, external effect,
   query: one field in the registry, consumed by recording, help and MCP
   discovery (today `journaled`/`unstepped` only separate two of the four).
2. ~~**Record resolved context.** A toolbar action on "the selection" should record
   its resolved sheet and range, so a replay does not target whatever is selected
   later. Check `recorder.ts` first; this is a 7.1 follow-up, no step yet.~~ Done
   2026-10-21. Checked: ranges were already explicit in the recorded line (the
   panels pass `v.range()`), but two things were not. A macro recorded on one
   sheet ran on whichever sheet was active when it was run: with several sheets
   a recording now starts `sheet.activate name=<the sheet recorded on>` (unless
   its first command already switches sheet), and the run leaves that sheet
   active. And `insert.shape`, `insert.picture` and `insert.textbox` without an
   `at=` took the selection at replay time: a command can now return `resolved`
   parameters (`Outcome.resolved`), which the recorder keeps, so the line says
   `at=C3`. Headless `sheet_recordcontext` (failed before); `sheet_macros`'
   expected programs gained the activate line.
3. **Feature participation in structural edits.** Require each feature (tables,
   charts, validation, conditional formats, names, merges) to show behaviour
   through insert, delete, copy, undo and reopen; one headless test that attaches
   all of them to a row insert.
4. ~~**Extension lifetimes.** Triggers and macros (7.2) are saved with the
   workbook; make sure closing a workbook removes its triggers and panels'
   callbacks, with no stale workbook state in globals.~~ Done 2026-10-21: a
   macro and its trigger are the workbook's own and go when it closes (tested);
   one global did leak: the date system's epoch (`dates.ts`, one for the page).
   Switching to a workbook counted from the epoch of the last one set, so a 1900
   workbook shown after a 1904 one made 1904 serials. `Workbook.claimDates()` now
   runs whenever a workbook becomes the active one (`books.ts`). Headless
   `sheet_booksglobals`. The new clock (`workbook.clock`) is the page's by design.
5. **Mutation journal (later).** Could cut snapshot cost and give blog-entry-todo
   B.4/B.5 their diff; do after the history-cost measurement (see IronCalc).

Proof: record a panel command, replay it over the bus with another selection;
it hits the original target. Undo and reopen restore everything attached.
