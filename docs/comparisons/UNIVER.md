# Univer: actions

**Licence and reuse:** Apache-2.0, but Paul decided ideas only: no source
copied. No canvas renderer, plugin/DI framework or JS UI is wanted.

Done, so dropped: one command registry reached from panels, REXX, bus and MCP;
labels, params, journal and grouped undo; `journaled` and `unstepped` flags
(`commands.ts`); recording as REXX and replay with arguments (TODO 7.1, 7.2).

1. ~~**Classify commands by effect.** Durable edit, view change, external effect,
   query: one field in the registry, consumed by recording, help and MCP
   discovery (today `journaled`/`unstepped` only separate two of the four).~~ Done
   2026-10-24: every command in `commands.ts` carries `effect: "edit" | "view" |
   "external" | "query"` (173 edits, 37 view changes, 20 external, 51 queries); the
   read view refuses exactly the edits (the prefix test `looksOnly` and the list
   `CHANGES_UNJOURNALED` are gone), and `catalogue()`, which the command bar and
   menus read, carries the effect and filters by it. Headless `sheet_effects` (it
   runs every command in the read view). Behaviour that changed with it: copy and
   cut are view changes, so a read view can copy; `macro.trigger`, `macro.stop` and
   `macro.run` are edits, so a read view refuses them (a trigger was let through).
   Not done: recording still records by `journaled`, and MCP discovery lists
   REXX functions (`described.ts`), not commands, so neither reads the field yet.
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
3. ~~**Feature participation in structural edits.** Require each feature (tables,
   charts, validation, conditional formats, names, merges) to show behaviour
   through insert, delete, copy, undo and reopen; one headless test that attaches
   all of them to a row insert.~~ Done 2026-10-21 with LibreOffice 1: headless
   `sheet_structural` attaches them all to a row insert and a column insert,
   undo, redo and reopen; names were the one feature that did not follow (fixed).
4. ~~**Extension lifetimes.** Triggers and macros (7.2) are saved with the
   workbook; make sure closing a workbook removes its triggers and panels'
   callbacks, with no stale workbook state in globals.~~ Done 2026-10-21: a
   macro and its trigger are the workbook's own and go when it closes (tested);
   one global did leak: the date system's epoch (`dates.ts`, one for the page).
   Switching to a workbook counted from the epoch of the last one set, so a 1900
   workbook shown after a 1904 one made 1904 serials. `Workbook.claimDates()` now
   runs whenever a workbook becomes the active one (`books.ts`). Headless
   `sheet_booksglobals`. The new clock (`workbook.clock`) is the page's by design.
5. ~~**Mutation journal (later).** Could cut snapshot cost and give blog-entry-todo
   B.4/B.5 their diff; do after the history-cost measurement (see IronCalc).~~ Done
   2026-10-25 (design and measurement: `docs/mutation_journal.md`): a step of undo is a
   list of journal entries (`journal.ts`), each a kind of state with what it was and
   what it became, and undo and redo apply them backwards and forwards; cells, notes,
   a sheet's names, formulas and tables, the workbook's names and cell styles, and
   widths, heights, pivots and outline levels are kept as the keys that changed. A
   one-cell edit on a 2,000-cell sheet held 2,021 map entries and holds 1; a script of
   fifty edits 3,043 and now 50 (`sheet_historycost`, `sheet_journal`). Left: recording at
   the mutation (the diff is taken when the command is done), and recording and MCP
   discovery reading `effect` (action 1's remainder). View changes stay undoable, as
   before: a decision for Paul (TODO.md).

Proof: record a panel command, replay it over the bus with another selection;
it hits the original target. Undo and reopen restore everything attached.
