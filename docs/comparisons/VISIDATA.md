# VisiData: actions

**Licence and reuse:** GPL-3.0. Ideas only: no Python or replay scripts copied.

Done, so dropped: recording a session as editable REXX (TODO 7.1, `recorder.ts`);
replay with `ARG` parameters and open/change triggers (7.2); replay as a test
(headless-vs-interactive, 7.7).

1. **Record by identity where possible.** Recorded lines use record or key IDs
   once tables have them (see Grist 1), ranges elsewhere; a missing or duplicate
   key fails clearly instead of editing whoever now sits in the old row. Follow-up
   to 7.1; needs Grist 1.
2. ~~**Replay preview.** Show a replay's targets and failures before applying it
   (parameters already exist, 7.2). State whether failure restores the workbook;
   macro runs are one undo step today.~~ Done 2026-10-24: `macro.preview name=
   [args=]` (`PREVIEWMACRO(name, args)`) runs the macro inside a group and takes it
   back whole, answering each command it ran, the cells it changed and whether it
   was refused, and saying that a run which fails keeps what it did before the
   failure as one step of undo (the macro goes on past a refused command, as ADDRESS
   does). Headless `sheet_macropreview`.
3. ~~**Lineage for group/pivot results.** Let a grouped result show its
   contributing rows; edit through a derived view only when the target is
   unambiguous; aggregates read-only. First fix `query.ts`'s textual grouping
   keys (`keyText`): define typed equality, null handling, duplicate keys, and
   live vs snapshot results.~~ Done 2026-10-25 in two parts. Typed keys
   (`groupkey.ts`, headless `sheet_groupkeys`): a key is a value with its type,
   so 1 and "1" are two groups, TRUE is not "TRUE", text matches without regard to
   case as `=` does (a group is labelled by the spelling first seen), -0 is 0, an
   error is its code, and a blank is a group of its own (not 0, not ""). `=` equates
   a blank with 0 and with "" but not those two with each other, so GROUPBY's old
   comparison (a row against a group's first row) gave a different answer by row
   order; GROUPBY, PIVOTBY and GROUP_BY now share one hash key. Duplicate keys are
   one group wherever their rows are. Live versus snapshot: a query or grouped
   formula is live (recomputed with the cells it reads, never stored); Paste
   Values is the snapshot. Lineage is LibreOffice 4 (`group.explain`, LINEAGE,
   EXPLAIN): the sheet rows behind each group, so a script edits the source rows
   (`cells`), and an aggregate is a result, never a cell to edit.
4. **Deferred patches for outbound object changes.** Staged inserts, updates and
   deletes with an apply/discard step. Fits blog-entry-todo B.5 (hold a change
   for review, keep or take back); needs Grist 2.
