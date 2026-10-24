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
3. **Lineage for group/pivot results.** Let a grouped result show its
   contributing rows; edit through a derived view only when the target is
   unambiguous; aggregates read-only. First fix `query.ts`'s textual grouping
   keys (`keyText`): define typed equality, null handling, duplicate keys, and
   live vs snapshot results.
4. **Deferred patches for outbound object changes.** Staged inserts, updates and
   deletes with an apply/discard step. Fits blog-entry-todo B.5 (hold a change
   for review, keep or take back); needs Grist 2.
