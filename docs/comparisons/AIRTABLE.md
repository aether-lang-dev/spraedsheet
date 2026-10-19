# Airtable: actions

**Licence and reuse:** proprietary. Ideas from its public documentation only
(support.airtable.com, checked 18 October 2026); no assets, no copying.

Most record/relation work is shared with Grist.md and is not repeated: record
identity, keyed reconcile, reference fields (Grist 1-3). Beyond it:

1. **Saved views over one table:** persisted filter/sort/group/column settings,
   several per table, over one record model. Cards/board/calendar later. No step
   yet; needs Grist 1.
2. **A generated record form** from the table's schema and the native field
   widgets; grid and form call the same commands and validation.
3. **Rollup semantics declared:** a total over related records uses all of them,
   the selected view, or an explicit predicate (Airtable's default ignores view
   filters). Make the choice inspectable. Needs Grist 3.
4. **Observable triggers.** `macro.trigger` runs on open and on change (TODO
   7.2). Missing: a run log (inputs, outcome, errors), transition semantics
   (fire on a committed change to Ready, no backfill of existing matches),
   self-trigger loop stop, idempotent retries, and a declared clock/timezone for
   scheduled work. Final committed state, not each write. No exactly-once claim
   for external effects. Runs only while the workbook is open.
5. **Saved views are not access control.** If sharing ever comes, check every
   data/command boundary including REXX and MCP. Sharing, public forms and hosted
   execution are out of scope for now.

Proof: a customers-and-orders workbook: import, reference, two saved views and a
detail pane editing one record, a per-customer total, a recorded parameterised
workflow triggered when an order becomes Ready.
