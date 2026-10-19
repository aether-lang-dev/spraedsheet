# Grist: actions

**Licence and reuse:** Apache-2.0. Paul's decision: ideas only, no source
copied, whatever the licence allows. No Python, SQLite or Grist data structures.

Not in the code today (checked `app/sheet/`): no record identity independent of
row position (only the auto-ID column, `autoId` in `control.ts`/`file.ts`), no
reference fields, no linked views. TODO.md is fully struck; these have no step
yet. Several belong to blog-entry-todo.md's object/sheet theme, none to B.1-B.7.

1. **Record identity for tables.** Stable table, field and record IDs, immutable
   and independent of row position; an optional, separately editable business key.
   Why: sort, rename and filtered edits then cannot hit the wrong record.
2. **Keyed reconcile of JSON imports.** Preview adds, changes, conflicts and
   optional deletes; keep missing / blank / null / delete distinct; commit as one
   command (one undo step). Builds on `tabular.ts` `fromJson`. Needs action 1.
3. **Reference fields and a linked detail view.** Orders -> Customer stores the
   identity, shows a label; rename keeps the link. Declare restrict / clear /
   cascade on delete before building it. Needs action 1.
4. **Record-aware invalidation.** A lookup that matches nothing must recalculate
   when a matching record is inserted. Measure before column-level optimising.
   Needs actions 1 and 3.

Constraints: keep free-form worksheets; a calculated column needs a stated
override policy; no collaboration server or permissions system for this.

Proof: import Customers and Orders from JSON, re-import by key without
duplicates, rename a customer, sort and edit Orders in a filtered view, undo,
reopen; the links and totals hold, through grid, REXX and MCP.
