# The mutation journal

Univer 5 (`docs/comparisons/UNIVER.md`). Undo and redo are replayed from a
journal of entries, one per change to one kind of state, each holding what the
state was and what it became. This replaces a step of undo being a copy of the
whole workbook.

## Why

The measurement (`tests/headless/sheet_historycost.ts`, IronCalc 4): a step of undo
was a `Saved` snapshot, so a one-cell edit on a 2,000-cell sheet kept 2,021 map
entries, an inserted row 3,043, a script of fifty changes 3,043; a 100,000-cell
sheet ran out of memory at the 47th edit (now only bounded by `HISTORY_ENTRIES`).
A journal makes a step cost what it changed.

## Shape

`app/sheet/journal.ts`.

- **Slot**: one kind of state. A workbook's own (`names`, `active`, `name`,
  `setupScript`, `styles`, `theme`, `bookProtection`) or one sheet's (`pname`, and
  its `language`, `inputs`, `names`, `formulas`, `tables`, `notes`, and each field
  of its view: rules, charts, merges, hidden rows, widths, frozen panes,
  protection ...).
- **Entry**: one of
  - `set`: a slot, what it was and became, whole (small kinds, and kinds not yet
    migrated);
  - `map`: a map slot, the keys that changed with what each was and became (undefined:
    no entry);
  - `pages`: the sheets before and after, in order, and the sheets in only one of
    them (a deleted sheet comes back as it was, an added one is done again as it was).
- **Step**: a label and the entries, and its cost (map entries held).

`Saved` (`workbook.ts`) stays the currency between a command and the journal: the
runner takes the workbook as it was (`saved()`, shared with the previous
snapshot where unchanged, so cheap), the command runs, and `changed()` builds
the step from `stepOf(before, now)`: an entry for each slot that differs. A
script (a group) is one step from the workbook before it to the workbook after.
Undo applies the entries backwards to the workbook as it is now (`applyStep`), redo
forwards, and the result is put back by the same `#apply` as before, so what is
worked out again after an undo is unchanged. The step moves between the undo and
redo stacks; it is not copied.

Because an undo patches the slots the step touched, rather than putting a whole
snapshot back, a change made to some other slot outside the journal is no longer
wiped by an undo of a different step (the old snapshot restore did).

The diff is taken when the command is done, not recorded as each mutation is made:
the cost per command is one `saved()` (copying a changed map, as before), but nothing
is retained beyond the entries. Recording at the mutation, so that a command
costs what it changed, is the later step and needs no change to the journal's entries.

## Migration, one kind at a time

Every step keeps all existing undo tests green and adds to `sheet_journal.ts` and
`sheet_historycost.ts`. `PATCHED` in `journal.ts` lists the map slots kept as the keys
that changed; the rest are `set` whole.

1. Done 2026-10-25: the journal, with `set` for every slot, and **cells** (a sheet's
   inputs) as `map`.
2. Done: **notes** (and so looks, comments, links, conditional styles).
3. Done: a sheet's **names, formulas and tables**.
4. Done: the workbook's **names** and **cell styles**.
5. Done: the sheet's view maps: **widths, heights, pivots, outline levels**.

What stays `set` for good is a container a command replaces whole and that is small
(rules, charts, merges, scenarios, drawings, slicers, edit ranges), a set of hidden
rows, and the scalar kinds (name, setup script, theme, protection, tab colour ...).
The active sheet is not an entry: undo goes to the sheet the change was made on,
and redo to the one undo was made from, swapped with each (`Step.active`).

## Decisions kept as today

- **View changes are undoable** (freeze panes, hidden rows and columns, widths,
  gridlines, tab colour): they were in `Saved`, so they are in the journal. Not changed.
  Whether they should be (Excel does not undo a scroll or a split, but does undo a
  hide and a width) is Paul's; see TODO.md.
- Not in the journal, as not in `Saved` before: the clipboard, traces, watches,
  Show Formulas, Circle Invalid, macros, triggers, checks, RexxJS's own state.
- Saved files are untouched: the journal is memory only, never written.

## Measurement

`sheet_historycost`, N = 2,000 cells, map entries held by one step, the snapshot
(HEAD before the journal) against the journal:

| change | snapshot | journal |
|---|---|---|
| one-cell edit | 2,021 | 1 |
| the next one-cell edit | 2,022 | 1 |
| a look over ten cells, after one | 31 | 10 |
| a 1,000-cell paste (with ten notes) | 31 | 1,010 |
| an inserted row (every cell moves) | 3,043 | 3,008 |
| a script of fifty edits | 3,043 | 50 |
| a hundred one-cell steps | about 200,000 | 100 |
| a defined cell style (21 in the gallery) | 43 | 1 |
| a second name, a second table | 3 each | 1 each |
| a column widened or a row made taller | 7 | 1 |

(The paste's 31 was the snapshot sharing the cells with the step before; the journal
counts the 1,000 cells it set.) A structural edit that moves every cell is the one
case that stays about the size of the sheet, as the cells did move. The history
budget (`HISTORY_ENTRIES`) is kept for it: twenty inserted rows on 2,000 cells
keep the newest four or five steps under a budget of 10,000.
