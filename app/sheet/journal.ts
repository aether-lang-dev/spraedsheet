// The mutation journal (docs/mutation_journal.md): a step of undo is a list of
// entries, each one change to one kind of state with what it was and what it
// became, so a step costs what it changed rather than a copy of the workbook.
//
// Saved (workbook.ts) stays the currency between a command and the journal: a
// command is given the workbook as it was, and when it is done the difference
// between that and the workbook now is the step. Undo applies a step's entries
// backwards to the workbook as it is, redo forwards; the result is a Saved that
// the workbook puts back as it always has.
import type { Page, Saved, SavedPage } from "./workbook.ts";

// One kind of state: a workbook's own (page absent) or a sheet's.
export interface Slot {
  readonly page?: Page;
  readonly key: string;
}

export interface Entry {
  readonly slot?: Slot;
  readonly kind: "set" | "map" | "pages";
  // The keys a map kind of state changed, each with what it was and became
  // (undefined: no entry).
  readonly changes?: readonly (readonly [unknown, unknown, unknown])[];
  // What the state was and became (set), whole.
  readonly was?: unknown;
  readonly now?: unknown;
  // Sheets in order before and after (pages), and the sheets that are in only
  // one of them, as they were or became.
  readonly order?: readonly [readonly Page[], readonly Page[]];
  readonly extra?: ReadonlyMap<Page, SavedPage>;
}

export interface Step {
  readonly label: string | undefined;
  readonly entries: readonly Entry[];
  // The sheet undo (or redo) goes to: the one the change was made on, and
  // after an undo the one it was undone from, so redo comes back to it. They
  // swap with each undo or redo, as the active sheet of a saved workbook did.
  active: Page;
  // Map entries held by the step (what historyCost sums).
  readonly cost: number;
}

const SHEET = new Set(["language", "inputs", "names", "formulas", "tables"]);
const BOOK = ["names", "name", "setupScript", "styles", "theme", "bookProtection"] as const;

// The kinds of state that are maps and are journaled as the keys that changed,
// not whole (the rest are set whole): "book:" a workbook's own, "sheet:" a
// sheet's. Each kind moved here is one step of docs/mutation_journal.md.
const PATCHED = new Set<string>(["sheet:widths", "sheet:heights", "sheet:pivots", "sheet:rowLevels", "sheet:colLevels", "sheet:inputs", "sheet:notes", "sheet:names", "sheet:formulas", "sheet:tables", "book:names", "book:styles"]);
const patched = (slot: Slot) => PATCHED.has(`${slot.page ? "sheet" : "book"}:${slot.key}`);

const sizeOf = (v: unknown): number => v instanceof Map || v instanceof Set ? v.size : Array.isArray(v) ? v.length : 1;

// Whether two values of a kind of state are the same: containers by what they
// hold (a snapshot copies its containers, never what is in them).
function same(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  if (a instanceof Map && b instanceof Map) {
    if (a.size !== b.size) return false;
    for (const [k, v] of a) if (!b.has(k) || b.get(k) !== v) return false;
    return true;
  }
  if (a instanceof Set && b instanceof Set) {
    if (a.size !== b.size) return false;
    for (const v of a) if (!b.has(v)) return false;
    return true;
  }
  if (Array.isArray(a) && Array.isArray(b)) return a.length === b.length && a.every((v, i) => v === b[i]);
  return false;
}

function pageOf(s: Saved, p: Page): SavedPage {
  const sp = s.pages.find((x) => x.page === p);
  if (!sp) throw new Error(`journal: sheet ${p.name} is not in the workbook`);
  return sp;
}

function get(s: Saved, slot: Slot): unknown {
  if (!slot.page) return (s as unknown as Record<string, unknown>)[slot.key];
  const sp = pageOf(s, slot.page);
  if (SHEET.has(slot.key)) return (sp.sheet as unknown as Record<string, unknown>)[slot.key];
  if (slot.key === "notes") return sp.notes;
  if (slot.key === "pname") return sp.name;
  return (sp.view as unknown as Record<string, unknown>)[slot.key];
}

function put(s: Saved, slot: Slot, v: unknown): Saved {
  if (!slot.page) return { ...s, [slot.key]: v };
  const p = slot.page;
  return {
    ...s,
    pages: s.pages.map((sp) => {
      if (sp.page !== p) return sp;
      if (SHEET.has(slot.key)) return { ...sp, sheet: { ...sp.sheet, [slot.key]: v } };
      if (slot.key === "notes") return { ...sp, notes: v as SavedPage["notes"] };
      if (slot.key === "pname") return { ...sp, name: v as string };
      return { ...sp, view: { ...sp.view, [slot.key]: v } };
    }),
  };
}

// The step from `before` to `after`: an entry for each kind of state that
// differs, and one for the sheets added, deleted or moved.
export function stepOf(before: Saved, after: Saved, label: string | undefined): Step {
  const entries: Entry[] = [];
  const add = (slot: Slot, was: unknown, now: unknown) => {
    if (same(was, now)) return;
    if (patched(slot) && was instanceof Map && now instanceof Map) {
      const changes: [unknown, unknown, unknown][] = [];
      for (const [k, v] of was) if (now.get(k) !== v || !now.has(k)) changes.push([k, v, now.get(k)]);
      for (const [k, v] of now) if (!was.has(k)) changes.push([k, undefined, v]);
      entries.push({ kind: "map", slot, changes });
    } else entries.push({ kind: "set", slot, was, now });
  };
  for (const k of BOOK) add({ key: k }, (before as unknown as Record<string, unknown>)[k], (after as unknown as Record<string, unknown>)[k]);
  const bp = before.pages.map((x) => x.page), ap = after.pages.map((x) => x.page);
  const inA = new Set(ap), inB = new Set(bp);
  if (bp.length !== ap.length || bp.some((p, i) => p !== ap[i])) {
    const extra = new Map<Page, SavedPage>();
    for (const sp of before.pages) if (!inA.has(sp.page)) extra.set(sp.page, sp);
    for (const sp of after.pages) if (!inB.has(sp.page)) extra.set(sp.page, sp);
    entries.push({ kind: "pages", order: [bp, ap], extra });
  }
  for (const b of before.pages) {
    if (!inA.has(b.page)) continue;
    const a = pageOf(after, b.page), page = b.page;
    add({ page, key: "pname" }, b.name, a.name);
    for (const k of SHEET) add({ page, key: k }, (b.sheet as unknown as Record<string, unknown>)[k], (a.sheet as unknown as Record<string, unknown>)[k]);
    add({ page, key: "notes" }, b.notes, a.notes);
    const keys = new Set([...Object.keys(b.view), ...Object.keys(a.view)]);
    for (const k of keys) add({ page, key: k }, (b.view as unknown as Record<string, unknown>)[k], (a.view as unknown as Record<string, unknown>)[k]);
  }
  let cost = 0;
  for (const e of entries) {
    if (e.kind === "map") cost += e.changes!.length;
    else if (e.kind === "set") cost += sizeOf(e.was) + sizeOf(e.now);
    else for (const sp of e.extra!.values()) cost += sp.sheet.inputs.size + sp.sheet.formulas.size + sp.sheet.names.size + sp.sheet.tables.size + sp.notes.size;
  }
  return { label, entries, cost, active: before.active };
}

// The workbook `s` with `step` taken back (backwards) or done again.
export function applyStep(step: Step, s: Saved, backwards: boolean): Saved {
  let out: Saved = { ...s, active: step.active };
  step.active = s.active;
  for (const e of step.entries) {
    if (e.kind === "set") out = put(out, e.slot!, backwards ? e.was : e.now);
    else if (e.kind === "map") {
      const m = new Map(get(out, e.slot!) as Map<unknown, unknown>);
      for (const [k, was, now] of e.changes!) {
        const v = backwards ? was : now;
        if (v === undefined) m.delete(k); else m.set(k, v);
      }
      out = put(out, e.slot!, m);
    } else {
      const to = backwards ? e.order![0] : e.order![1];
      out = { ...out, pages: to.map((p) => out.pages.find((x) => x.page === p) ?? e.extra!.get(p)!) };
    }
  }
  return out;
}
