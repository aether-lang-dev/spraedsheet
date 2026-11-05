// A snapshot of a map taken from the changes made since the last one, not by
// copying the map (docs/mutation_journal.md, "Recording at the mutation").
//
// The sheet's cells and its notes are the maps that grow with the workbook. A
// command used to copy each one it changed (saved() after it ran) and then
// compare the copy with the one before, key by key. Now the container notes
// each key as it changes it (the mutation recorded as it is made), and the
// next snapshot is an Overlay: the snapshot before, plus those keys with what
// they hold now. Taking it costs what changed; so does the step of undo, read
// from the overlay's own keys (Overlay.changesFrom).
//
// An Overlay is a Map for every reader (instanceof Map, iteration, size, get,
// has), and is never changed after it is made. Overlays stack, one per
// snapshot; at DEPTH a snapshot is a plain copy again, so a lookup never walks
// more than DEPTH layers and an old layer is not kept for ever. That copy, once
// in DEPTH snapshots, is the whole cost left (copies.entries counts it).
export const DEPTH = 32;

// The entries copied whole into snapshots (a plain Map made from the live
// one), for the tests' measurement (sheet_historycost).
export const copies = { maps: 0, entries: 0 };

export class Overlay<K, V> extends Map<K, V> {
  readonly base: ReadonlyMap<K, V>;
  // The keys changed since base, each with what it holds now (undefined: none).
  readonly delta: ReadonlyMap<K, V | undefined>;
  readonly depth: number;
  readonly #size: number;

  constructor(base: ReadonlyMap<K, V>, delta: ReadonlyMap<K, V | undefined>) {
    super();
    this.base = base;
    this.delta = delta;
    this.depth = base instanceof Overlay ? base.depth + 1 : 1;
    let n = base.size;
    for (const [k, v] of delta) {
      const had = base.has(k);
      if (v === undefined) { if (had) n--; } else if (!had) n++;
    }
    this.#size = n;
  }

  override get size(): number { return this.#size; }
  override get(k: K): V | undefined { return this.delta.has(k) ? this.delta.get(k) : this.base.get(k); }
  override has(k: K): boolean { return this.delta.has(k) ? this.delta.get(k) !== undefined : this.base.has(k); }
  override *entries(): IterableIterator<[K, V]> {
    for (const [k, v] of this.base) if (!this.delta.has(k)) yield [k, v];
    for (const [k, v] of this.delta) if (v !== undefined) yield [k, v];
  }
  override [Symbol.iterator](): IterableIterator<[K, V]> { return this.entries(); }
  override *keys(): IterableIterator<K> { for (const [k] of this.entries()) yield k; }
  override *values(): IterableIterator<V> { for (const [, v] of this.entries()) yield v; }
  override forEach(f: (v: V, k: K, m: Map<K, V>) => void, self?: unknown): void {
    for (const [k, v] of this.entries()) f.call(self, v, k, this);
  }
  override set(): this { throw new Error("a snapshot is not changed"); }
  override delete(): boolean { throw new Error("a snapshot is not changed"); }
  override clear(): void { throw new Error("a snapshot is not changed"); }

  // The keys that differ between `older` (a snapshot this one was layered on,
  // however far down) and this one, each with what it was and what it is
  // here; null when `older` is not below this one (then only a comparison of
  // the two whole maps can say).
  changesFrom(older: ReadonlyMap<K, V>): [K, V | undefined, V | undefined][] | null {
    const keys = new Set<K>();
    let m: ReadonlyMap<K, V> = this;
    while (m !== older) {
      if (!(m instanceof Overlay)) return null;
      for (const k of m.delta.keys()) keys.add(k);
      m = m.base;
    }
    const out: [K, V | undefined, V | undefined][] = [];
    for (const k of keys) {
      const was = older.get(k), now = this.get(k);
      if (was !== now) out.push([k, was, now]);
    }
    return out;
  }
}

// The snapshot of `live` after `prev`, given the keys changed since `prev`
// was taken: `prev` itself when none did, an Overlay of them, or a copy (no
// snapshot before, or DEPTH reached).
export function snapshot<K, V>(live: ReadonlyMap<K, V>, prev: ReadonlyMap<K, V> | null, touched: ReadonlySet<K> | null): ReadonlyMap<K, V> {
  if (prev && touched) {
    if (!touched.size) return prev;
    if (!(prev instanceof Overlay) || prev.depth < DEPTH) {
      const delta = new Map<K, V | undefined>();
      for (const k of touched) delta.set(k, live.get(k));
      return new Overlay(prev, delta);
    }
  }
  copies.maps++;
  copies.entries += live.size;
  return new Map(live);
}
