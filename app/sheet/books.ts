// Several workbooks open at once (Excel's windows, as one sae page has
// them: one window, the workbooks switched in it as Excel's View > Switch
// Windows does): each its own Workbook and Commands, the same store of
// saved files; one active, the one the page shows and the commands reach.
// book.* commands (commands.ts hands them here): new, open (a saved file or
// a sample, into a new workbook), activate, close, list. A formula reads
// its own workbook only: no [Book2]Sheet1!A1 references across them.
import { Commands, memoryStore } from "./commands.ts";
import type { Outcome, Store } from "./commands.ts";
import { Workbook } from "./workbook.ts";

export interface Book {
  readonly wb: Workbook;
  readonly cmds: Commands;
}

export class Workbooks {
  readonly #books: Book[] = [];
  #active = 0;
  readonly #store: Store;
  // Told when another workbook is made active (the page shows it).
  onSwitch: (b: Book) => void = () => {};
  // Each new Workbook made ready as the page needs it (its URL opener).
  onMade: (wb: Workbook) => void = () => {};

  constructor(store?: Store) {
    this.#store = store ?? memoryStore();
    this.#make();
  }

  get active(): Book { return this.#books[this.#active]; }
  get all(): readonly Book[] { return this.#books; }

  #make(name?: string): Book {
    const wb = new Workbook();
    // The first keeps a workbook's own first name (Untitled); others Book2 ...
    if (name || this.#books.length) wb.name = name ?? this.#free();
    const cmds = new Commands(wb, this.#store);
    cmds.books = this;
    const b = { wb, cmds };
    this.onMade(wb);
    this.#books.push(b);
    return b;
  }

  // Book2, Book3 ... the first not open.
  #free(): string {
    for (let i = 2; ; i++) if (!this.#books.some((b) => b.wb.name.toLowerCase() === `book${i}`)) return `Book${i}`;
  }

  #find(name: string): number {
    const i = this.#books.findIndex((b) => b.wb.name.toLowerCase() === name.trim().toLowerCase());
    if (i < 0) throw new RangeError(`no workbook ${name} is open (${this.#books.map((b) => b.wb.name).join(", ")})`);
    return i;
  }

  #switch(i: number): void {
    this.#active = i;
    this.active.wb.claimDates();
    this.onSwitch(this.active);
  }

  async run(id: string, p: Record<string, string>): Promise<Outcome> {
    const names = () => this.#books.map((b) => b.wb.name);
    try {
      switch (id) {
        case "book.new": {
          const name = p.name?.trim();
          if (name && this.#books.some((b) => b.wb.name.toLowerCase() === name.toLowerCase())) throw new RangeError(`${name} is open already`);
          this.#make(name || undefined);
          this.#switch(this.#books.length - 1);
          return { ok: true, message: `${this.active.wb.name} opened (${names().length} workbooks)`, changed: this.active.wb.sheet.keys(), output: this.active.wb.name };
        }
        case "book.open": {
          if (p.name === undefined && p.sample === undefined) throw new RangeError("book.open needs a name (a saved file) or a sample");
          const b = this.#make();
          const o = await b.cmds.execute("file.open", { ...(p.name !== undefined ? { name: p.name } : {}), ...(p.sample !== undefined ? { sample: p.sample } : {}) });
          if (!o.ok) {
            this.#books.pop();
            return o;
          }
          let n = 2;
          const base = b.wb.name;
          while (this.#books.some((x) => x !== b && x.wb.name.toLowerCase() === b.wb.name.toLowerCase())) b.wb.name = `${base} (${n++})`;
          this.#switch(this.#books.length - 1);
          return { ok: true, message: `${b.wb.name}: ${o.message}`, changed: b.wb.sheet.keys(), output: b.wb.name };
        }
        case "book.activate": {
          this.#switch(this.#find(p.name ?? ""));
          return { ok: true, message: `${this.active.wb.name} active`, changed: this.active.wb.sheet.keys(), output: this.active.wb.name };
        }
        case "book.close": {
          const i = p.name === undefined ? this.#active : this.#find(p.name);
          if (this.#books.length === 1) throw new RangeError("the last workbook stays open (file.new empties it)");
          const [gone] = this.#books.splice(i, 1);
          const next = Math.min(this.#active > i ? this.#active - 1 : this.#active, this.#books.length - 1);
          this.#switch(next);
          return { ok: true, message: `${gone.wb.name} closed; ${this.active.wb.name} active`, changed: this.active.wb.sheet.keys(), output: this.active.wb.name };
        }
        case "book.list":
          return { ok: true, message: names().map((n, i) => (i === this.#active ? `${n} (active)` : n)).join(", "), changed: [], output: names() };
      }
      throw new RangeError(`no command ${id}`);
    } catch (e) {
      if (e instanceof RangeError) return { ok: false, message: e.message, changed: [] };
      throw e;
    }
  }
}
