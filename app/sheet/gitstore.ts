// A git repository the workbook's versions are kept in (blog-entry-todo.md
// B.4): each save a commit of the workbook's text form (textform.ts) as the
// one file `workbook.sheet.txt`, so `git log`, `git diff` and `git show` read
// them as they read any repository. Loose objects only (zlib, SHA-1), written
// and read here; the folder is a bare repository, `git --git-dir=FOLDER`.
// What reaches the disk is a RepoFs: the page's `fs` within its grant, or a
// map in memory for the tests.
//
// What it does not do: read a packed object (a repository `git gc` has
// packed is refused with why), branches (it commits to `main`), merges.
import { sha1, hex } from "./crypto.ts";
import { deflate, inflate, utf8Bytes, utf8Text } from "./zip.ts";

export interface RepoFs {
  // A file's bytes, or null where there is none.
  read(path: string): Uint8Array | null;
  // A file written (its folders made), atomically.
  write(path: string, bytes: Uint8Array): void;
}

export const FILE = "workbook.sheet.txt";
const REF = "refs/heads/main";

export interface Version {
  readonly id: string;
  readonly parent: string | null;
  readonly message: string;
  readonly author: string;
  readonly time: number;
}

function adler32(b: Uint8Array): number {
  let a = 1, c = 0;
  for (let i = 0; i < b.length; i++) { a = (a + b[i]) % 65521; c = (c + a) % 65521; }
  return ((c << 16) | a) >>> 0;
}

// The zlib stream git's loose objects are: a header, deflate, an Adler-32.
function zlib(b: Uint8Array): Uint8Array {
  const body = deflate(b), sum = adler32(b), out = new Uint8Array(body.length + 6);
  out[0] = 0x78; out[1] = 0x9c;
  out.set(body, 2);
  out[out.length - 4] = sum >>> 24; out[out.length - 3] = (sum >>> 16) & 255; out[out.length - 2] = (sum >>> 8) & 255; out[out.length - 1] = sum & 255;
  return out;
}

function joined(...parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let at = 0;
  for (const p of parts) { out.set(p, at); at += p.length; }
  return out;
}

// An object's id and its bytes as git stores them (the header, then the content).
function objectOf(type: string, content: Uint8Array): { id: string; stored: Uint8Array } {
  const whole = joined(utf8Bytes(`${type} ${content.length}\0`), content);
  return { id: hex(sha1(whole)), stored: whole };
}

const pathOf = (id: string) => `objects/${id.slice(0, 2)}/${id.slice(2)}`;

function put(fs: RepoFs, type: string, content: Uint8Array): string {
  const { id, stored } = objectOf(type, content);
  if (!fs.read(pathOf(id))) fs.write(pathOf(id), zlib(stored));
  return id;
}

function get(fs: RepoFs, id: string): { type: string; content: Uint8Array } {
  const z = fs.read(pathOf(id));
  if (!z) throw new RangeError(`no loose object ${id.slice(0, 7)} (a packed repository is not read here)`);
  const whole = inflate(z.subarray(2));
  const nul = whole.indexOf(0);
  const [type] = utf8Text(whole.subarray(0, nul)).split(" ");
  return { type, content: whole.subarray(nul + 1) };
}

const text = (b: Uint8Array | null) => (b ? utf8Text(b).trim() : "");

// The commit the branch is at, or null in an empty repository.
export function head(fs: RepoFs): string | null {
  return text(fs.read(REF)) || null;
}

// A commit of this text; the repository made if it is not there yet. The id.
export function save(fs: RepoFs, workbook: string, message: string, author: string, time: number): string {
  if (!fs.read("HEAD")) {
    fs.write("HEAD", utf8Bytes("ref: refs/heads/main\n"));
    fs.write("config", utf8Bytes("[core]\n\trepositoryformatversion = 0\n\tfilemode = false\n\tbare = true\n"));
  }
  const blob = put(fs, "blob", utf8Bytes(workbook));
  const entry = joined(utf8Bytes(`100644 ${FILE}\0`), Uint8Array.from(blob.match(/../g)!.map((h) => parseInt(h, 16))));
  const tree = put(fs, "tree", entry);
  const parent = head(fs), who = `${author.replace(/[<>\n]/g, "")} <${author.replace(/[^A-Za-z0-9]/g, "").toLowerCase() || "sheet"}@spreadsheet> ${time} +0000`;
  const commit = `tree ${tree}\n${parent ? `parent ${parent}\n` : ""}author ${who}\ncommitter ${who}\n\n${message.replace(/\s+$/, "")}\n`;
  const id = put(fs, "commit", utf8Bytes(commit));
  fs.write(REF, utf8Bytes(`${id}\n`));
  return id;
}

function commitOf(fs: RepoFs, id: string): Version & { tree: string } {
  const { type, content } = get(fs, id);
  if (type !== "commit") throw new RangeError(`${id.slice(0, 7)} is a ${type}, not a version`);
  const [headers, ...body] = utf8Text(content).split("\n\n");
  const field = (name: string) => headers.split("\n").find((l) => l.startsWith(`${name} `))?.slice(name.length + 1) ?? "";
  const who = /^(.*) <[^>]*> (\d+) [+-]\d{4}$/.exec(field("author"));
  return { id, tree: field("tree"), parent: field("parent") || null, message: body.join("\n\n").replace(/\n$/, ""), author: who?.[1] ?? "", time: Number(who?.[2] ?? 0) };
}

// The versions, newest first.
export function versions(fs: RepoFs, limit = 1000): Version[] {
  const out: Version[] = [];
  for (let id = head(fs); id && out.length < limit;) {
    const { tree: _tree, ...v } = commitOf(fs, id);
    out.push(v);
    id = v.parent;
  }
  return out;
}

// A version named by an id or the start of one, "HEAD", or "HEAD~n".
export function resolve(fs: RepoFs, name: string): string {
  const n = name.trim();
  const m = /^HEAD(?:~(\d+))?$/i.exec(n);
  if (m) {
    const all = versions(fs, Number(m[1] ?? 0) + 1), v = all[Number(m[1] ?? 0)];
    if (!v) throw new RangeError(`no version ${n}`);
    return v.id;
  }
  const hits = versions(fs).filter((v) => /^[0-9a-f]{4,40}$/i.test(n) && v.id.startsWith(n.toLowerCase()));
  if (hits.length !== 1) throw new RangeError(hits.length ? `${n} names several versions` : `no version ${n}`);
  return hits[0].id;
}

// The workbook's text at a version.
export function textAt(fs: RepoFs, id: string): string {
  const { tree } = commitOf(fs, id);
  const { content } = get(fs, tree);
  // Entries: "<mode> <name>\0<20 bytes>"
  let at = 0;
  while (at < content.length) {
    const nul = content.indexOf(0, at);
    const [, name] = utf8Text(content.subarray(at, nul)).split(" ");
    const sha = hex(content.subarray(nul + 1, nul + 21));
    if (name === FILE) return utf8Text(get(fs, sha).content);
    at = nul + 21;
  }
  throw new RangeError(`version ${id.slice(0, 7)} has no ${FILE}`);
}

// ---- The difference of two texts, cell by cell.

export interface Change {
  readonly where: string;   // "Sheet1!B2", or "Sheet1 @names TWO", or "@formulas"
  readonly was: string;     // "" where it was not there
  readonly now: string;     // "" where it is no longer there
}

// A text's lines by what they are about: a cell (sheet and address) by the
// rest of its line; a setting by its name and entry; a list's element by
// the whole of it.
function lineMap(t: string): Map<string, string> {
  const m = new Map<string, string>();
  let sheet = "";
  for (const line of t.split("\n")) {
    if (line === "" || line.startsWith("#")) continue;
    const f = line.split("\t");
    if (f[0] === "@sheet") { sheet = f[1]; continue; }
    const head = sheet ? `${sheet} ` : "";
    if (f[0].startsWith("@")) {
      if (f[0].endsWith("[]")) m.set(`${head}${f[0]} ${f.slice(1).join("\t")}`, f.slice(1).join("\t"));
      else if (f.length === 3) m.set(`${head}${f[0]} ${f[1]}`, f[2]);
      else m.set(`${head}${f[0]}`, f[1] ?? "");
    } else m.set(`${sheet}!${f[0]}`, f.slice(1).join("\t"));
  }
  return m;
}

export function diffText(before: string, after: string): Change[] {
  const a = lineMap(before), b = lineMap(after), out: Change[] = [];
  for (const [k, v] of b) if (a.get(k) !== v) out.push({ where: k.replace(/ @(\w+)\[\] .*$/s, " @$1[]"), was: a.get(k) ?? "", now: v });
  for (const [k, v] of a) if (!b.has(k)) out.push({ where: k.replace(/ @(\w+)\[\] .*$/s, " @$1[]"), was: v, now: "" });
  return out;
}

// A change as a line: "Sheet1!B2: =B1*21 → =B1*22", "Sheet1!C3: added =SUM(B1:B2)".
export function changeLine(c: Change): string {
  const t = (s: string) => s.replace(/\t/g, " ");
  return c.was === "" ? `${c.where}: added ${t(c.now)}` : c.now === "" ? `${c.where}: removed ${t(c.was)}` : `${c.where}: ${t(c.was)} → ${t(c.now)}`;
}
