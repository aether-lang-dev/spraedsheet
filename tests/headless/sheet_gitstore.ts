// Versions kept in a git repository (blog-entry-todo.md B.4,
// app/sheet/gitstore.ts and the version.* commands): two saves, the log
// lists both, the diff names the changed cells; the bytes are git's (a
// blob's id is git's, a commit names its tree and parent). A real git reads
// the same folder in tests/cli/test_versions.sh.
import { Commands } from "./sheet/commands.ts";
import type { Store } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { changeLine, diffText, head, save, textAt, versions } from "./sheet/gitstore.ts";
import type { RepoFs } from "./sheet/gitstore.ts";
import { sha1, hex } from "./sheet/crypto.ts";
import { utf8Bytes, utf8Text } from "./sheet/zip.ts";
import { eq } from "./lib/check.ts";

// A repository in memory: files by path.
const files = new Map<string, Uint8Array>();
const repo: RepoFs = { read: (p) => files.get(p) ?? null, write: (p, b) => { files.set(p, b); } };
const store: Store = { get: () => null, set: () => {}, remove: () => {}, repo };

const wb = new Workbook();
const cmds = new Commands(wb, store);
await cmds.run("formulas.use language=Excel");
await cmds.run("cell.set cell=A1 input=Item");
await cmds.run("cell.set cell=B1 input=2");
await cmds.run('cell.set cell=B2 input="=B1*21"');
const one = await cmds.execute("version.save", { message: "First" });
eq("a first save", [one.ok, /^saved [0-9a-f]{7}: First \(.* lines\)$/.test(one.message), versions(repo).length, files.has("HEAD"), files.has("refs/heads/main")], [true, true, 1, true, true]);
eq("nothing changed: no second version", [(await cmds.execute("version.save", { message: "Again" })).message.startsWith("nothing changed since"), versions(repo).length], [true, 1]);

await cmds.run("cell.set cell=B1 input=3");
await cmds.run("cell.set cell=C3 input=new");
await cmds.run("cell.comment cell=A1 text=heading");
const diffNow = await cmds.execute("version.diff");
eq("before it is kept: the diff of the workbook against the last version, cell by cell", (diffNow.output as { where: string; was: string; now: string }[]).map(changeLine),
  ["Sheet1!A1: Item → Item {\"comment\":\"heading\"}", "Sheet1!B1: 2 → 3", "Sheet1!C3: added new"]);
const two = await cmds.execute("version.save", { message: "Second\n\nwith a body" });
const log = (await cmds.execute("version.list")).output as { id: string; message: string; author: string }[];
eq("the log: newest first, with messages", [log.map((v) => v.message), log.length], [["Second\n\nwith a body", "First"], 2]);
const [v2, v1] = log;
eq("the diff of two versions names the changed cells", (await cmds.execute("version.diff", { a: v1.id.slice(0, 7), b: "HEAD" })).output as unknown[], [
  { where: "Sheet1!A1", was: "Item", now: "Item\t{\"comment\":\"heading\"}" }, { where: "Sheet1!B1", was: "2", now: "3" }, { where: "Sheet1!C3", was: "", now: "new" }]);
eq("HEAD~1 is the first", (await cmds.execute("version.diff", { a: "HEAD~1", b: "HEAD" })).message.includes("3 changes"), true);

// It is git's format: a blob's id is the SHA-1 of "blob <n>\0<text>", the branch holds the last commit.
const text = textAt(repo, v1.id), body = utf8Bytes(text);
const header = utf8Bytes(`blob ${body.length}\0`), whole = new Uint8Array(header.length + body.length);
whole.set(header); whole.set(body, header.length);
eq("the blob's object id is git's", [files.has(`objects/${hex(sha1(whole)).slice(0, 2)}/${hex(sha1(whole)).slice(2)}`), head(repo) === v2.id], [true, true]);
eq("a commit's parent is the one before", [utf8Text(files.get("refs/heads/main")!).trim() === v2.id, versions(repo)[0].parent === v1.id, versions(repo)[1].parent], [true, true, null]);

// Restore the first: the workbook as it was.
await cmds.execute("version.restore", { id: v1.id.slice(0, 7) });
eq("restored: B1 is 2 again, no C3, nothing different from the first", [wb.sheet.input("B1"), wb.sheet.input("C3"), (await cmds.execute("version.diff", { a: v1.id.slice(0, 7) })).message], ["2", "", "no changes"]);

// The difference, as a function of two texts.
eq("a text diff: a setting, a section's entry, a list's element", diffText("# rexxsheet text 1\n@formulas\t\"excel\"\n@sheet\tS\n@names\tX\t\"A1\"\n@rules[]\t{\"a\":1}\n", "# rexxsheet text 1\n@formulas\t\"rexxjs\"\n@sheet\tS\n@names\tX\t\"A2\"\n@rules[]\t{\"a\":2}\n").map(changeLine),
  ["@formulas: \"excel\" → \"rexxjs\"", "S @names X: \"A1\" → \"A2\"", "S @rules[]: added {\"a\":2}", "S @rules[]: removed {\"a\":1}"]);

// REXX.
const said: string[] = [];
await cmds.script('LET id = SAVEVERSION("Third")\nSAY LENGTH(id)\nLET v = VERSIONS()\nSAY v.0 SUBSTR(v.1, 9)\nLET r = SETCELL("B1", 9)\nLET d = DIFFVERSIONS("HEAD")\nSAY d.0 d.1', (l) => said.push(l));
eq("SAVEVERSION, VERSIONS and DIFFVERSIONS from REXX", said, ["7", "3 Third", "1 Sheet1!B1: 2 → 9"]);

// Without a repository (no fs grant): refused, with why.
const bare = new Commands(new Workbook(), { get: () => null, set: () => {}, remove: () => {} });
eq("no repository: refused", (await bare.execute("version.save", { message: "x" })).message, "no version repository: this page has no fs grant to keep one in");

// expect: ok a first save
// expect: ok nothing changed: no second version
// expect: ok before it is kept: the diff of the workbook against the last version, cell by cell
// expect: ok the log: newest first, with messages
// expect: ok the diff of two versions names the changed cells
// expect: ok HEAD~1 is the first
// expect: ok the blob's object id is git's
// expect: ok a commit's parent is the one before
// expect: ok restored: B1 is 2 again, no C3, nothing different from the first
// expect: ok a text diff: a setting, a section's entry, a list's element
// expect: ok SAVEVERSION, VERSIONS and DIFFVERSIONS from REXX
// expect: ok no repository: refused
