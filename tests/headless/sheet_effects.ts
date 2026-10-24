// Commands classified by effect (docs/comparisons/UNIVER.md action 1): each
// command in the registry says whether it is an edit (changes what the
// workbook holds), a view change (what is looked at, selected or copied),
// an external effect (reaches outside the workbook: files, versions,
// printing, links) or a query (reads and reports). The read view refuses
// exactly the edits; the catalogue (help, the command bar) carries it.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const READ = "read view: switch edit mode on to change the workbook";
const all = new Commands(new Workbook()).catalogue();
const effects = new Set(["edit", "view", "external", "query"]);
const of = new Map(all.map((c) => [c.id, (c as { effect?: string }).effect]));

eq("every command has an effect, one of four", all.filter((c) => !effects.has(of.get(c.id) ?? "")).map((c) => c.id), []);

const pick = (e: string) => all.filter((c) => of.get(c.id) === e).length;
eq("and each of the four is used", [pick("edit") > 100, pick("view") > 10, pick("external") > 10, pick("query") > 40], [true, true, true, true]);

const want: Record<string, string[]> = {
  edit: ["cell.set", "range.clear", "format.fill", "sheet.add", "edit.paste", "edit.undo", "table.define", "macro.set", "macro.trigger", "check.add", "version.restore", "review.takeback"],
  view: ["view.zoom", "view.freezeRows", "view.select", "view.editMode", "sheet.activate", "edit.copy", "edit.cut", "formulas.showFormulas"],
  external: ["file.open", "file.save", "file.print", "file.delete", "file.importCsv", "cell.followHyperlink", "version.save"],
  query: ["cell.get", "cell.info", "names.list", "comment.list", "chart.list", "macro.list", "version.list", "version.diff", "review.diff", "edit.find", "range.stats", "file.exportPdf"],
};
for (const [e, ids] of Object.entries(want)) eq(`${e}: ${ids.length} commands classified`, ids.map((id) => [id, of.get(id)]), ids.map((id) => [id, e]));

// The catalogue can be asked for one effect.
const queries = new Commands(new Workbook()).catalogue("", "query").map((c) => c.id);
eq("the catalogue by effect", [queries.includes("cell.get"), queries.includes("cell.set"), queries.length === pick("query")], [true, false, true]);

// The read view refuses exactly the edits: every command run in it with
// its required parameters filled in, each on a workbook of its own.
const refused: string[] = [], notEdit: string[] = [];
for (const c of all) {
  const cmds = new Commands(new Workbook());
  await cmds.run("cell.set cell=A1 input=1");
  await cmds.execute("view.editMode", { on: "0" });
  const required = c.params === "" ? [] : c.params.split(", ").filter((p) => !p.endsWith("?"));
  const params: Record<string, string> = {};
  for (const p of required) params[p] = "A1";
  let message = "";
  try { message = (await cmds.execute(c.id, params)).message; } catch (e) { message = ""; }
  const readViewRefused = message === READ;
  if (readViewRefused && of.get(c.id) !== "edit") refused.push(c.id);
  if (!readViewRefused && of.get(c.id) === "edit") notEdit.push(c.id);
}
eq("the read view refuses nothing but edits", refused, []);
eq("and refuses every edit", notEdit, []);

// expect: ok every command has an effect, one of four
// expect: ok and each of the four is used
// expect: ok edit: 12 commands classified
// expect: ok view: 8 commands classified
// expect: ok external: 7 commands classified
// expect: ok query: 12 commands classified
// expect: ok the catalogue by effect
// expect: ok the read view refuses nothing but edits
// expect: ok and refuses every edit
