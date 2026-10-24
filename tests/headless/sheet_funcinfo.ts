// One function metadata contract (docs/comparisons/HYPERFORMULA.md action 4, GNUMERIC.md
// action 1): formulas.functionInfo says what the app knows of one function in one
// record (arity, kind, volatility, whether TRUE/FALSE are refused, the Spanish
// name, a declared deviation from Excel), and formulas.functionStatus gives it for
// all of them; the tables that used to be parallel lists are checked to agree.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { ES_FUNCTIONS } from "./sheet/locale.ts";
import { VOLATILE_FORMS, SPECIAL_FORMS } from "./sheet/evaluate.ts";
import { NO_LOGICALS_LIST } from "./sheet/functions.ts";
import { KNOWN } from "./lib/ironcalc_calc_known.ts";
import { DIFFERS } from "./sheet/funcinfo.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook(), cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
type Info = { name: string; kind: string; min: number | null; max: number | null; arity: string; volatile: boolean; refusesLogicals: boolean; es: string | null; deviation: string | null };
const info = async (name: string) => (await cmds.execute("formulas.functionInfo", { name })).output as Info;

const sum = await info("SUM");
eq("SUM: a function of one or more values", [sum.kind, sum.min, sum.max, sum.arity, sum.volatile, sum.refusesLogicals, sum.es], ["function", 1, null, "at least 1 value", false, false, "SUMA"]);
const rand = await info("rand");
eq("RAND: volatile, takes none, any case of name", [rand.name, rand.volatile, rand.arity, rand.es], ["RAND", true, "no values", "ALEATORIO"]);
eq("DATEDIF: refuses TRUE and FALSE", (await info("DATEDIF")).refusesLogicals, true);
eq("LET: a special form", [(await info("LET")).kind, (await info("IF")).kind], ["special form", "form"]);
eq("OFFSET: volatile", (await info("OFFSET")).volatile, true);
eq("a deviation declared", typeof (await info("EDATE")).deviation, "string");
eq("a name that is not a function", [(await cmds.execute("formulas.functionInfo", { name: "NOPE" })).message], ["NOPE is not a function"]);

// The tables agree.
const all = ((await cmds.execute("formulas.functionStatus")).output as Info[]);
const names = new Set(all.map((i) => i.name));
const functions = (await cmds.execute("formulas.functions")).output as string[];
eq("the status lists every function the language has", [all.length, [...names].every((n) => functions.includes(n))], [functions.length, true]);
eq("every Spanish name is for a function that exists", Object.keys(ES_FUNCTIONS).filter((n) => !names.has(n)), []);
eq("every volatile form is a function that exists", [...VOLATILE_FORMS].filter((n) => !names.has(n)), []);
eq("every function that refuses logicals exists", NO_LOGICALS_LIST.filter((n) => !names.has(n)), []);
eq("every special form is listed as one", SPECIAL_FORMS.filter((n) => all.find((i) => i.name === n)?.kind !== "special form"), []);
eq("every ordinary function has an arity", all.filter((i) => i.kind === "function" && (i.min === null || i.arity === "")).map((i) => i.name), []);

// The declared deviations are the ledger, recounted: a workbook's cells count for
// the functions its name gives (ARABIC_ROMAN for both, _issue_N dropped).
const counted: Record<string, number> = {};
for (const [file, cells] of Object.entries(KNOWN)) {
  const base = (file.split("/").pop() ?? "").replace(/\.xlsx$/, "").replace(/_issue_\d+/, "");
  const fns = names.has(base) ? [base] : base.split("_").filter((t) => names.has(t));
  for (const f of fns) counted[f] = (counted[f] ?? 0) + Object.keys(cells).length;
}
eq("the declared deviations are what the ledger of differing cells comes to", JSON.stringify(counted, Object.keys(counted).sort()), JSON.stringify(DIFFERS, Object.keys(DIFFERS).sort()));

// REXX.
const said: string[] = [];
await cmds.script('LET m = FUNCTIONINFO("SUM")\nSAY m', (l) => said.push(l));
eq("FUNCTIONINFO from REXX", said, ["SUM: at least 1 value; differs from Excel somewhere"]);

// expect: ok SUM: a function of one or more values
// expect: ok RAND: volatile, takes none, any case of name
// expect: ok DATEDIF: refuses TRUE and FALSE
// expect: ok LET: a special form
// expect: ok OFFSET: volatile
// expect: ok a deviation declared
// expect: ok a name that is not a function
// expect: ok the status lists every function the language has
// expect: ok every Spanish name is for a function that exists
// expect: ok every volatile form is a function that exists
// expect: ok every function that refuses logicals exists
// expect: ok every special form is listed as one
// expect: ok every ordinary function has an arity
// expect: ok the declared deviations are what the ledger of differing cells comes to
// expect: ok FUNCTIONINFO from REXX
