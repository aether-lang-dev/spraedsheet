// Reference examples and boundary cases, per function (docs/comparisons/GNUMERIC.md
// action 1): the cases are lib/examples.ts (shared with the generated status page,
// docs/function_status.md); this works each out and holds the library to having
// them. The last cases count the functions covered against the library.
import { names } from "./sheet/functions.ts";
import { SPECIAL_FORMS } from "./sheet/evaluate.ts";
import { CASES, check } from "./lib/examples.ts";
import { eq } from "./lib/check.ts";

let bad: string[] = [], checked = 0;
for (const fn of Object.keys(CASES)) {
  if (!names().includes(fn) && !SPECIAL_FORMS.includes(fn)) bad.push(`${fn}: not in the library`);
  if (CASES[fn].length < 2) bad.push(`${fn}: needs an example and a boundary`);
  for (const [src, want] of CASES[fn]) {
    checked++;
    const { ok, got } = check(src, want);
    if (!ok) bad.push(`${src}: got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`);
  }
}
eq("every example and boundary comes out as its definition says", bad, []);
eq("the cases are many enough to mean something", checked >= 1200, true);
eq("every function in the library has an example and boundary cases", names().filter((n) => !(n in CASES)), []);
eq("and so does each special form (LET, LAMBDA, MAP ...)", SPECIAL_FORMS.filter((n) => !(n in CASES)), []);

// expect: ok every example and boundary comes out as its definition says
// expect: ok the cases are many enough to mean something
// expect: ok every function in the library has an example and boundary cases
// expect: ok and so does each special form (LET, LAMBDA, MAP ...)
