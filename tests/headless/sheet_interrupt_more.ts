// What sae's page time limit leaves behind, beyond the one sheet's turn
// (sheet_interrupt.ts): a recalculation that was passing changes between
// sheets, a script the UI ran inside a group of undo, and the work waiting
// for the next change when it could be done at once (recoverInterrupted).
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { EXCEL } from "./sheet/excel.ts";
import type { Language } from "./sheet/language.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const settle = async () => { for (let i = 0; i < 200; i++) await Promise.resolve(); };

// Excel, but a formula whose text ends in "*10" waits for ever while `stuck`:
// the job the time limit killed.
let stuck = false;
const gated: Language = {
  ...EXCEL,
  compile: (...a) => {
    const f = EXCEL.compile(...a);
    return f.kind === "formula" && a[0].endsWith("*10")
      ? { ...f, run: (env) => (stuck ? new Promise<never>(() => {}) : f.run(env)) }
      : f;
  },
};

// Several sheets: the edit's change reaches Sheet2, whose recalculation is
// stopped, and Sheet3, which the workbook had not yet been to.
{
  const wb = new Workbook();
  const cmds = new Commands(wb);
  await wb.setLanguage(gated);
  const on = (sheet: string, k: string) => show(wb.page(sheet)!.sheet.value(k));
  await cmds.run("sheet.add name=Sheet2");
  await cmds.run("sheet.add name=Sheet3");
  await cmds.run("sheet.activate name=Sheet2");
  await cmds.run("cell.set cell=A1 input==Sheet1!A1+1");
  await cmds.run("cell.set cell=B1 input==A1*10");
  await cmds.run("sheet.activate name=Sheet3");
  await cmds.run("cell.set cell=A1 input==Sheet1!A1+100");
  await cmds.run("sheet.activate name=Sheet1");
  await cmds.run("cell.set cell=A1 input=1");
  eq("worked out across the sheets", [on("Sheet2", "A1"), on("Sheet2", "B1"), on("Sheet3", "A1")], ["2", "20", "101"]);

  stuck = true;
  const edit = cmds.run("cell.set cell=A1 input=5").then((o) => o.ok ? "done" : `failed ${o.message}`, (e) => `failed ${e.message}`);
  await settle();
  eq("the change waits in Sheet2, and Sheet3 has not been reached", [wb.interruptTurns("page time limit"), on("Sheet3", "A1")], [true, "101"]);
  await edit;
  stuck = false;
  eq("recovering at once works out the sheet it stopped in and the ones after it",
    [await wb.recoverInterrupted(), on("Sheet2", "A1"), on("Sheet2", "B1"), on("Sheet3", "A1")], [true, "6", "60", "105"]);
  eq("and nothing is left to recover", await wb.recoverInterrupted(), false);
}

// A script run from the UI, in a group of undo, stopped part way.
{
  const wb = new Workbook();
  const cmds = new Commands(wb);
  await wb.setLanguage(gated);
  stuck = true;
  const run = cmds.script('LET a = SETCELL("A1", "1")\nLET b = SETCELL("A2", "2")\nLET c = SETCELL("B1", "=A1*10")', () => {})
    .then((o) => o.message, (e) => `failed ${e.message}`);
  await settle();
  eq("the script is in a group", wb.grouping, true);
  cmds.interrupted();
  await run;
  stuck = false;
  eq("an interrupted script's group is ended", wb.grouping, false);
  await cmds.run("edit.undo");
  eq("its changes are one step of undo", [show(wb.sheet.value("A1")), show(wb.sheet.value("A2"))], ["", ""]);
}

// No work, nothing to recover.
{
  const wb = new Workbook();
  eq("with nothing interrupted there is nothing to recover", await wb.recoverInterrupted(), false);
}

// expect: ok worked out across the sheets
// expect: ok the change waits in Sheet2, and Sheet3 has not been reached
// expect: ok recovering at once works out the sheet it stopped in and the ones after it
// expect: ok and nothing is left to recover
// expect: ok the script is in a group
// expect: ok an interrupted script's group is ended
// expect: ok its changes are one step of undo
// expect: ok with nothing interrupted there is nothing to recover
