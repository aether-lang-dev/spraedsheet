// The Turns queue (sheet.ts) holds one change at a time. A turn whose job
// throws, or is interrupted (sae's page time limit killing a promise job:
// its await never resumes), fails that turn with an error and lets the queue
// move on; before, an interrupted job left every later turn waiting for good.
import { Turns } from "./sheet/sheet.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

// Microtasks given to a turn before it is called stuck.
const settle = async () => { for (let i = 0; i < 200; i++) await Promise.resolve(); };

{
  const turns = new Turns();
  const log: string[] = [];
  const bad = turns.take(async () => { throw new Error("boom"); }).then(() => "ran", (e) => `failed ${e.message}`);
  const good = turns.take(async () => { log.push("second"); return 2; });
  eq("a job that throws fails its turn, and the next turn runs", [await bad, await good, log], ["failed boom", 2, ["second"]]);
}

{
  const turns = new Turns();
  const log: string[] = [];
  const stuck = turns.take(() => new Promise<number>(() => {})).then(() => "ran", (e) => `failed ${e.message}`);
  const next = turns.take(async () => { log.push("next"); return 1; });
  await settle();
  eq("a turn waiting on a job that never comes back holds the queue", [turns.busy, log], [true, []]);
  let said = false;
  try { said = turns.interrupt("page time limit"); } catch (e) { said = false; }
  await settle();
  eq("interrupting it fails that turn with an error, says it did, and the queue moves on",
    [said, await stuck, log, turns.busy], [true, "failed interrupted: page time limit", ["next"], false]);
  eq("the turn behind it ran", await next, 1);
  eq("interrupting with no turn in progress does nothing", turns.interrupt("again"), false);
  const after = await turns.take(async () => "still works");
  eq("and the queue is still usable", after, "still works");
}

{
  // The interrupted job coming back late must not end a later turn.
  const turns = new Turns();
  let wake: (n: number) => void = () => {};
  const slow = turns.take(() => new Promise<number>((r) => { wake = r; })).then((v) => `got ${v}`, (e) => `failed ${e.message}`);
  await settle();
  turns.interrupt("x");
  let during = "";
  const next = turns.take(async () => { wake(9); await settle(); during = turns.busy ? "busy" : "idle"; return "next"; });
  eq("a late answer from an interrupted job changes nothing", [await slow, await next, during], ["failed interrupted: x", "next", "busy"]);
}

{
  // Through the commands, as sae's bus tells the app (bus.serve's stopped).
  const wb = new Workbook();
  const cmds = new Commands(wb);
  await cmds.run("cell.set cell=A1 input=1");
  eq("a workbook's turns are interrupted by an interrupted request",
    [wb.interruptTurns("page time limit"), (await cmds.run("cell.set cell=A2 input=2")).ok, wb.sheet.value("A2")], [false, true, 2]);
}

// expect: ok a job that throws fails its turn, and the next turn runs
// expect: ok a turn waiting on a job that never comes back holds the queue
// expect: ok interrupting it fails that turn with an error, says it did, and the queue moves on
// expect: ok the turn behind it ran
// expect: ok interrupting with no turn in progress does nothing
// expect: ok and the queue is still usable
// expect: ok a late answer from an interrupted job changes nothing
// expect: ok a workbook's turns are interrupted by an interrupted request
