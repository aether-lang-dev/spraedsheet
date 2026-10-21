// An injectable clock (docs/comparisons/IRONCALC.md action 6): TODAY() and NOW()
// read the host's clock, so a test could not pin them. workbook.clock sets a
// fixed instant (not saved with the file), or with no at= gives the host's
// clock back; the volatile cells are worked out again at once.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
await cmds.run('cell.set cell=A1 input="=TODAY()"');
await cmds.run('cell.set cell=A2 input="=NOW()"');
await cmds.run('cell.set cell=A3 input="=DAY(A1)&""/""&MONTH(A1)&""/""&YEAR(A1)"');
const host = wb.sheet.value("A1");

const set = await cmds.execute("workbook.clock", { at: "2026-10-21 09:30" });
eq("a fixed clock", [set.ok, set.message], [true, "the clock is 2026-10-21 09:30:00"]);
eq("TODAY() is that day, worked out at once", [wb.sheet.value("A1"), wb.sheet.value("A3")], [46316, "21/10/2026"]);
eq("NOW() is that day and 9:30", Math.round((Number(wb.sheet.value("A2")) - 46316) * 24 * 60), 570);

await cmds.execute("workbook.clock", { at: "2030-01-02" });
eq("another instant: midnight, volatile cells again", [wb.sheet.value("A1"), wb.sheet.value("A2"), wb.sheet.value("A3")], [47485, 47485, "2/1/2030"]);
eq("with the same clock, the same value however often it is read", [(await cmds.execute("formulas.recalculate")).ok, wb.sheet.value("A2")], [true, 47485]);

// A serial number works too.
await cmds.execute("workbook.clock", { at: "45000.25" });
eq("a serial: 45000 and a quarter day", [wb.sheet.value("A1"), wb.sheet.value("A2")], [45000, 45000.25]);

// Not the file's: saved JSON and text carry nothing of it, and an opened file leaves it.
const file = JSON.stringify((await cmds.execute("file.export")).output);
eq("not saved with the file", file.includes("clock"), false);

// Reading it back.
const state = await cmds.execute("workbook.clockState");
eq("the clock as it stands", [state.output, state.message], [45000.25, "the clock is fixed at 2023-03-15 06:00:00"]);

// Refused.
eq("refused: not a date or time", (await cmds.execute("workbook.clock", { at: "yesterday" })).message, "yesterday is not a date or time (2026-10-21, 2026-10-21 09:30, or a serial)");
eq("refused: a time past the day", (await cmds.execute("workbook.clock", { at: "2026-10-21 25:00" })).message, "2026-10-21 25:00 is not a date or time (2026-10-21, 2026-10-21 09:30, or a serial)");

// REXX.
const said: string[] = [];
await cmds.script('LET m = SETCLOCK("2026-10-21 12:00")\nSAY m\nSAY GETCLOCK()\nLET r = SETCELL("B1", "=TODAY()")\nSAY GETCELL("B1")\nLET m = SETCLOCK()\nSAY m', (l) => said.push(l));
eq("SETCLOCK and GETCLOCK from REXX", said, ["the clock is 2026-10-21 12:00:00", "46316.5", "46316", "the clock is the host's"]);

// No at=: the host's clock again.
await cmds.execute("workbook.clock");
eq("the host's clock back: today's serial is the host's, not the fixed one", [wb.sheet.value("A1") === host, (await cmds.execute("workbook.clockState")).message], [true, "the clock is the host's"]);

// expect: ok a fixed clock
// expect: ok TODAY() is that day, worked out at once
// expect: ok NOW() is that day and 9:30
// expect: ok another instant: midnight, volatile cells again
// expect: ok with the same clock, the same value however often it is read
// expect: ok a serial: 45000 and a quarter day
// expect: ok not saved with the file
// expect: ok the clock as it stands
// expect: ok refused: not a date or time
// expect: ok refused: a time past the day
// expect: ok SETCLOCK and GETCLOCK from REXX
// expect: ok the host's clock back: today's serial is the host's, not the fixed one
