// Record Macro, in REXX (GridCraft's Automate, as this app has it): every
// command a session runs while recording (the ones that change the
// workbook, from the panels, the keys, the command bar or a script) is a
// line of a RexxJS program, which runs them again. A cell set is SETCELL,
// as a person would write it; every other command goes to ADDRESS SHEET as
// its command line, so the program says exactly what was done and can be
// edited as REXX: a loop round a line, a variable in it. Pure.
import type { Addr } from "./address.ts";

// A command line ("cell.set cell=B2 input=5") as its id and parameters,
// a value quoted when it must be ("" a quote within).
export function parseLine(line: string): { id: string; params: Record<string, string> } {
  const m = /^\s*(\S+)/.exec(line);
  const params: Record<string, string> = {};
  let rest = line.slice(m ? m[0].length : 0);
  for (;;) {
    const p = /^\s*([A-Za-z_][\w]*)=("(?:[^"]|"")*"|\S*)/.exec(rest);
    if (!p) break;
    params[p[1]] = p[2].startsWith('"') ? p[2].slice(1, -1).replace(/""/g, '"') : p[2];
    rest = rest.slice(p[0].length);
  }
  return { id: m ? m[1] : "", params };
}

// A REXX string literal for a CALL's argument or an assignment, in double
// quotes as RexxJS reads one: a quote doubled, and the escapes a
// double-quoted literal has (\\, \n, \t, \r, \uXXXX) for a backslash and the
// control characters, which it cannot hold. (An ADDRESS command line has
// none: it is passed on as it is.)
export function literal(s: string): string {
  const body = s.replace(/\\/g, "\\\\").replace(/"/g, '""').replace(/[\u0000-\u001f]/g, (c) =>
    c === "\n" ? "\\n" : c === "\t" ? "\\t" : c === "\r" ? "\\r" : `\\u${c.charCodeAt(0).toString(16).padStart(4, "0")}`);
  return `"${body}"`;
}
// An ADDRESS command line: in double quotes, a quote doubled.
const quoted = (s: string): string => `"${s.replace(/"/g, '""')}"`;

// The recorded command lines as a program, headed by its name.
export function toRexx(name: string, lines: readonly string[], counted = lines.length): string {
  const body = lines.map((line) => {
    const { id, params } = parseLine(line);
    if (id === "cell.set" && Object.keys(params).length === 2) return `CALL SETCELL ${literal(params.cell)}, ${literal(params.input)}`;
    return `ADDRESS SHEET ${quoted(line)}`;
  });
  return [`/* ${name}: recorded by the spreadsheet, ${counted} command${counted === 1 ? "" : "s"} */`, ...body].join("\n");
}

// A macro's parameters as the REXX it starts with: "CELL=B2;N=5" is
// CELL = "B2" and N = "5", each a variable's name and its value.
export function parameters(args: string): string {
  return args.split(";").map((a) => a.trim()).filter((a) => a).map((a) => {
    const m = /^([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$/.exec(a);
    if (!m) throw new RangeError(`${a} is not a parameter (NAME=value)`);
    return `${m[1].toUpperCase()} = ${literal(m[2])}\n`;
  }).join("");
}

// A macro run when the workbook opens, or after a change to a range.
export interface Trigger {
  readonly macro: string;
  readonly on: "open" | "change";
  readonly range?: readonly [Addr, Addr];
  readonly sheet?: string;
}

// A macro's name: a REXX-ish word, as Excel's Record Macro asks for.
export const MACRO_NAME = /^[A-Za-z_][A-Za-z0-9_]{0,63}$/;
