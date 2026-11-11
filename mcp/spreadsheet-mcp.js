#!/usr/bin/env node
/*!
 * spreadsheet-mcp | MIT License (../LICENSE-PROTOCOL-MIT)
 */
/**
 * An MCP server for the sae spreadsheet (../app/): an AI client (Claude, or
 * any MCP client) drives the spreadsheet, which this server starts itself
 * and talks to over sae's stdin and stdout (sae's control channel), or one
 * already running, through its control bus (--attach).
 *
 *   claude mcp add spreadsheet -- node /path/to/mcp/spreadsheet-mcp.js
 *
 * MCP over stdio: one JSON-RPC message a line in, one a line out. Each tool
 * call is one control-bus request. Which tools a client is offered is set
 * when the server starts, by --tools=json|rexx|all (or SPREADSHEET_MCP_TOOLS):
 *
 *   json (the default)  arguments as JSON, the server writing the app's
 *                       commands from them: nothing for a model to escape
 *     spreadsheet_write_cells    rows of values from a cell
 *     spreadsheet_call           one function, or several as one step,
 *                                its arguments checked before it is sent
 *     spreadsheet_read_range     a range's values and formulas, as rows
 *     spreadsheet_inspect        the workbook at a glance: its sheets (each
 *                                its used range, cells, names, tables,
 *                                charts, pivots, drawings), its names
 *     spreadsheet_list_commands  the functions, each with its parameters
 *   rexx                for those who would write the app's REXX themselves
 *     spreadsheet_run            a command, a string of REXX the app runs
 *     spreadsheet_read_range, spreadsheet_inspect, spreadsheet_list_commands
 *   all                 all of them
 *
 * How it reaches the app, set when it starts:
 *
 *   (the default)      it starts the app itself at the first call that needs
 *                      it, `sae --app app --control-stdio`, and holds sae's
 *                      stdin and stdout as the channel (./sae-channel.js):
 *                      nothing on the file system, nothing listening, no
 *                      key; only this server and that sae hold the channel.
 *                      The window stays open for the person when the server
 *                      goes; closed by the person, the next call opens it
 *                      again.
 *   --headless         the same with no window (AETHER_UI_HEADLESS=1), its
 *   (SPREADSHEET_MCP_HEADLESS=1)   data in a folder of its own (or
 *                      SAE_APPDATA_DIR); the app goes when the server does.
 *   --attach           an app the person started with its control bus on
 *   (SPREADSHEET_MCP_ATTACH=1)   (sae --app app --control-bus): a loopback
 *                      port and a key fresh for every launch, in a courier
 *                      file only this user can read, every message sealed
 *                      (sae's docs/control-bus.md). This server runs sae's
 *                      `sae-bus` for each call, as rexx/spreadsheet-address.js
 *                      does.
 *
 *   SAE_BIN            the sae to start (default ../sae/target/build/bin/sae)
 *   SAE_BUS            --attach: the sae-bus to run (default:
 *                      ../sae/target/sae-bus, beside this repository's sae link)
 *   SPREADSHEET_APP    --attach: the app's name, which finds its courier file
 *                      (default "Spreadsheet", app/app.json's name)
 *   SAE_APPDATA_DIR    the app's data folder: with --attach, where it was
 *                      started with one (sae-bus reads it); otherwise the
 *                      one the app is started with
 */
'use strict';

const { execFile } = require('child_process');
const path = require('path');
const readline = require('readline');

const PROTOCOL_VERSIONS = ['2025-06-18', '2025-03-26', '2024-11-05'];
const SERVER = { name: 'spreadsheet', version: '1.1.0' };
// sae-bus waits up to 10 s for the app's answer; a little more here, so its
// own reason is what comes back.
const BUS_DEADLINE_MS = 12000;
// The most cells one spreadsheet_write_cells writes: each is a command, and
// all of them must be done within the bus's 10 s.
const MOST_CELLS = 2000;
const SURFACES = ['json', 'rexx', 'all'];

const LIMITS =
  'Each call is one request to the running app and one step of its undo (the person takes it back with one Undo); ' +
  'one that fails is taken back whole. ' +
  'Request and reply only: you are not told when the person edits the sheet, so read again before relying on what you read earlier. ' +
  'One call at a time, each answered within 10 seconds.';

// ---- the tools -------------------------------------------------------------

const WRITE_CELLS = {
  name: 'spreadsheet_write_cells',
  title: 'Write rows of values',
  description:
    'Write a block of values into the active sheet: `rows` is a list of rows, each a list of values, written from `cell` ' +
    'rightwards and down. A value is text, a number, true/false, null (empties the cell), or a formula starting with "=" ' +
    '(e.g. "=SUM(B2:B9)"). Nothing needs escaping. At most ' + MOST_CELLS + ' cells a call. ' + LIMITS,
  inputSchema: {
    type: 'object',
    properties: {
      cell: { type: 'string', description: 'The top-left cell, e.g. "A1"' },
      rows: {
        type: 'array',
        description: 'e.g. [["Date", "Item", "Amount"], ["2026-10-01", "Taxi", 23.5]]',
        items: { type: 'array', items: { type: ['string', 'number', 'boolean', 'null'] } },
      },
    },
    required: ['cell', 'rows'],
    additionalProperties: false,
  },
};

const CALL = {
  name: 'spreadsheet_call',
  title: 'Call a spreadsheet function',
  description:
    'Call one of the spreadsheet\'s functions by name, its arguments as a JSON list: ' +
    '{"function": "SORTRANGE", "args": ["A1:C10", "B", 0]}. Or several in turn as one step of undo, all or nothing: ' +
    '{"calls": [{"function": "INSERTROW", "args": [2]}, {"function": "SETCELL", "args": ["A2", "=SUM(B3:B9)"]}]}. ' +
    'spreadsheet_list_commands lists the functions and what each takes; the arguments are checked against that before anything is ' +
    'sent. Common ones: GETCELL(cell), SETFORMAT(cell, format) with "currency:USD", "percent:0.0%", "number:0.00" or ' +
    '"bold;color:red"; SORTRANGE(range, by, [ascending]); SUM_RANGE(range); INSERTROW(row); ADDSHEET(name); ' +
    'SETACTIVESHEET(name); GETSHEETNAMES(); UNDO(). A list argument is a JSON list; a JSON one may be an object. ' + LIMITS,
  inputSchema: {
    type: 'object',
    properties: {
      function: { type: 'string', description: 'A function\'s name, e.g. "SORTRANGE"' },
      args: { type: 'array', description: 'Its arguments, in order', items: {} },
      calls: {
        type: 'array',
        description: 'Instead of function and args: several calls, done in turn as one step',
        items: {
          type: 'object',
          properties: { function: { type: 'string' }, args: { type: 'array', items: {} } },
          required: ['function'],
          additionalProperties: false,
        },
      },
    },
    additionalProperties: false,
  },
};

const RUN = {
  name: 'spreadsheet_run',
  title: 'Run a spreadsheet command (REXX)',
  description:
    'Run one command in the spreadsheet: a REXX expression calling its control functions, as text. ' +
    'Values go inside the command as REXX strings in double quotes, and a quote inside one is written twice ("say ""hi"""). ' +
    'RexxJS reads \\n in a string as a new line, \\t as a tab and \\r as a return, and a backslash before one of those ' +
    'cannot be written in a string at all ("C:\\temp" arrives with a tab): put text with backslashes in through ' +
    'spreadsheet_write_cells or spreadsheet_call (--tools=all), which escape nothing. Common ones:\n' +
    '  SETCELL("A1", "100")          a value, or a formula: SETCELL("A3", "=A1+A2")\n' +
    '  GETCELL("A3")                 what a cell shows\n' +
    '  GETEXPRESSION("A3")           its formula, without the =\n' +
    '  CLEARCELL("A1:B2")            empty a range\n' +
    '  SETFORMAT("D2", "currency:USD")   or "percent:0.0%", "number:0.00", "bold;color:red;align:center"\n' +
    '  SETCOMMENT("A1", "checked")\n' +
    '  INSERTROW(3) / DELETEROW(3) / INSERTCOLUMN("B") / DELETECOLUMN("B")\n' +
    '  ADDSHEET("Costs") / SETACTIVESHEET("Costs") / GETSHEETNAMES() / GETACTIVESHEET()\n' +
    '  SORTRANGE("A1:C10", "B", 0)   by column B, descending (1 or left out: ascending)\n' +
    '  RESULT(RANGE("A1:C9") |> WHERE("amount > 100"))   a query over a range\'s rows\n' +
    '  UNDO() / REDO()\n' +
    '  EXPORT()                      the whole workbook, as its file\'s JSON\n' +
    'Several can go in one command, separated by ";": they are one step of undo, and the answer is the FIRST one\'s ' +
    '(read with a call of its own).\n' +
    'spreadsheet_list_commands lists every function with its parameters. ' + LIMITS,
  inputSchema: {
    type: 'object',
    properties: { command: { type: 'string', description: 'The REXX to run, e.g. SETCELL("A1", "100")' } },
    required: ['command'],
    additionalProperties: false,
  },
};

const RUN_PROGRAM = {
  name: 'spreadsheet_run_program',
  title: 'Run a REXX program',
  description:
    'Run a whole REXX program in the spreadsheet, as one step of undo, all or nothing: lines, DO i = 1 TO n and DO WHILE cond loops ' +
    '(no DO FOREVER or LEAVE: end a loop by its condition), IF/THEN/ELSE, SELECT, your own routines (CALL name / name: ... RETURN), and SAY. It calls the ' +
    'same functions as spreadsheet_call, called as functions (not with CALL): LET r = SETCELL("A" || i, i * 10), v = GETCELL("A1"), ' +
    'LET r = SORTRANGE("A1:C9", "C", 0); spreadsheet_list_commands lists them. Cell names read the sheet (A1 + B1). Set RESULT to give a ' +
    'value back; what it SAYs comes back too. The program is sent as it is, so nothing in it needs escaping on the way; inside it, REXX rules ' +
    'hold: a quote in a string is written twice, and \\n, \\t and \\r in a string are a new line, a tab and a return (so put text such as ' +
    'C:\\temp into cells with spreadsheet_write_cells, not in a program). It runs in the app, next to the data: loop over hundreds of cells in one call ' +
    'rather than one call a cell. It is stopped if one stretch of work runs past 5 s, or the whole past 10 s, and then taken back whole. ' +
    LIMITS,
  inputSchema: {
    type: 'object',
    properties: { program: { type: 'string', description: 'The REXX program, lines separated by new lines' } },
    required: ['program'],
    additionalProperties: false,
  },
};

const READ_RANGE = {
  name: 'spreadsheet_read_range',
  title: 'Read a range',
  description:
    'Read a rectangle of the active sheet, e.g. "A1:C5": each row\'s values as shown, and the formulas behind ' +
    'any that have one. ' + LIMITS,
  inputSchema: {
    type: 'object',
    properties: { range: { type: 'string', description: 'A1-style, e.g. "A1:C5" (or one cell, "B2")' } },
    required: ['range'],
    additionalProperties: false,
  },
};

const INSPECT = {
  name: 'spreadsheet_inspect',
  title: 'Inspect the workbook',
  description:
    'The workbook at a glance, before reading or writing it: its name, the active sheet, the formula language, each ' +
    'sheet (its used range such as "A1:D20", how many cells and formulas, whether hidden, its own names, tables, ' +
    'charts, pivot tables and drawings) and the workbook\'s names. ' + LIMITS,
  inputSchema: { type: 'object', properties: {}, additionalProperties: false },
};

const listCommands = (via) => ({
  name: 'spreadsheet_list_commands',
  title: 'List the spreadsheet functions',
  description:
    'Every function ' + via + ' may call (about 130), each as its signature, what it does and what it gives: ' +
    'SORTRANGE(range, by, [ascending]): optional parameters in brackets. Give a name for one function in full, with what ' +
    'each parameter takes. ' + LIMITS,
  inputSchema: {
    type: 'object',
    properties: { name: { type: 'string', description: 'One function, e.g. "SETCELLVALIDATION"; left out: all of them' } },
    additionalProperties: false,
  },
});

// The window itself (7.5): what a person at it has, over the bus's ui channel
// (app.json capabilities.controlBus.ui, a sae-driver build with its driver
// on): look, press keys, click a widget.
const UI_NEEDS = 'Needs the app\'s ui channel (app.json capabilities.controlBus.ui, a sae-driver build run with AETHER_UI_TEST_PORT); refused without. ';
const SCREENSHOT = {
  name: 'spreadsheet_screenshot',
  title: 'Look at the window',
  description: 'A picture of the spreadsheet\'s window as it is: the grid, the panels, the drawings and charts floating over it. ' + UI_NEEDS + LIMITS,
  inputSchema: { type: 'object', properties: {}, additionalProperties: false },
};
const KEY = {
  name: 'spreadsheet_key',
  title: 'Press a key',
  description: 'A key pressed in the window, as a person at the keyboard presses it, to whatever has the focus (the grid: arrows, Tab, Return, Escape, Ctrl+Z, a letter opens the editor). ' +
    'combo: e.g. "Down", "Shift+Tab", "Ctrl+C", "F2", "7". ' + UI_NEEDS + LIMITS,
  inputSchema: {
    type: 'object',
    properties: { combo: { type: 'string', description: 'The key, with modifiers joined by +' } },
    required: ['combo'],
    additionalProperties: false,
  },
};
const UI = {
  name: 'spreadsheet_ui',
  title: 'Ask the window\'s driver',
  description: 'One request to the window\'s UI driver: GET /widgets (every widget with its text, geometry and state), GET /widget/<id>, GET /focus, ' +
    'GET /window/pick?x=&y= (what a press at a point reaches), POST /widget/<id>/click (or set_text?v=, focus, hover, toggle, set_value?v=), ' +
    'POST /window/key?combo=, POST /window/resize?w=&h=. Nothing else is passed. ' + UI_NEEDS + LIMITS,
  inputSchema: {
    type: 'object',
    properties: {
      method: { type: 'string', enum: ['GET', 'POST'] },
      path: { type: 'string', description: 'e.g. "/widgets", "/widget/12/click", "/window/key?combo=Down"' },
    },
    required: ['method', 'path'],
    additionalProperties: false,
  },
};

// One request, one reviewable diff (--review): what a request changed waits
// to be kept or taken back whole.
const REVIEW_ACTIONS = ['hold', 'diff', 'keep', 'take_back'];
const REVIEW = {
  name: 'spreadsheet_review',
  title: 'Review what a request changed',
  description:
    'With review on (the server started with --review), one request, one diff: before a request that writes, the app holds the workbook as it is; ' +
    'after it, the cells it changed (old and new value or formula) are shown, and they wait. action "diff" shows them again, ' +
    '"keep" keeps them, "take_back" takes every one back whole (not undoable), "hold" holds by hand what is done from now on. ' +
    'Until the changes are kept or taken back, another request that writes is refused. The diff also holds any edit the person makes meanwhile. ' + LIMITS,
  inputSchema: {
    type: 'object',
    properties: { action: { type: 'string', enum: REVIEW_ACTIONS, description: 'hold, diff, keep or take_back' } },
    required: ['action'],
    additionalProperties: false,
  },
};

function toolsFor(surface, review = false) {
  const tools = surface === 'rexx'
    ? [RUN, RUN_PROGRAM, READ_RANGE, INSPECT, listCommands('spreadsheet_run or a program'), SCREENSHOT, KEY, UI]
    : surface === 'all'
      ? [WRITE_CELLS, CALL, RUN_PROGRAM, READ_RANGE, INSPECT, listCommands('spreadsheet_call, a program, or spreadsheet_run'), RUN, SCREENSHOT, KEY, UI]
      : [WRITE_CELLS, CALL, RUN_PROGRAM, READ_RANGE, INSPECT, listCommands('spreadsheet_call or a program'), SCREENSHOT, KEY, UI];
  return review ? [...tools, REVIEW] : tools;
}

// --review, or SPREADSHEET_MCP_REVIEW=1.
const reviewOf = (argv, env) => argv.includes('--review') || env.SPREADSHEET_MCP_REVIEW === '1';

// --tools=X, or SPREADSHEET_MCP_TOOLS; json when neither.
function surfaceOf(argv, env) {
  const flag = argv.find((a) => a.startsWith('--tools='));
  const asked = (flag ? flag.slice('--tools='.length) : env.SPREADSHEET_MCP_TOOLS || 'json').trim().toLowerCase();
  if (!SURFACES.includes(asked)) throw new Error('--tools is json, rexx or all, not ' + JSON.stringify(asked));
  return asked;
}

// ---- the bus ---------------------------------------------------------------

// A string as a REXX literal: double quotes, any inside written twice.
function rexxString(s) {
  return '"' + String(s).replace(/"/g, '""') + '"';
}

// How this server reaches the app: one of these, made once in main.
//
//   launch (the default)   it starts the app itself (sae --app app
//                          --control-stdio) at the first call that needs it,
//                          and holds sae's stdin and stdout as the channel
//                          (./sae-channel.js): no courier file, no port, no
//                          key. --headless: with no window, its data in a
//                          folder of its own; the app goes when the server
//                          does. With a window the app stays open for the
//                          person when the server goes, and if the person
//                          closes it, the next call starts it again.
//   attach (--attach)      the app the person started with its control bus
//                          on (sae --app app --control-bus), through sae's
//                          sae-bus, a request a process.
const { SaeChannel } = require('./sae-channel');

function launcher({ headless }) {
  const fs = require('fs'), os = require('os');
  let app = null, data = '', ownData = false;
  return {
    async request(command) {
      if (!app || !app.running()) {
        if (headless && !data) {
          data = process.env.SAE_APPDATA_DIR || fs.mkdtempSync(path.join(os.tmpdir(), 'sheet-mcp-app-'));
          ownData = !process.env.SAE_APPDATA_DIR;
        }
        app = new SaeChannel({
          bin: process.env.SAE_BIN || path.join(__dirname, '..', 'sae', 'target', 'build', 'bin', 'sae'),
          app: path.join(__dirname, '..', 'app'),
          headless,
          env: data ? { SAE_APPDATA_DIR: data } : {},
        });
        try {
          await app.start();
        } catch (e) {
          return { ok: false, error: 'no spreadsheet to ask: ' + e.message };
        }
      }
      return app.request(command, BUS_DEADLINE_MS);
    },
    async close() {
      if (app) await app.close();
      if (ownData) fs.rmSync(data, { recursive: true, force: true });
      app = null;
    },
  };
}

// One request over the bus, through sae-bus.
function attacher() {
  return {
    request(command) {
      const saeBus = process.env.SAE_BUS || path.join(__dirname, '..', 'sae', 'target', 'sae-bus');
      const app = process.env.SPREADSHEET_APP || 'Spreadsheet';
      return new Promise((resolve) => {
        execFile(saeBus, [app, command], { encoding: 'utf8', timeout: BUS_DEADLINE_MS, maxBuffer: 16 * 1024 * 1024 }, (error, stdout, stderr) => {
          if (error) {
            const why = (stderr || '').trim() || (error.killed ? 'no answer within ' + BUS_DEADLINE_MS / 1000 + ' s' : error.message);
            resolve({ ok: false, error: 'no spreadsheet to ask: ' + why });
            return;
          }
          let answer;
          try {
            answer = JSON.parse(stdout);
          } catch (e) {
            resolve({ ok: false, error: 'sae-bus answered with something other than JSON: ' + stdout.slice(0, 200) });
            return;
          }
          if (answer.success) resolve({ ok: true, result: answer.result === undefined ? null : answer.result });
          else resolve({ ok: false, error: answer.error || 'the command failed (error ' + (answer.errorCode || 1) + ')' });
        });
      });
    },
    async close() {},
  };
}

// --attach, or SPREADSHEET_MCP_ATTACH=1; --headless, or SPREADSHEET_MCP_HEADLESS=1.
const attachOf = (argv, env) => argv.includes('--attach') || env.SPREADSHEET_MCP_ATTACH === '1';
const headlessOf = (argv, env) => argv.includes('--headless') || env.SPREADSHEET_MCP_HEADLESS === '1';

function transportOf(argv, env) {
  if (attachOf(argv, env)) {
    if (headlessOf(argv, env)) throw new Error('--attach drives an app already running; --headless starts one: not both');
    return attacher();
  }
  return launcher({ headless: headlessOf(argv, env) });
}

// A REXX stem ({0: count, 1: ..., n: ...}) as an array; anything else as is.
function unstem(v) {
  if (!isStem(v)) return v;
  const out = [];
  for (let i = 1; i <= Number(v['0']); i++) out.push(v[String(i)]);
  return out;
}

function isStem(v) {
  if (v === null || typeof v !== 'object' || Array.isArray(v) || !('0' in v)) return false;
  const n = Number(v['0']);
  if (!Number.isInteger(n) || n < 0) return false;
  for (let i = 1; i <= n; i++) if (!(String(i) in v)) return false;
  return true;
}

// Stems as arrays, all the way down.
function plain(v) {
  if (Array.isArray(v)) return v.map(plain);
  if (isStem(v)) return unstem(v).map(plain);
  if (v !== null && typeof v === 'object') return Object.fromEntries(Object.entries(v).map(([k, x]) => [k, plain(x)]));
  return v;
}

// ---- ranges and cells ------------------------------------------------------

const RANGE = /^\s*([A-Za-z]{1,3})([1-9][0-9]*)\s*(?::\s*([A-Za-z]{1,3})([1-9][0-9]*)\s*)?$/;
const REF = /^([A-Za-z]+)([0-9]+)$/;

function colNumber(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function colLetters(n) {
  let s = '';
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

// "A1:C5" as its corners, or null when it is not a range.
function parseRange(text) {
  const m = RANGE.exec(String(text));
  if (!m) return null;
  const c1 = colNumber(m[1]), r1 = Number(m[2]);
  const c2 = m[3] ? colNumber(m[3]) : c1, r2 = m[4] ? Number(m[4]) : r1;
  return { top: Math.min(r1, r2), bottom: Math.max(r1, r2), left: Math.min(c1, c2), right: Math.max(c1, c2) };
}

// GETCELLS's cells laid out as the range's rows, and the formulas by cell.
function asRows(span, cells) {
  const width = span.right - span.left + 1;
  const rows = [];
  for (let r = span.top; r <= span.bottom; r++) rows.push(new Array(width).fill(''));
  const formulas = {};
  for (const c of cells) {
    if (!c || typeof c !== 'object') continue;
    const m = REF.exec(String(c.ref));
    if (!m) continue;
    const r = Number(m[2]) - span.top, k = colNumber(m[1]) - span.left;
    if (r < 0 || r >= rows.length || k < 0 || k >= width) continue;
    rows[r][k] = c.value === undefined || c.value === null ? '' : c.value;
    if (c.expression) formulas[String(c.ref).toUpperCase()] = '=' + c.expression;
  }
  return { rows, formulas };
}

// ---- the functions, described by the app (LISTCOMMANDS(1)) -----------------

// LISTCOMMANDS(1)'s answer as the functions, each with its parameters, sorted.
function catalogueOf(result) {
  const listed = unstem(result);
  return (Array.isArray(listed) ? listed : [])
    .filter((c) => c && typeof c === 'object')
    .map((c) => ({ ...c, params: (unstem(c.params) || []).map((x) => ({ ...x, optional: x.optional === 1, variadic: x.variadic === 1 })) }))
    .sort((x, y) => String(x.name).localeCompare(String(y.name)));
}

const isNumberish = (v) => (typeof v === 'number' && Number.isFinite(v)) || (typeof v === 'string' && v.trim() !== '' && Number.isFinite(Number(v)));

// One argument as the app takes it, or { error }. `where` names it.
function argument(param, v, where) {
  const bad = (wanted) => ({ error: where + ' (' + param.name + ') must be ' + wanted + ', not ' + JSON.stringify(v) });
  switch (param.type) {
    case 'row':
      return isNumberish(v) && Number(v) >= 1 && Number.isInteger(Number(v)) ? { value: Number(v) } : bad('a row number from 1');
    case 'number':
      return isNumberish(v) ? { value: Number(v) } : bad('a number');
    case 'flag':
      if (v === true || v === 1 || v === '1') return { value: 1 };
      if (v === false || v === 0 || v === '0') return { value: 0 };
      return bad('1 or 0 (or true or false)');
    case 'column':
      if (typeof v === 'string' && /^[A-Za-z]{1,3}$/.test(v.trim())) return { value: v.trim().toUpperCase() };
      return isNumberish(v) && Number(v) >= 1 && Number.isInteger(Number(v)) ? { value: Number(v) } : bad('a column, "B" or 2');
    case 'stem':
      return Array.isArray(v) ? { value: v } : bad('a list');
    case 'json':
      return typeof v === 'string' ? { value: v } : v !== null && typeof v === 'object' ? { value: JSON.stringify(v) } : bad('JSON, as text or an object');
    case 'cell':
    case 'range':
    case 'text':
      return typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean' ? { value: String(v) } : bad('text');
    default:
      return { value: v };
  }
}

// A call checked against the function's description: { command, args } as
// BATCH_EXECUTE takes it, or { error }.
function checked(list, call, which) {
  const name = String(call && call.function || '').trim().toUpperCase();
  const where = which ? 'call ' + which + ', ' : '';
  if (!name) return { error: where + 'no function named; spreadsheet_list_commands lists them' };
  const fn = list.find((c) => c.name === name);
  if (!fn) return { error: where + 'no function ' + JSON.stringify(call.function) + '; spreadsheet_list_commands lists them' };
  if (fn.params.some((x) => x.type === 'query')) {
    return { error: where + name + ' works on a query (RANGE(...) |> ' + name + '(...)): run it with spreadsheet_run, with --tools=rexx or all' };
  }
  const args = call.args === undefined ? [] : call.args;
  if (!Array.isArray(args)) return { error: where + name + '\'s args must be a list' };
  const fixed = fn.params.filter((x) => !x.variadic);
  const rest = fn.params.find((x) => x.variadic);
  const needed = fixed.filter((x) => !x.optional).length;
  if (args.length < needed || (!rest && args.length > fixed.length)) {
    const takes = rest ? 'at least ' + needed : needed === fixed.length ? String(needed) : needed + ' to ' + fixed.length;
    return { error: where + fn.signature + ' takes ' + takes + ' argument' + (takes === '1' ? '' : 's') + ', not ' + args.length };
  }
  const out = [];
  for (let i = 0; i < args.length; i++) {
    const param = i < fixed.length ? fixed[i] : rest;
    const a = argument(param, args[i], where + name + '\'s argument ' + (i + 1));
    if (a.error) return { error: a.error };
    out.push(a.value);
  }
  return { command: name, args: out };
}

// The REXX that carries calls to the app with nothing in it to escape. A
// value's text never appears in a REXX string: RexxJS reads \n, \t and \r
// in its strings, in single quotes as in double, before it reads \\, so
// "C:\temp" arrives with a tab in it and no escaping gets it there whole
// (RexxJS's TODO.md). The calls go as JSON in base64,
// which RexxJS's BASE64_DECODE gives back; that reads bytes as Latin-1, so
// the JSON is first made ASCII (\uXXXX for the rest, which JSON.parse in
// the app turns back).
function batchCommand(calls) {
  return "BATCH_EXECUTE(" + carried(calls) + ", 1)";
}

// A whole program, the same way: RUN_PROGRAM(<its JSON in base64>, "json").
function programCommand(text) {
  return 'RUN_PROGRAM(' + carried(text) + ', "json")';
}

// A value as the app's BASE64_DECODE gives it back: ASCII JSON, in base64.
function carried(value) {
  const ascii = JSON.stringify(value).replace(/[\u007f-\uffff]/g, (c) => '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0'));
  return "BASE64_DECODE('" + Buffer.from(ascii, 'ascii').toString('base64') + "')";
}

// ---- answers ---------------------------------------------------------------

function text(value) {
  return typeof value === 'string' ? value : JSON.stringify(value, null, 2);
}

function ok(value, structured) {
  const r = { content: [{ type: 'text', text: text(value) }], isError: false };
  if (structured !== undefined) r.structuredContent = structured;
  return r;
}

function failed(message) {
  return { content: [{ type: 'text', text: message }], isError: true };
}

const WRITES = new Set(['spreadsheet_write_cells', 'spreadsheet_call', 'spreadsheet_run', 'spreadsheet_run_program']);

// One server: how it reaches the app, which tools it offers, where its
// replies go, and what it has learned of the app (its functions, asked for
// once: they do not change while it runs). main makes one; a test may make
// its own with a transport of its own. Nothing here is module state.
class Server {
  constructor({ transport, surface = 'json', reviewing = false, out = process.stdout }) {
    this.transport = transport;
    this.surface = surface;
    this.reviewing = reviewing;
    this.out = out;
    this.catalogue = null;
  }

  // One request to the app: { ok, result } or { ok: false, error }.
  bus(command) {
    return this.transport.request(command);
  }

  async functions() {
    if (this.catalogue) return { ok: true, list: this.catalogue };
    const a = await this.bus('LISTCOMMANDS(1)');
    if (!a.ok) return a;
    this.catalogue = catalogueOf(a.result);
    return { ok: true, list: this.catalogue };
  }

  // Calls as one request, BATCH_EXECUTE's, stopping at the first that fails
  // (so the request is taken back whole): each one's result, or { error }.
  async batch(calls) {
    const a = await this.bus(batchCommand(calls));
    if (!a.ok) return { error: a.error };
    let results;
    try {
      results = typeof a.result === 'string' ? JSON.parse(a.result) : a.result;
    } catch (e) {
      return { error: 'the app answered with something other than its results: ' + String(a.result).slice(0, 200) };
    }
    return { results: (Array.isArray(results) ? results : []).map((r) => plain(r.result === undefined ? null : r.result)) };
  }


  // What the app tells of a review, as a list of its changes.
  async changesOf() {
    const d = await this.bus('REVIEWDIFF()');
    if (!d.ok) return { error: d.error };
    const list = unstem(d.result);
    return { changes: (Array.isArray(list) ? list : []).map(String) };
  }

  async review(args) {
    if (!REVIEW_ACTIONS.includes(args.action)) return failed('spreadsheet_review: action is hold, diff, keep or take_back, not ' + JSON.stringify(args.action));
    if (args.action === 'diff') {
      const c = await this.changesOf();
      if (c.error) return failed(c.error);
      return ok(c.changes.length ? c.changes.join('\n') : 'no changes', { changes: c.changes });
    }
    const a = await this.bus({ hold: 'HOLDREVIEW()', keep: 'KEEPREVIEW()', take_back: 'TAKEBACKREVIEW()' }[args.action]);
    return a.ok ? ok(String(a.result), { result: a.result }) : failed(a.error);
  }

  // A tool that may write, with review on: held first (refused while another
  // request's changes wait), its diff read after; nothing changed, or a
  // failure (the app took the request back whole), leaves no hold.
  async reviewed(name, args, surface) {
    const h = await this.bus('HOLDREVIEW(1)');
    if (!h.ok) return failed(h.error);
    const r = await this.callTool(name, args, surface, false);
    const c = r.isError ? { changes: [] } : await this.changesOf();
    if (c.error || c.changes.length === 0) {
      await this.bus('KEEPREVIEW()');
      return c.error ? failed(c.error) : r;
    }
    const lines = c.changes.length + ' change' + (c.changes.length === 1 ? '' : 's');
    r.content[0].text += '\n\nheld for review: ' + lines + ':\n' + c.changes.join('\n') +
      '\nspreadsheet_review with keep or take_back ends the hold; no other request that writes is taken until then.';
    r.structuredContent = { ...(r.structuredContent || {}), review: { changes: c.changes } };
    return r;
  }

  async callTool(name, args, surface = 'all', reviewing = false) {
    args = args || {};
    if (!toolsFor(surface, reviewing).some((t) => t.name === name)) return null;
    if (name === 'spreadsheet_review') return this.review(args);
    if (reviewing && WRITES.has(name)) return this.reviewed(name, args, surface);
    if (name === 'spreadsheet_screenshot') {
      const a = await this.bus('@ui GET /screenshot?b64=1');
      if (!a.ok) return failed(a.error);
      return { content: [{ type: 'image', data: String(a.result), mimeType: 'image/png' }], isError: false };
    }
    if (name === 'spreadsheet_key') {
      if (typeof args.combo !== 'string' || args.combo.trim() === '') return failed('spreadsheet_key needs a combo, e.g. "Down" or "Ctrl+C"');
      const a = await this.bus('@ui POST /window/key?combo=' + encodeURIComponent(args.combo.trim()));
      if (!a.ok) return failed(a.error);
      return ok(a.result === null ? '(done)' : a.result, { result: a.result });
    }
    if (name === 'spreadsheet_ui') {
      if (!['GET', 'POST'].includes(args.method) || typeof args.path !== 'string' || !args.path.startsWith('/')) return failed('spreadsheet_ui needs a method (GET or POST) and a path starting with /');
      const a = await this.bus('@ui ' + args.method + ' ' + args.path);
      if (!a.ok) return failed(a.error);
      return ok(a.result === null ? '(done)' : a.result, { result: a.result });
    }
    if (name === 'spreadsheet_inspect') {
      const a = await this.bus('INSPECTWORKBOOK()');
      if (!a.ok) return failed(a.error);
      let w;
      try { w = JSON.parse(String(a.result)); } catch (e) { return failed('the app answered with something other than JSON'); }
      return ok(w, w);
    }
    if (name === 'spreadsheet_run') {
      if (typeof args.command !== 'string' || args.command.trim() === '') return failed('spreadsheet_run needs a command, e.g. SETCELL("A1", "100")');
      const a = await this.bus(args.command);
      if (!a.ok) return failed(a.error);
      const result = plain(a.result);
      return ok(result === null ? '(done; no result)' : result, { result });
    }
    if (name === 'spreadsheet_call') {
      const many = args.calls !== undefined;
      if (many === (args.function !== undefined)) return failed('spreadsheet_call takes a function (with its args) or calls, one of the two');
      if (many && args.args !== undefined) return failed('spreadsheet_call: args go inside each of the calls');
      const calls = many ? args.calls : [{ function: args.function, args: args.args }];
      if (!Array.isArray(calls) || calls.length === 0) return failed('spreadsheet_call: calls must be a list of at least one');
      const f = await this.functions();
      if (!f.ok) return failed(f.error);
      const sent = [];
      for (let i = 0; i < calls.length; i++) {
        const c = checked(f.list, calls[i], many ? i + 1 : 0);
        if (c.error) return failed(c.error + (many ? ' (nothing was sent)' : ''));
        sent.push(c);
      }
      const b = await this.batch(sent);
      if (b.error) return failed(b.error);
      if (!many) {
        const result = b.results[0] === undefined ? null : b.results[0];
        return ok(result === null || result === '' ? '(done; no result)' : result, { result });
      }
      return ok(b.results, { results: b.results });
    }
    if (name === 'spreadsheet_run_program') {
      if (typeof args.program !== 'string' || args.program.trim() === '') return failed('spreadsheet_run_program needs a program: lines of REXX');
      const a = await this.bus(programCommand(args.program));
      if (!a.ok) return failed(a.error);
      const got = plain(a.result) || {};
      const said = Array.isArray(got.said) ? got.said.map(String) : [];
      const result = got.result === undefined || got.result === null ? '' : got.result;
      const lines = said.slice();
      if (result !== '') lines.push('RESULT: ' + text(result));
      return ok(lines.length ? lines.join('\n') : '(ran; said nothing, no RESULT)', { said, result });
    }
    if (name === 'spreadsheet_write_cells') {
      const m = REF.exec(String(args.cell || '').trim());
      if (!m) return failed('spreadsheet_write_cells needs a cell to start from, such as "A1", not ' + JSON.stringify(args.cell));
      if (!Array.isArray(args.rows) || args.rows.some((r) => !Array.isArray(r))) return failed('spreadsheet_write_cells: rows must be a list of lists');
      const left = colNumber(m[1]), top = Number(m[2]);
      const sets = [];
      let width = 0;
      for (let r = 0; r < args.rows.length; r++) {
        width = Math.max(width, args.rows[r].length);
        for (let k = 0; k < args.rows[r].length; k++) {
          const v = args.rows[r][k];
          if (v !== null && !['string', 'number', 'boolean'].includes(typeof v)) {
            return failed('spreadsheet_write_cells: row ' + (r + 1) + ', value ' + (k + 1) + ' is ' + JSON.stringify(v) + ': text, a number, true/false or null');
          }
          sets.push({ command: 'SETCELL', args: [colLetters(left + k) + (top + r), v === null ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v)] });
        }
      }
      if (sets.length === 0) return failed('spreadsheet_write_cells: no values to write');
      if (sets.length > MOST_CELLS) return failed('spreadsheet_write_cells writes at most ' + MOST_CELLS + ' cells a call, not ' + sets.length + ': split the rows');
      const b = await this.batch(sets);
      if (b.error) return failed(b.error);
      const range = colLetters(left) + top + ':' + colLetters(left + width - 1) + (top + args.rows.length - 1);
      return ok('wrote ' + sets.length + ' cell' + (sets.length === 1 ? '' : 's') + ' in ' + range, { range, cells: sets.length });
    }
    if (name === 'spreadsheet_read_range') {
      const span = parseRange(args.range);
      if (!span) return failed('spreadsheet_read_range needs a range such as "A1:C5", not ' + JSON.stringify(args.range));
      const range = String(args.range).replace(/\s+/g, '').toUpperCase();
      const a = await this.bus('GETCELLS(' + rexxString(range.includes(':') ? range : range + ':' + range) + ')');
      if (!a.ok) return failed(a.error);
      const cells = unstem(a.result);
      const laid = asRows(span, Array.isArray(cells) ? cells : []);
      const structured = { range, rows: laid.rows, formulas: laid.formulas };
      return ok(structured, structured);
    }
    if (name === 'spreadsheet_list_commands') {
      const f = await this.functions();
      if (!f.ok) return failed(f.error);
      const commands = f.list;
      if (args.name !== undefined) {
        const one = commands.find((c) => c.name === String(args.name).trim().toUpperCase());
        if (!one) return failed('no function ' + JSON.stringify(args.name) + '; spreadsheet_list_commands with no name lists them');
        const lines = [one.signature, one.does + '; gives ' + one.gives + '.'];
        for (const x of one.params) lines.push('  ' + x.name + ': ' + x.type + (x.optional ? ', optional' : '') + (x.variadic ? ', any number' : '') + (x.about ? ' (' + x.about + ')' : ''));
        return ok(lines.join('\n'), one);
      }
      return ok(commands.map((c) => c.signature + ': ' + c.does + '; gives ' + c.gives).join('\n'), { commands });
    }
    return null;
  }

  // ---- JSON-RPC over stdio -------------------------------------------------

  send(message) {
    this.out.write(JSON.stringify(message) + '\n');
  }

  reply(id, result) {
    this.send({ jsonrpc: '2.0', id, result });
  }

  error(id, code, message) {
    this.send({ jsonrpc: '2.0', id, error: { code, message } });
  }

  async handle(msg) {
    const { surface, reviewing } = this;
    const isRequest = msg.id !== undefined && msg.id !== null;
    switch (msg.method) {
      case 'initialize': {
        const asked = msg.params && msg.params.protocolVersion;
        this.reply(msg.id, {
          protocolVersion: PROTOCOL_VERSIONS.includes(asked) ? asked : PROTOCOL_VERSIONS[0],
          capabilities: { tools: { listChanged: false } },
          serverInfo: SERVER,
          instructions:
            'Drives a sae spreadsheet: one this server starts at the first call (over its own stdin and stdout), or with --attach one already running with its control bus on. ' +
            'Read before you write, keep each call small, ' +
            'and tell the person what you changed: each call is one Undo for them.',
        });
        return;
      }
      case 'ping':
        if (isRequest) this.reply(msg.id, {});
        return;
      case 'tools/list':
        this.reply(msg.id, { tools: toolsFor(surface, reviewing) });
        return;
      case 'tools/call': {
        const p = msg.params || {};
        const result = await this.callTool(p.name, p.arguments, surface, reviewing);
        if (result === null) this.error(msg.id, -32602, 'no tool named ' + JSON.stringify(p.name));
        else this.reply(msg.id, result);
        return;
      }
      default:
        // Notifications (initialized, cancelled) need no answer; an unknown
        // request does.
        if (isRequest) this.error(msg.id, -32601, 'no method ' + JSON.stringify(msg.method));
    }
  }
}

if (require.main === module) {
  const argv = process.argv.slice(2);
  let server;
  try {
    server = new Server({ surface: surfaceOf(argv, process.env), reviewing: reviewOf(argv, process.env), transport: transportOf(argv, process.env) });
  } catch (e) {
    process.stderr.write('spreadsheet-mcp: ' + e.message + '\n');
    process.exit(2);
  }
  // Requests are answered in order: the app takes one command at a time.
  let queue = Promise.resolve();
  for (const sig of ['SIGINT', 'SIGTERM']) process.on(sig, () => server.transport.close().then(() => process.exit(0)));
  const lines = readline.createInterface({ input: process.stdin });
  lines.on('line', (line) => {
    if (line.trim() === '') return;
    let msg;
    try {
      msg = JSON.parse(line);
    } catch (e) {
      // In its turn, like any answer.
      queue = queue.then(() => server.error(null, -32700, 'not JSON: ' + e.message));
      return;
    }
    queue = queue.then(() => server.handle(msg)).catch((e) => {
      if (msg.id !== undefined && msg.id !== null) server.error(msg.id, -32603, String(e && e.message || e));
    });
  });
  lines.on('close', () => queue.then(() => server.transport.close()).then(() => process.exit(0)));
}

module.exports = { Server, reviewOf, toolsFor, surfaceOf, attachOf, headlessOf, transportOf, rexxString, batchCommand, programCommand, unstem, plain, parseRange, asRows, checked, catalogueOf };
