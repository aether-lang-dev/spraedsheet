#!/usr/bin/env node
// sheet: the spreadsheet from a shell, as GridCraft's CLI covers it:
// evaluate a formula, inspect a workbook, convert a file, run commands,
// list the commands and the functions, start the MCP server. Each run
// starts the app itself, headless (sae with AETHER_UI_HEADLESS=1, its
// control bus on, its data in a folder of its own), sends its requests over
// the bus with sae's sae-bus, and stops it; SHEET_ATTACH=1 uses an app
// already running instead. Needs node and sae beside this checkout.
//
//   sheet eval <formula> [--file F]       what the formula comes to (F opened first)
//   sheet inspect <file>                  its sheets, cells, names, tables and charts
//   sheet convert <in> <out>              .xlsx .xls .ods .csv .json in; .xlsx .csv .json .html .pdf out
//   sheet run <command line>... [--file F] [--save OUT]   commands in turn, what each gave
//   sheet commands [name]                 the control functions (a script's), or one described
//   sheet functions                       the formula functions
//   sheet mcp [--tools=...]               the MCP server, over stdio (mcp/spreadsheet-mcp.js)
//
// --json (after the subcommand) makes the result machine-readable: convert
// and run print one JSON object on stdout, {"ok":true, ...} (convert with
// "report", the list of what the file left out or the writer could not
// keep; run with "results", each {"line","said"}), and a failure is one JSON
// object on stderr, {"ok":false,"command","error","refused"[,"line"]}.
//
// A file goes to the app inside one bus request (an argument to sae-bus),
// so the system's limit on a command line bounds it: about a megabyte on
// macOS. Exit status: 0, 1 when a request was refused (said on stderr), 2
// when there was no app to ask.
'use strict';
const { execFile, spawn } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const SAE = process.env.SAE_DIR || path.join(ROOT, 'sae');
const SAE_BUS = process.env.SAE_BUS || path.join(SAE, 'target', 'sae-bus');
const SAE_BIN = process.env.SAE_BIN || path.join(SAE, 'target', 'build', 'bin', 'sae');
const APP = 'Spreadsheet';

class Refused extends Error {}

const rexxString = (s) => '"' + String(s).replace(/"/g, '""') + '"';
// Text the app reads back whole, whatever is in it: ASCII JSON in base64.
const carried = (value) => "BASE64_DECODE('" + Buffer.from(JSON.stringify(value).replace(/[\u007f-￿]/g, (c) => '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0')), 'ascii').toString('base64') + "')";

let appdata = process.env.SAE_APPDATA_DIR || '';
function bus(command) {
  return new Promise((resolve, reject) => {
    execFile(SAE_BUS, [APP, command], { encoding: 'utf8', timeout: 20000, maxBuffer: 64 * 1024 * 1024, env: { ...process.env, ...(appdata ? { SAE_APPDATA_DIR: appdata } : {}) } }, (error, stdout, stderr) => {
      if (error) return reject(new Error((stderr || '').trim() || error.message));
      let a;
      try { a = JSON.parse(stdout); } catch (e) { return reject(new Error('sae-bus answered with something other than JSON')); }
      if (a.success) resolve(a.result === undefined ? null : a.result);
      else reject(new Refused(a.error || 'refused'));
    });
  });
}
const unstem = (v) => (v && typeof v === 'object' && '0' in v ? Array.from({ length: Number(v['0']) }, (_, i) => v[String(i + 1)]) : v);

// The app, started headless with its control bus, until stop().
let child = null;
async function start() {
  if (process.env.SHEET_ATTACH === '1') return;
  // SHEET_APPDATA names the app's data folder (kept after the run: the version repository, versions.git, is in it); else one of its own.
  appdata = process.env.SHEET_APPDATA ? (fs.mkdirSync(process.env.SHEET_APPDATA, { recursive: true }), process.env.SHEET_APPDATA) : fs.mkdtempSync(path.join(os.tmpdir(), 'sheet-cli-'));
  child = spawn(SAE_BIN, ['--app', path.join(ROOT, 'app'), '--control-bus'], { env: { ...process.env, AETHER_UI_HEADLESS: '1', SAE_APPDATA_DIR: appdata }, stdio: 'ignore' });
  child.on('error', () => {});
  const courier = path.join(appdata, '.control-bus');
  for (let i = 0; i < 150; i++) {
    if (fs.existsSync(courier)) {
      try { await bus('SPREADSHEET_VERSION()'); return; } catch (e) { /* not answering yet */ }
    }
    await new Promise((r) => setTimeout(r, 100));
  }
  throw new Error('the app did not start (' + SAE_BIN + ')');
}
function stop() {
  if (child) child.kill();
  if (appdata && process.env.SHEET_ATTACH !== '1' && !process.env.SHEET_APPDATA) fs.rmSync(appdata, { recursive: true, force: true });
}

// A file into the workbook, by its kind.
async function open(file) {
  const ext = path.extname(file).toLowerCase(), bytes = fs.readFileSync(file), b64 = bytes.toString('base64'), name = path.basename(file, ext);
  if (ext === '.xlsx' || ext === '.xlsm') return bus(`OPENXLSX('${b64}', ${rexxString(name)})`);
  if (ext === '.xls') return bus(`OPENXLS('${b64}', ${rexxString(name)})`);
  if (ext === '.ods') return bus(`OPENODS('${b64}', ${rexxString(name)})`);
  if (ext === '.csv' || ext === '.tsv' || ext === '.txt') return bus(`IMPORTCSV("A1", '${b64}', "", "utf-8")`);
  if (ext === '.json') {
    // The file's own JSON (carried re-serialises it, ASCII), not a string of it.
    let data;
    try { data = JSON.parse(bytes.toString('utf8')); } catch (e) { throw new Refused(file + ': not JSON (' + e.message + ')'); }
    await bus(`IMPORT(${carried(data)})`);
    return `opened ${typeof data.name === 'string' && data.name ? data.name : name}`;
  }
  throw new Refused(file + ': not a file this opens (.xlsx, .xls, .ods, .csv, .json)');
}
// The workbook out, by the file's kind.
async function save(file) {
  const ext = path.extname(file).toLowerCase();
  if (ext === '.xlsx') return fs.writeFileSync(file, Buffer.from(String(await bus('SAVEXLSX()')), 'base64'));
  if (ext === '.pdf') return fs.writeFileSync(file, Buffer.from(String(await bus('EXPORTPDF()')), 'base64'));
  if (ext === '.csv') return fs.writeFileSync(file, String(await bus('EXPORTCSV()')));
  if (ext === '.html' || ext === '.htm') return fs.writeFileSync(file, String(await bus('EXPORTHTML("", 1)')));
  if (ext === '.json') { const j = await bus('EXPORT()'); return fs.writeFileSync(file, typeof j === 'string' ? j : JSON.stringify(j, null, 2)); }
  throw new Refused(file + ': not a file this writes (.xlsx, .csv, .json, .html, .pdf)');
}

const flag = (args, name) => { const i = args.indexOf(name); if (i < 0) return null; const v = args[i + 1]; args.splice(i, 2); return v; };

// What the last file read or written left out (FILEREPORT), as a list.
async function report() {
  const r = unstem(await bus('FILEREPORT()'));
  return Array.isArray(r) ? r.map(String) : [];
}

async function main(argv) {
  const [sub, ...args] = argv;
  const jsonAt = args.indexOf('--json'), json = jsonAt >= 0;
  if (json) args.splice(jsonAt, 1);
  let current = sub, line = null;
  if (sub === 'mcp') {
    const p = spawn(process.execPath, [path.join(ROOT, 'mcp', 'spreadsheet-mcp.js'), ...args], { stdio: 'inherit' });
    return new Promise((r) => p.on('exit', (c) => r(c ?? 0)));
  }
  const usage = 'sheet eval|inspect|convert|run|commands|functions|mcp ... (cli/sheet.js says how)';
  if (!['eval', 'inspect', 'convert', 'run', 'commands', 'functions'].includes(sub)) { process.stderr.write(usage + '\n'); return 2; }
  const file = flag(args, '--file'), out = flag(args, '--save');
  await start();
  try {
    if (file) await open(file);
    if (sub === 'eval') {
      if (!args.length) throw new Refused('sheet eval <formula>');
      const v = await bus(`EVALUATE(${rexxString(args.join(' '))})`);
      console.log(typeof v === 'string' ? v : JSON.stringify(v));
    } else if (sub === 'inspect') {
      if (!args.length) throw new Refused('sheet inspect <file>');
      console.log(await open(args[0]));
      const sheets = unstem(await bus('GETSHEETNAMES()'));
      for (const s of sheets) {
        await bus(`SETACTIVESHEET(${rexxString(s)})`);
        const last = await bus('GETLASTCELL()');
        console.log(`sheet ${s}: used to ${last}`);
      }
      const names = (await bus('GETALLNAMEDRANGES()')) || {};
      for (let i = 1; i <= Number(names['0'] || 0); i++) console.log(`name ${names[i + '.NAME']}: ${names[i + '.RANGE']}`);
      for (const t of unstem(await bus('LISTTABLES()')) || []) console.log(`table ${t}`);
      for (const c of unstem(await bus('LISTCHARTS()')) || []) console.log(`chart ${c}`);
    } else if (sub === 'convert') {
      if (args.length !== 2) throw new Refused('sheet convert <in> <out>');
      await open(args[0]);
      const read = await report();
      await save(args[1]);
      const wrote = path.extname(args[1]).toLowerCase() === '.xlsx' ? await report() : [];
      if (json) console.log(JSON.stringify({ ok: true, command: 'convert', in: args[0], out: args[1], read, write: wrote, report: [...read, ...wrote] }));
      else console.log(`${args[0]} -> ${args[1]}`);
    } else if (sub === 'run') {
      const results = [];
      for (line of args) {
        // RUNCOMMAND, the line carried as JSON: nothing in it read as REXX on
        // the way in, and one bus request per line, its answer what it gave.
        let said;
        try { said = await bus(`RUNCOMMAND(${carried(line)}, "json")`); } catch (e) {
          if (e instanceof Refused) throw new Refused(`${line}: ${e.message}`);
          throw e;
        }
        if (json) results.push({ line, said }); else console.log(`${line}: ${said}`);
      }
      line = null;
      if (json) console.log(JSON.stringify({ ok: true, command: 'run', results }));
    } else if (sub === 'commands') {
      const list = unstem(await bus('LISTCOMMANDS(1)')).filter((c) => c && typeof c === 'object');
      const one = args[0] ? list.find((c) => String(c.name).toUpperCase() === args[0].toUpperCase()) : null;
      if (args[0] && !one) throw new Refused(`no function ${args[0]}`);
      for (const c of one ? [one] : list.sort((a, b) => String(a.name).localeCompare(String(b.name)))) console.log(`${c.signature}${c.does ? ': ' + c.does : ''}`);
    } else if (sub === 'functions') {
      for (const f of unstem(await bus('LISTFUNCTIONS()'))) console.log(f);
    }
    if (out) { await save(out); console.log(`saved ${out}`); }
    return 0;
  } catch (e) {
    if (json) process.stderr.write(JSON.stringify({ ok: false, command: current, ...(line ? { line } : {}), error: e.message, refused: e instanceof Refused }) + '\n');
    else process.stderr.write(`sheet: ${e.message}\n`);
    return e instanceof Refused ? 1 : 2;
  } finally {
    stop();
  }
}

main(process.argv.slice(2)).then((code) => process.exit(code), (e) => { process.stderr.write(`sheet: ${e.message}\n`); stop(); process.exit(2); });
