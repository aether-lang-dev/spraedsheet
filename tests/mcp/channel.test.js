// The MCP server driving the real app over sae's control channel, end to
// end: the server starts sae (`--headless`, so `sae --app app
// --control-stdio` with no window) at the first call and holds its stdin and
// stdout; each tool call is a frame there and back. While the app runs,
// nothing it does may leave a way in for anyone else: no courier file in its
// data folder, and no listening socket, TCP or Unix, in sae or in the
// server. When the client goes, the server goes, and sae with it.
//   node --test tests/mcp/channel.test.js     (./run-tests.sh mcp_channel)
// Needs node, lsof (macOS, Linux, the BSDs) and sae's target/build/bin/sae
// (SAE_DIR names another sae checkout).
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { spawn, spawnSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');
const readline = require('readline');

const ROOT = path.join(__dirname, '..', '..');
const SAE = process.env.SAE_DIR || path.join(ROOT, 'sae');
const SAE_BIN = path.join(SAE, 'target', 'build', 'bin', process.platform === 'win32' ? 'sae.exe' : 'sae');
const args = (name) => JSON.parse(fs.readFileSync(path.join(__dirname, 'args', name), 'utf8'));

// An MCP client over the server's stdio: call(method, params) is a promise
// for the reply to that request.
function client(env) {
  const server = spawn(process.execPath, [path.join(ROOT, 'mcp', 'spreadsheet-mcp.js'), '--headless'], { env: { ...process.env, ...env }, stdio: ['pipe', 'pipe', 'pipe'] });
  let stderr = '';
  server.stderr.on('data', (d) => (stderr += d));
  const waiting = new Map();
  let ids = 0;
  readline.createInterface({ input: server.stdout }).on('line', (line) => {
    const m = JSON.parse(line);
    const w = waiting.get(m.id);
    if (w) { waiting.delete(m.id); w(m); }
  });
  const call = (method, params) => new Promise((resolve) => {
    const id = ++ids;
    waiting.set(id, resolve);
    server.stdin.write(JSON.stringify({ jsonrpc: '2.0', id, method, params }) + '\n');
  });
  const tool = async (name, a) => {
    const m = await call('tools/call', { name, arguments: a });
    assert.ok(m.result, 'a result: ' + JSON.stringify(m));
    return m.result;
  };
  const exited = new Promise((r) => server.on('close', (code) => r(code)));
  return { server, call, tool, exited, stderr: () => stderr };
}

// The processes the server started (sae), by parent.
function children(pid) {
  const r = spawnSync('pgrep', ['-P', String(pid)], { encoding: 'utf8' });
  return r.stdout.split('\n').map((s) => s.trim()).filter((s) => s);
}

// What `pid` listens on, as lsof names it ('' for nothing): its listening
// TCP sockets, and its Unix sockets that have a name in the file system (one
// a client could connect to; the system's unnamed pairs, "->0x...", are not).
function listening(pid) {
  const tcp = spawnSync('lsof', ['-nP', '-a', '-p', pid, '-iTCP', '-sTCP:LISTEN', '-Fn'], { encoding: 'utf8' }).stdout;
  const unix = spawnSync('lsof', ['-nP', '-a', '-p', pid, '-U', '-Fn'], { encoding: 'utf8' }).stdout;
  return [...tcp.split('\n').filter((l) => l.startsWith('n')), ...unix.split('\n').filter((l) => l.startsWith('n/'))].join(' ');
}

// The files in `dir`, all the way down.
function files(dir) {
  const out = [];
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) out.push(...files(p).map((f) => path.join(e.name, f)));
    else out.push(e.name);
  }
  return out;
}

test('an AI client drives the app through the MCP server over the control channel: no courier file, nothing listening, and the app gone with the client', { skip: !fs.existsSync(SAE_BIN) && 'no ' + SAE_BIN }, async () => {
  assert.equal(spawnSync('lsof', ['-v']).error, undefined, 'lsof, to see what listens');
  const data = fs.mkdtempSync(path.join(os.tmpdir(), 'sheet-mcp-channel-'));
  const c = client({ SAE_BIN, SAE_APPDATA_DIR: data, SPREADSHEET_MCP_TOOLS: 'all' });
  let sae = '';
  try {
    const init = await c.call('initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'test', version: '0' } });
    assert.equal(init.result.serverInfo.name, 'spreadsheet');
    assert.deepEqual(children(c.server.pid), [], 'no app before the first call');

    const list = await c.tool('spreadsheet_list_commands', { name: 'SETCELL' });
    assert.equal(list.isError, false, JSON.stringify(list));
    assert.match(list.content[0].text, /^SETCELL\(cell, value\)/);
    sae = children(c.server.pid)[0];
    assert.ok(sae, 'the server started sae');
    const cmd = spawnSync('ps', ['-o', 'command=', '-p', sae], { encoding: 'utf8' }).stdout.trim();
    assert.match(cmd, /--app .*app --control-stdio$/, 'started on the channel: ' + cmd);

    const wrote = await c.tool('spreadsheet_write_cells', args('expenses.json'));
    assert.equal(wrote.isError, false, JSON.stringify(wrote));
    assert.deepEqual(wrote.structuredContent.range, 'A1:C5');
    const read = await c.tool('spreadsheet_read_range', { range: 'A1:C5' });
    assert.deepEqual(read.structuredContent.rows[1], ['2026-10-01', "Taxi to C:\\temp's office", 23.5], 'the backslash and the quote as sent');
    assert.deepEqual(read.structuredContent.formulas, { C5: '=SUM(C2:C4)' });
    assert.equal(read.structuredContent.rows[4][2], 75.5, 'the total worked out');

    const sorted = await c.tool('spreadsheet_call', args('sort.json'));
    assert.equal(sorted.isError, false, JSON.stringify(sorted));
    const after = await c.tool('spreadsheet_read_range', { range: 'B2:B4' });
    assert.deepEqual(after.structuredContent.rows.map((r) => r[0]), ['Lunch "client" meeting', "Taxi to C:\\temp's office", 'Train'], 'sorted by amount, dearest first');
    const program = await c.tool('spreadsheet_run_program', args('program.json'));
    assert.equal(program.isError, false, JSON.stringify(program));
    assert.deepEqual(program.structuredContent.said, ['total 75.5 over 20: 2']);
    const rexx = await c.tool('spreadsheet_run', { command: 'GETCELL("D2")' });
    assert.equal(rexx.isError, false, JSON.stringify(rexx));
    assert.match(JSON.stringify(rexx.structuredContent ?? rexx.content), /0\.5298/, 'the share the program wrote, 40 / 75.5');

    // The window's tools go over the channel too: a release sae has no UI
    // driver, and says so (the app's own answer, not the server's).
    const shot = await c.tool('spreadsheet_screenshot', {});
    assert.equal(shot.isError, true);
    assert.match(shot.content[0].text, /no UI driver/);

    // Nothing for anyone else to find while it runs.
    assert.deepEqual(files(data).filter((f) => path.basename(f).startsWith('.control-bus')), [], 'no courier file in the app\'s data folder');
    assert.equal(listening(sae), '', 'sae listens on nothing, TCP or Unix');
    assert.equal(listening(String(c.server.pid)), '', 'nor does the server');
    assert.match(c.stderr(), /sae: control channel: on/, 'sae\'s own notes on stderr, not in the channel');
  } finally {
    c.server.stdin.end();
  }
  assert.equal(await c.exited, 0, 'the server exits when the client goes');
  const alive = () => { try { process.kill(Number(sae), 0); return true; } catch (e) { return false; } };
  const deadline = Date.now() + 5000;
  while (alive() && Date.now() < deadline) await new Promise((r) => setTimeout(r, 100));
  assert.equal(alive(), false, 'and sae with it');
  assert.match(c.stderr(), /control channel: closed by the program that started sae/);
  fs.rmSync(data, { recursive: true, force: true });
});
