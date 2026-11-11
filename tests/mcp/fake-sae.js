#!/usr/bin/env node
// A stand-in for sae on its control channel, for the MCP server's launch
// tests: started as `fake-sae.js --app <dir> --control-stdio`, it logs how
// it was started to $FAKE_SAE_LOG, says a page serves, and answers each
// request frame on its stdin with fake-sae-bus.js's canned answer for that
// command (so every request is in $FAKE_BUS_LOG as on the bus), on its
// stdout, as frames. When its stdin ends it logs "stopped" and exits, as a
// headless sae does. With FAKE_SAE_DEAF=1 it says a page serves and then
// never reads its stdin, nor exits: a sae that has stopped reading.
'use strict';
const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const log = (x) => process.env.FAKE_SAE_LOG && fs.appendFileSync(process.env.FAKE_SAE_LOG, JSON.stringify(x) + '\n');
const appdata = process.env.SAE_APPDATA_DIR || '';
log({ args: process.argv.slice(2), headless: process.env.AETHER_UI_HEADLESS || '', appdata });

function frame(obj) {
  const payload = Buffer.from(JSON.stringify(obj), 'utf8');
  const head = Buffer.alloc(4);
  head.writeUInt32BE(payload.length, 0);
  process.stdout.write(Buffer.concat([head, payload]));
}

frame({ event: 'serving' });
if (process.env.FAKE_SAE_DEAF === '1') {
  setInterval(() => {}, 60000);
  return;
}
let buf = Buffer.alloc(0);
process.stdin.on('data', (d) => {
  buf = Buffer.concat([buf, d]);
  while (buf.length >= 4 && buf.length >= 4 + buf.readUInt32BE(0)) {
    const n = buf.readUInt32BE(0);
    const req = JSON.parse(buf.subarray(4, 4 + n).toString('utf8'));
    buf = buf.subarray(4 + n);
    const r = spawnSync(process.execPath, [path.join(__dirname, 'fake-sae-bus.js'), 'Spreadsheet', req.command], { encoding: 'utf8' });
    let answer;
    try {
      answer = JSON.parse(r.stdout);
    } catch (e) {
      answer = { success: false, result: null, errorCode: 1, error: (r.stderr || 'no answer').trim() };
    }
    frame({ id: req.id, ...answer });
  }
});
process.stdin.on('end', () => {
  log('stopped');
  process.exit(0);
});
