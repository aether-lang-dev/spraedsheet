#!/usr/bin/env node
// A stand-in for sae's sae-bus, for mcp/'s protocol tests: `fake-sae-bus.js
// <app> <command>` answers as sae-bus does (one line of JSON, or a reason on
// stderr and a non-zero exit), from canned answers chosen by the command.
// Each call's arguments are appended to $FAKE_BUS_LOG, one JSON array a line.
'use strict';
const fs = require('fs');

const [app, command] = process.argv.slice(2);
if (process.env.FAKE_BUS_LOG) fs.appendFileSync(process.env.FAKE_BUS_LOG, JSON.stringify(process.env.FAKE_BUS_LOG_ENV ? [app, command, process.env.SAE_APPDATA_DIR || ''] : [app, command]) + '\n');

// The review (HOLDREVIEW and the rest): held or not, and the changes since,
// in a file beside the log, since each call is a process of its own.
const path = require('path');
const REVIEW = process.env.FAKE_BUS_LOG ? process.env.FAKE_BUS_LOG + '.review' : null;
const review = () => (REVIEW && fs.existsSync(REVIEW) ? JSON.parse(fs.readFileSync(REVIEW, 'utf8')) : { held: false, changes: [] });
const keepReview = (r) => REVIEW && fs.writeFileSync(REVIEW, JSON.stringify(r));

const answer = (result) => process.stdout.write(JSON.stringify({ success: true, result, errorCode: 0, error: '' }) + '\n');
const refuse = (error) => process.stdout.write(JSON.stringify({ success: false, result: null, errorCode: 1, error }) + '\n');
const cell = (ref, value, expression = '') => ({ ref, value, expression, comment: '', format: '' });
const stem = (xs) => Object.fromEntries([[0, xs.length], ...xs.map((x, i) => [i + 1, x])]);
const param = (name, type, optional = 0, about = '', variadic = 0) => ({ name, type, optional, variadic, about });
const described = (name, signature, does, gives, params) => ({ name, signature, does, gives, params: stem(params) });

// What LISTCOMMANDS(1) gives, for a few of the app's functions.
const FUNCTIONS = [
  described('SORTRANGE', 'SORTRANGE(range, by, [ascending])', 'Sorts a range\'s rows by a column', '"OK"',
    [param('range', 'range'), param('by', 'column'), param('ascending', 'flag', 1, '1 or left out: ascending; 0: descending')]),
  described('GETCELL', 'GETCELL(cell)', 'Reads a cell', 'what it shows', [param('cell', 'cell')]),
  described('SETCELL', 'SETCELL(cell, value)', 'Sets a cell, as typing it does', 'the value set', [param('cell', 'cell'), param('value', 'text')]),
  described('INSERTROW', 'INSERTROW(row)', 'Inserts a row there', '"OK"', [param('row', 'row')]),
  described('SETCELLS', 'SETCELLS(range, values)', 'Sets a range\'s cells', '"OK"', [param('range', 'range'), param('values', 'stem')]),
  described('SETCELLEDITOR', 'SETCELLEDITOR(cell, type, [config])', 'Gives a cell an editor', 'the cell',
    [param('cell', 'cell'), param('type', 'text'), param('config', 'json', 1)]),
  described('SETCELLVALIDATION', 'SETCELLVALIDATION(range, type, [limits...])', 'Sets the values a range may take', '"OK"',
    [param('range', 'range'), param('type', 'text'), param('limits', 'text', 1, '', 1)]),
  described('GETSHEETNAMES', 'GETSHEETNAMES()', 'The sheets', 'a stem of names', []),
  described('WHERE', 'WHERE(query, condition)', 'Keeps the rows that meet a condition', 'a query', [param('query', 'query'), param('condition', 'text')]),
];

// BATCH_EXECUTE(BASE64_DECODE('<json, base64>'), 1): each call's canned result; a
// SORTRANGE by column Q (outside its range, as the app would find) fails,
// which stops it, as the app's does.
function batch(json) {
  const calls = JSON.parse(json);
  const results = [];
  for (const [index, c] of calls.entries()) {
    if (c.command === 'SORTRANGE' && c.args[1] === 'Q') return refuse(`command ${index + 1} (SORTRANGE): Q is not in A1:B2 (nothing it did was kept)`);
    const result = c.command === 'SETCELL' ? c.args[1] : c.command === 'GETCELL' ? 'shown' : c.command === 'GETSHEETNAMES' ? stem(['Sheet1', 'Costs']) : 'OK';
    results.push({ index, command: c.command, success: true, result });
  }
  const r = review();
  if (r.held) {
    for (const c of calls) if (c.command === 'SETCELL') r.changes.push('Sheet1!' + c.args[0] + ': added ' + c.args[1]);
    keepReview(r);
  }
  answer(JSON.stringify(results));
}

const BATCH = /^BATCH_EXECUTE\(BASE64_DECODE\('([A-Za-z0-9+/=]*)'\), 1\)$/;
const PROGRAM = /^RUN_PROGRAM\(BASE64_DECODE\('([A-Za-z0-9+/=]*)'\), "json"\)$/;

// RUN_PROGRAM(BASE64_DECODE('<json string, base64>'), "json"): the program's
// lines counted, its SAY lines said back, RESULT from a `LET RESULT =` line;
// a program with BOOM in it fails, as one calling an unknown routine does.
function program(json) {
  const text = JSON.parse(json);
  if (text.includes('BOOM')) return refuse("Subroutine 'BOOM' not found (nothing it did was kept)");
  const lines = text.split('\n').filter((l) => l.trim() !== '');
  const said = lines.filter((l) => l.trim().startsWith('SAY ')).map((l) => l.trim().slice(4).replace(/^"|"$/g, ''));
  const m = /LET RESULT = (.*)/.exec(text);
  answer({ said: stem(said), result: m ? m[1].trim() : '' });
}

if (process.env.FAKE_BUS_NO_APP) {
  process.stderr.write('sae-bus: no courier file at /nonexistent/.control-bus\n');
  process.exit(1);
} else if (process.env.FAKE_BUS_GARBLED) {
  process.stdout.write('not json\n');
} else if (PROGRAM.test(command)) {
  program(Buffer.from(PROGRAM.exec(command)[1], 'base64').toString('latin1'));
} else if (BATCH.test(command)) {
  // As the app's BASE64_DECODE (atob): the bytes as Latin-1.
  batch(Buffer.from(BATCH.exec(command)[1], 'base64').toString('latin1'));
} else if (command === 'GETCELLS("A1:B2")') {
  answer({ 0: 4, 1: cell('A1', '2'), 2: cell('B1', 'x'), 3: cell('A2', 5, 'A1+3'), 4: cell('B2', '') });
} else if (command === 'GETCELLS("C3:C3")') {
  answer({ 0: 1, 1: cell('C3', 'one') });
} else if (command === 'INSPECTWORKBOOK()') {
  answer(JSON.stringify({ name: 'Untitled', active: 'Sheet1', language: 'Excel', names: ['TOTAL=Sheet1!B1:B3'],
    sheets: [{ name: 'Sheet1', used: 'A1:C4', cells: 9, formulas: 2, hidden: 0, names: [], tables: ['Sales=A1:C4'], charts: [], pivots: [], drawings: [] },
      { name: 'Costs', used: '', cells: 0, formulas: 0, hidden: 0, names: [], tables: [], charts: [], pivots: [], drawings: [] }] }));
} else if (command === '@ui GET /screenshot?b64=1') {
  answer('iVBORw0KGgoAAAANSUhEUg==');
} else if (command === '@ui POST /window/key?combo=Shift%2BTab') {
  answer({ fired: true });
} else if (command === '@ui GET /widgets') {
  answer([{ id: 1, type: 'vstack' }]);
} else if (command.startsWith('@ui ')) {
  refuse('@ui: ' + command.slice(4) + ' is not a path the control bus passes to the UI driver');
} else if (command === 'HOLDREVIEW(1)' || command === 'HOLDREVIEW()') {
  const r = review();
  if (r.held && command === 'HOLDREVIEW(1)') refuse('a review is waiting: ' + r.changes.length + ' changes; keep them or take them back first');
  else { if (!r.held) keepReview({ held: true, changes: [] }); answer(r.held ? 'already held: ' + r.changes.length + ' changes so far' : 'holding the workbook as it is for review'); }
} else if (command === 'REVIEWDIFF()') {
  const r = review();
  if (!r.held) refuse('nothing is held for review: review.hold first'); else answer(stem(r.changes));
} else if (command === 'KEEPREVIEW()' || command === 'TAKEBACKREVIEW()') {
  const r = review();
  if (!r.held) refuse('nothing is held for review: review.hold first');
  else { keepReview({ held: false, changes: [] }); answer((command === 'KEEPREVIEW()' ? 'kept ' : 'took back ') + r.changes.length + ' changes'); }
} else if (command === 'SPREADSHEET_VERSION()') {
  answer('1.0');
} else if (command === 'LISTCOMMANDS(1)') {
  answer(stem(FUNCTIONS));
} else if (command === 'GETSHEETNAMES()') {
  answer({ 0: 2, 1: 'Sheet1', 2: 'Costs' });
} else if (command.startsWith('BOGUS')) {
  refuse('Function BOGUS is not available');
} else if (command.startsWith('SETCELL')) {
  answer('100');
} else {
  answer(null);
}
