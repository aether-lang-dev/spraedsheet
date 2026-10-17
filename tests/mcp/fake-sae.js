#!/usr/bin/env node
// A stand-in for sae, for the MCP server's --headless test: started as
// `fake-sae.js --app <dir> --control-bus`, it writes the courier file into
// $SAE_APPDATA_DIR, as sae does, logs how it was started to $FAKE_SAE_LOG,
// and waits to be stopped, logging that too.
'use strict';
const fs = require('fs');
const path = require('path');
const log = (x) => process.env.FAKE_SAE_LOG && fs.appendFileSync(process.env.FAKE_SAE_LOG, JSON.stringify(x) + '\n');
const appdata = process.env.SAE_APPDATA_DIR || '';
log({ args: process.argv.slice(2), headless: process.env.AETHER_UI_HEADLESS || '', appdata });
if (appdata) fs.writeFileSync(path.join(appdata, '.control-bus'), '{"url":"http://127.0.0.1:1","key":"x"}');
process.on('SIGTERM', () => { log('stopped'); process.exit(0); });
setInterval(() => {}, 1000);
