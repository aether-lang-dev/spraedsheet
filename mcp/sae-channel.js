/*!
 * sae-channel | MIT License (../LICENSE-PROTOCOL-MIT)
 */
/**
 * The app, started by this program and driven over its own stdin and stdout:
 * sae's control channel (sae's docs/app-backend.md, "The control channel").
 *
 *   const app = new SaeChannel({ bin, app: 'app', headless: true });
 *   await app.start();                       // resolves once a page serves
 *   const a = await app.request('SPREADSHEET_VERSION()');   // { ok, result } or { ok: false, error }
 *   await app.close();
 *
 * `sae --app <dir> --control-stdio` takes its stdin and stdout for the
 * channel before anything else, so only this program and that sae hold it:
 * no courier file, no port, no key. Its stderr is this program's (the app's
 * print lines and sae's own notes go there). Frames are std.ipc's, a 4-byte
 * big-endian length and then one JSON object:
 *
 *   to sae     {"id": n, "command": "<string>"}
 *   from sae   {"id": n, "success": bool, "result": <JSON>, "errorCode": n, "error": "<why>"}
 *              {"event": "serving"}       a page serves: commands reach it from now
 *
 * The app's answers are the control bus's (sae-bus prints the same object),
 * so a request means the same on either.
 *
 * Closing the channel ends a sae with no window (headless); a sae with one
 * stays open for the person, and only the channel goes. Each SaeChannel is
 * its own sae; nothing here is shared.
 */
'use strict';

const { spawn } = require('child_process');

// sae answers a request within its own 10 s (the bus's deadline); a little
// more here, so its reason is what comes back.
const ANSWER_MS = 12000;
// From start to a page serving: a cold start of a large app, or a slow host.
const START_MS = 60000;
// The largest answer read (a screenshot, base64).
const MOST_BYTES = 64 * 1024 * 1024;
// A request handed to sae's stdin must be taken within this: sae reads its
// channel every 20 ms whatever its page is doing, so one it has not taken
// for this long has stopped reading (sae holds the same bound the other
// way, its frames.WRITE_MS).
const WRITE_MS = 5000;

class SaeChannel {
  // bin: the sae to start; app: the app's folder; headless: no window
  // (AETHER_UI_HEADLESS=1); env: more of sae's environment; args: more
  // arguments after the app's folder.
  constructor({ bin, app, headless = false, env = {}, stderr = 'inherit' }) {
    this.bin = bin;
    this.app = app;
    this.headless = headless;
    this.env = env;
    this.stderr = stderr;
    this.child = null;
    this.ids = 0;
    this.waiting = new Map();   // id -> { resolve, timer }
    this.buf = Buffer.alloc(0);
    this.serving = false;
    this.onServing = null;
    this.why = '';              // why it is not running ('' while it is, or before start)
    this.exited = null;         // resolves with the exit status
    this.stall = null;          // the timer on bytes sae has not taken, while there are some
  }

  // Start sae; resolves once a page serves, rejects with why it did not.
  start() {
    return new Promise((resolve, reject) => {
      const env = { ...process.env, ...this.env };
      if (this.headless) env.AETHER_UI_HEADLESS = '1';
      let child;
      try {
        child = spawn(this.bin, ['--app', this.app, '--control-stdio'], { env, stdio: ['pipe', 'pipe', this.stderr] });
      } catch (e) {
        this.why = 'the app did not start: ' + e.message;
        reject(new Error(this.why));
        return;
      }
      this.child = child;
      let settled = false;
      const done = (err) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        this.onServing = null;
        if (err) reject(err); else resolve();
      };
      const timer = setTimeout(() => {
        this.why = 'the app did not serve within ' + START_MS / 1000 + ' s';
        done(new Error(this.why));
        this.kill();
      }, START_MS);
      this.onServing = () => done(null);
      this.exited = new Promise((r) => child.on('close', (code, signal) => r(code === null ? signal : code)));
      child.on('error', (e) => {
        this.gone('the app did not start (' + this.bin + '): ' + e.message);
        done(new Error(this.why));
      });
      child.on('close', (code, signal) => {
        this.drained();
        this.gone('the app has exited (' + (code === null ? 'signal ' + signal : 'status ' + code) + ')');
        done(new Error(this.why));
      });
      // Its stdin closing early is the same as its exit, said there.
      child.stdin.on('error', () => {});
      child.stdin.on('drain', () => this.drained());
      child.stdout.on('data', (d) => this.read(d));
    });
  }

  // True while sae runs and the channel is open.
  running() {
    return this.child !== null && this.why === '';
  }

  // Frames out of what arrived.
  read(data) {
    this.buf = Buffer.concat([this.buf, data]);
    while (this.buf.length >= 4) {
      const n = this.buf.readUInt32BE(0);
      if (n > MOST_BYTES) {
        this.broke('a frame of ' + n + ' bytes, more than ' + MOST_BYTES);
        return;
      }
      if (this.buf.length < 4 + n) return;
      const text = this.buf.subarray(4, 4 + n).toString('utf8');
      this.buf = this.buf.subarray(4 + n);
      let m;
      try {
        m = JSON.parse(text);
      } catch (e) {
        this.broke('a frame that is not JSON: ' + text.slice(0, 80));
        return;
      }
      if (m && typeof m.event === 'string') {
        if (m.event === 'serving') {
          this.serving = true;
          if (this.onServing) this.onServing();
        }
        continue;
      }
      const w = m && this.waiting.get(m.id);
      if (!w) continue;   // answered late, after its own deadline here
      this.waiting.delete(m.id);
      clearTimeout(w.timer);
      if (m.success) w.resolve({ ok: true, result: m.result === undefined ? null : m.result });
      else w.resolve({ ok: false, error: m.error || 'the command failed (error ' + (m.errorCode || 1) + ')' });
    }
  }

  // The channel is no use: sae is stopped.
  broke(why) {
    this.gone('the app broke the channel: ' + why);
    this.kill();
  }

  // Every request waiting fails with why, and so does every later one.
  gone(why) {
    if (this.why === '') this.why = why;
    for (const [, w] of this.waiting) {
      clearTimeout(w.timer);
      w.resolve({ ok: false, error: this.why });
    }
    this.waiting.clear();
  }

  // One command: { ok: true, result } or { ok: false, error }.
  request(command, timeoutMs = ANSWER_MS) {
    if (!this.running()) return Promise.resolve({ ok: false, error: this.why || 'the app is not running' });
    const id = ++this.ids;
    const payload = Buffer.from(JSON.stringify({ id, command }), 'utf8');
    const head = Buffer.alloc(4);
    head.writeUInt32BE(payload.length, 0);
    return new Promise((resolve) => {
      const timer = setTimeout(() => {
        this.waiting.delete(id);
        resolve({ ok: false, error: 'no answer within ' + timeoutMs / 1000 + ' s' });
      }, timeoutMs);
      this.waiting.set(id, { resolve, timer });
      if (!this.child.stdin.write(Buffer.concat([head, payload]))) this.undrained();
    });
  }

  // node buffers what sae's stdin has not taken, without bound: past its
  // high-water mark a write says so, and 'drain' says when it has all gone.
  // Not gone within WRITE_MS: sae has stopped reading, and is dropped.
  undrained() {
    if (this.stall) return;
    this.stall = setTimeout(() => {
      this.stall = null;
      if (!this.child || this.child.stdin.writableLength === 0) return;
      this.stopped('the app stopped reading its channel: ' + this.child.stdin.writableLength +
        ' bytes not taken within ' + WRITE_MS / 1000 + ' s');
    }, WRITE_MS);
  }

  drained() {
    if (this.stall) clearTimeout(this.stall);
    this.stall = null;
  }

  // The channel is given up on (sae stopped reading): every request fails
  // with why. Without a window sae is stopped; with one it is the person's,
  // and only the channel goes (as in close()).
  stopped(why) {
    this.gone(why);
    if (this.headless) {
      this.kill();
      return;
    }
    try { this.child.stdin.destroy(); } catch (e) { /* gone already */ }
    this.child.stdout.destroy();
    this.child.unref();
  }

  // Close the channel. A headless sae goes by itself; one that has not
  // within `graceMs` is killed. A sae with a window is left open for the
  // person (its channel closed).
  async close(graceMs = 3000) {
    if (!this.child) return;
    const child = this.child;
    this.drained();
    try { child.stdin.end(); } catch (e) { /* gone already */ }
    if (!this.headless) {
      // Not waited for: it is the person's now.
      this.gone('the channel is closed');
      child.stdout.destroy();
      child.unref();
      return;
    }
    const stopped = await Promise.race([this.exited, new Promise((r) => setTimeout(() => r('timeout'), graceMs))]);
    if (stopped === 'timeout') this.kill();
  }

  kill() {
    if (this.child && this.child.exitCode === null) {
      try { this.child.kill('SIGKILL'); } catch (e) { /* gone already */ }
    }
  }
}

module.exports = { SaeChannel, ANSWER_MS, WRITE_MS };
