# RexxSheet: a spreadsheet you can script in REXX

A spreadsheet with two formula languages, Excel's and RexxJS's, and every
action a command that a REXX script can drive. This repository has four
parts:

| Folder | What it is | License |
|---|---|---|
| [`app/`](app/) | **The spreadsheet, current.** A [sae](https://github.com/aether-lang-dev/sae-it-aint-so) app in sae's TypeScript dialect: it runs as a page in the sae browser or as an installed app, natively on aether-ui (GTK4, AppKit, Win32). | MIT |
| [`rexx/`](rexx/) | REXX scripts that drive the running sae app from outside, over its control bus (`ADDRESS SPREADSHEET`, through sae's `sae-bus`). | MIT |
| [`server/`](server/) | `sheet-server`: one workbook, a small team, one in-house server ([Teamware over HTTP](#teamware-over-http)). An Aether program. | MIT |
| [`mcp/`](mcp/) | An MCP server, so an AI client (Claude, or any MCP client) drives the app: one it starts itself and talks to over sae's stdin and stdout (no file, nothing listening), or with `--attach` one already running with its control bus on. | MIT |

The original React + Tauri app the sae one grew from was here as `tauri/`
until 2026-10-15; it is in this repository's history (last at commit
`bd1a002`), AGPL as it was.

## The sae app

`app/` began in sae as its demo 18 and moved here on 2026-10-14; its
history up to then is in sae's git log (to sae commit `61e4adb`). What it
does and what is still to do: [docs/spreadsheet_demo_done_and_todo.md](docs/spreadsheet_demo_done_and_todo.md).
The bugs it found in the Tauri app: [docs/tauri-spreadsheet-upstream-notes.md](docs/tauri-spreadsheet-upstream-notes.md).
Where it is going: [TODO.md](TODO.md), small tested steps toward the
feature set of GridCraft (where it stands: [docs/comparisons/gridcraft_parity.md](docs/comparisons/gridcraft_parity.md);
the yardstick: [docs/comparisons/gridcraft_feature_reference.md](docs/comparisons/gridcraft_feature_reference.md),
its source never read) and a toolbar and menus over the grid, each step
reachable from REXX.

Action lists drawn from other spreadsheets for the current Aether/AetherUI/sae app:
[LibreOffice, Grist, HyperFormula, Univer, IronCalc, Gnumeric, VisiData, and
Airtable](docs/comparisons/README.md).

```
app/
  app.json          the installed app's manifest: its start page and grants
                    (documents, the control bus)
  sheet.ts          the page: panels declared over the commands
  sheet/            the spreadsheet with no ui: workbook, sheets, both formula
                    languages, every command (commands.ts), RexxJS (rexxjs.js,
                    generated)
  sheet/view/       the ui: each panel reads the workbook and runs commands
tests/
  headless/         the model and commands, no window: sae's lowerer harness
                    on sae's QuickJS (sheet -> ../../app/sheet)
  specs/            the app driven in sae-driver: grid, keys, clipboard, files,
                    the control bus, rexx/'s scripts and mcp/'s server run against it
  lib/griddriver.ae reads and clicks the datagrid for the specs
  mcp/              the MCP server's tests (node --test): its protocol, with
                    stand-ins for sae-bus and sae; and the real app driven
                    over its control channel (channel.test.js)
vendor/rexxjs/      RexxJS's interpreter and parser, vendored (MIT, its LICENSE)
tools/              bundle-rexxjs.mjs: re-vendors RexxJS and writes app/sheet/rexxjs.js
docs/               the plan, and notes for the Tauri app
rexx/               REXX scripts that drive the app, and their ADDRESS library
mcp/                an MCP server, so an AI client drives the app (sae-channel.js:
                    the app started on sae's stdin and stdout)
sae -> ../sae       the sae checkout it builds and tests against
RexxJS -> ../RexxJS the RexxJS checkout the REXX scripts run on
```

### Running it

It needs sae checked out beside this repository (the `sae` symlink points
at `../sae`) and built, as sae's AGENTS.md "Build and test" says. Then:

```bash
../sae/target/build/bin/sae --app app
```

`--control-bus` after it switches on the control bus, so other programs on
this machine can drive the sheet (sae's `docs/control-bus.md`): a loopback
port, and its key in a file only you can read. A program that starts the
app itself needs neither: the MCP server and `cli/sheet.js` start it with
`--control-stdio` and drive it over its own stdin and stdout, which only
they hold (sae's `docs/app-backend.md`, "The control channel").

### Testing it

```bash
./run-tests.sh
```

It needs `node`, and RexxJS checked out beside this repository, for the
REXX scripts' spec and the MCP server's tests (and, off Linux, `lsof`, for
what listens). That runs the headless tests, the MCP server's tests and the
specs, one line of PASS or FAIL
each, with logs under `target/run-tests/`. `./run-tests.sh headless` runs
only the headless tests; `./run-tests.sh spec_sheet` only that spec.
`SAE_DIR=...` picks another sae checkout. On Linux with no display it runs
under `xvfb-run`.

### Re-vendoring RexxJS

```bash
node tools/bundle-rexxjs.mjs ../RexxJS
```

copies RexxJS's interpreter and parser into `vendor/rexxjs/` (recording the
commit in `UPSTREAM`) and regenerates `app/sheet/rexxjs.js`. RexxJS is the
canonical source of the REXX language; changes to it go there first.

## Teamware over HTTP

A small team can share one workbook: one person runs `sheet-server` on the
team's network, and everyone opens the workbook from it in their own sae.
Each person's sae runs the spreadsheet as it always does; the server keeps
the order of everyone's edits, refuses the one that comes second when two
people change the same cell (and says so, by name, in the team line above
the grid), and keeps every workbook on its disk as plain files. The design,
and what is Paul's to decide: [docs/design/teamware-over-http.md](docs/design/teamware-over-http.md).

Run the server (Aether 0.923 or later; the pages it serves are this
checkout's `app/`):

```bash
ae build server/sheetserver.ae -o target/sheet-server
target/sheet-server add-person /srv/sheets ann     # prints Ann's token, once
target/sheet-server add-person /srv/sheets bob
target/sheet-server serve /srv/sheets app 8390
```

Each person then opens `http://<server>:8390/sheet?book=budget` in sae
(`sae http://...`), pastes their token into the team line once (the page keeps
it in that server's storage), and edits. A workbook that does not exist yet
is made from what the first person's page shows. `/srv/sheets/people` is the
list of who may sign in (a name and the SHA-256 of their token); take a line
out to stop someone.

What each person sees:

- Others' edits within a moment (a long-poll; about 0.1 to 0.2 s on a LAN).
- `Here: Ann on Sheet1!B4`: who else has it open, and where.
- A refusal in words when someone changed the same thing first: `Not kept:
  your B2 = 200. ann changed Sheet1!B2 (B2 = 100) at 14:02 (revision 6) after
  you began; the sheet shows theirs now. Make your change again to replace
  it.` Edits to different cells never conflict; a row insert, a sort or a
  cut and paste conflicts with any edit to that sheet made from before it.
- Undo takes back your own last kept edit (refused if someone has since
  changed those cells); Redo sends it again.
- **Who changed this cell** lists every kept change to the selected address;
  **Look at revision** opens the workbook as it was at any revision, beside
  the team one, in the read view.

On the server, each workbook is a folder: `ops.log` (every edit ever kept,
one JSON line each: who, when, the command, the lines it changed; written
and fsynced before the person's sae is told), `head.sheet.txt` (the workbook
now, in the app's text form, which `IMPORTTEXT` opens and `git diff` reads
cell by cell), `snapshots/` (the text every 500 revisions) and
`checkpoints/`. Back it up with rsync, or keep it in git; there is no
database. A crash mid-write leaves a torn last line, which the server cuts
off when it starts and says so; nothing it had acknowledged is lost.

Not yet: exports to `.xlsx` and CSV on the server (the page's own exports
need the installed app), TLS (plain http on the team's network for now), a
workbook bigger than a browser page's 32 MB heap.

## REXX scripts

With the app started with `--control-bus`, from the repository's root:

```bash
RexxJS/core/rexx rexx/test-spreadsheet-address.rexx
```

See [rexx/README.md](rexx/README.md): how a command reaches the app, and
where the scripts differ from the Tauri app's.
