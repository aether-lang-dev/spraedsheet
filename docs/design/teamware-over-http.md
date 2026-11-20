# Teamware over HTTP: one workbook, a small team, one in-house server

Status: built, 2026-11-19 (slices 1 to 4). Written before the build as the
design; section 8 says what was built, what was measured and where it
differs.

Paul's brief: "Getting something like Spreadsheet to work over http would be
great. Sure it's not a google-sheets but for in-house teamware it could be
great, especially with creativeness around data-at-rest in a backend."

The aim: a small team (two to twenty people) each opens the same workbook
from their own sae, served by one server on the team's network, and sees
each other's edits within a second. Not character-by-character co-editing:
an edit is a whole command (a cell set, a fill, a format, a sort). Honest
about concurrency: no edit is ever silently lost; when two people change the
same thing, the second is told, plainly, and shown what won. And the
workbook lives on the server's disk in a form an ops person trusts and a
team can read with ordinary tools.

## 1. Where the spreadsheet runs: browser mode

**Decision: browser mode.** The server serves the app's own pages
(`app/sheet.ts` and its modules, unchanged and not forked) and the workbook
API from one origin. Each person runs `sae http://teamserver:8390/sheet?book=budget`.
The model (workbook, both formula languages, every command) runs in each
person's sae, as it does today; the server holds the truth (the operation
log) and never runs the model.

Why not app mode pointed at a server:

- App mode reaches a server only through `app.json`'s `capabilities.http`,
  a fixed list of URL prefixes in the bundle. Every team's server has its
  own address, so each team would edit the bundle's manifest, which is what
  the installer was shown and (once signing lands, roadmap 2.3) what is
  signed. Browser mode needs no grant: a page reaches its own origin by
  construction.
- Browser mode's rules are what a team server wants: same-origin by
  construction, no ambient credentials (no cookie jar; the page keeps its
  token in per-origin `storage` and sends it itself, so CSRF cannot exist),
  no `fs`, no `shell`.

What the app uses today that browser mode does not allow, measured by
loading `/sheet` from a page server in browser mode (`SAE_NO_WINDOW=1`,
sae 9c70240, 2026-11-19: it loads and runs, 0.41 s, within the 32 MB heap):

| The app uses | In browser mode | Effect on the team sheet |
|---|---|---|
| `documents` (open/save files the person picks) | absent | Open/Save/Export to a local file are unavailable; the server is where the workbook lives, and its exports (section 4) are where an `.xlsx` comes from |
| `fs` (`$APPDATA/versions.git`, the version repository) | absent | the version panel says so; the operation log is the history instead (section 4) |
| `bus` (control bus) | absent | REXX and MCP drive the app, not the team page |
| `shell.open`, `printing` | absent | links and printing unavailable |
| the menu bar | absent ("only an installed app has a menu bar") | the toolbar and panels do everything the menus do |
| a 512 MB heap (`app.json` `memory`) | 32 MB per page | fine for team-sized workbooks (thousands of cells); a large one would not fit: **decision P1** |

Every one of these is already declared `"seeks ..., reduced functionality
without"` at the top of `app/sheet.ts`, so the page loads without them. The
one addition is `"seeks outgoing-http GET,POST /api/**, reduced functionality
without"`: in the browser that is its own origin's API; in an installed app
without a `capabilities.http` grant it is absent and the team code stays off.

The team code is one module, `app/sheet/team.ts` (the client: sign in, open,
send, follow), and one pure module, `app/sheet/teamtext.ts` (the keyed text
and its deltas, section 2), both tested headless. `sheet.ts` turns team mode
on when the page came from an `http(s)` origin with `?book=` in its URL.

## 2. The unit of change: a command and the lines it changed

Every user action is already a command with an id and string parameters
(`commands.ts`), and the workbook already has a canonical text form
(`textform.ts`: one cell per line, a sheet's names, rules and settings an
entry per line, keys sorted, built for diffing and kept in git by
`gitstore.ts`). An operation is both:

```json
{"id":"7f3c...","base":41,"cmd":"cell.set cell=B2 input=5",
 "structure":[],
 "del":[["Sheet1\tB2","B2\t4"]],
 "put":[["Sheet1\tB2","B2\t5"]]}
```

- `cmd` is the command line the person ran, with the parameters as the
  command resolved them (`Outcome.resolved`: a default taken from the
  selection made explicit), as Record Macro keeps it. It is for people: the
  history panel says "Ann: Set B2 to 5". Replaying it is not how others
  apply the op, because a command can read state that is the author's alone
  (`edit.paste` pastes the author's clipboard; Ctrl+; types the author's
  clock), and a command run on a newer workbook than its author saw can
  read cells someone else has since changed.
- `del` and `put` are the lines of the text form the command changed, each
  under its key, with the line as it was and as it became. A key is the
  sheet and the cell (`Sheet1\tB2`), the sheet and the entry
  (`Sheet1\t@names\tQTY`), a whole list of a sheet (`Sheet1\t@rules[]`, its
  elements in order), a workbook setting (`\t@formulas`), or the order of
  the sheets (`\t@sheet`). This is what everyone applies, so every sae and
  the server fold the same deltas and cannot drift apart.
- The client works the delta out by writing its workbook as text before and
  after the command (measured headless on the Mac: 8 ms at 500 cells, 44 ms
  at 2,000, 120 ms at 5,000), and keeps the last text so a command costs one
  write.
- `structure` names the sheets whose shape the command changed: rows or
  columns inserted, deleted or moved, a sort, a cut and paste (`rows.insert`,
  `rows.delete`, `columns.insert`, `columns.delete`, `columns.move`,
  `data.sort*`, `table.sort`, `data.subtotal`, `data.removeDuplicates`,
  `edit.paste` of a cut).
- `id` is made by the client, so a retried post (an answer lost on the way
  back) is recognised and answered with the revision it already has,
  never appended twice.

The server adds `rev` (its order), `who` (from the token, never from the
client) and `at` (its clock).

Only commands whose effect is `edit` are sent. `view` commands (selection,
zoom, the panel shown) stay each person's own; `query` changes nothing;
`external` commands that would replace the workbook (`file.open`,
`file.importText`, `version.restore`) are refused in a team workbook with a
message saying so.

## 3. Concurrency: compare-and-set on lines, a server-ordered log

The server keeps one ordered log per workbook and the workbook's keyed text
at its head. An op made against revision N is accepted when, for every key
it touches, the server's line now is the line the op says it was (`del`),
and no op after N changed the shape of a sheet the op touches (or, for a
structural op, touched that sheet at all). Then it is appended as revision
head+1, durable on disk before the answer goes back. Otherwise it is
refused, `409`, naming the first key in conflict and who changed it, when,
in which revision, with what command.

This is compare-and-set per line, which needs no merge logic and no model
on the server, and is simple to state:

| Two edits from the same starting revision | Result |
|---|---|
| different cells, entries or settings | both kept, in the server's order (the second is "rebased": its lines are applied on the newer text, which is safe because no line of it changed) |
| the same cell (its input, look, note, validation: all on its one line) | the second refused |
| the same named range, macro, style, or workbook setting | the second refused |
| two changes to a sheet's list of conditional-format rules, charts, merges, tables, drawings ... (a list is one key) | the second refused |
| a structural change (insert/delete/move rows or columns, sort, cut and paste) and any other change to the same sheet | the second refused |
| two sheets added, renamed, moved or deleted (the order of sheets is one key) | the second refused |
| an edit to a cell a formula reads, and an edit to the formula | both kept; the formula is worked out again on every sae, as in any spreadsheet |

What the person sees when theirs is refused: their sae had already shown
the edit (it applies locally first), so it puts back the server's workbook
and says so in the team line above the grid, in words:

> Not kept: your Set B2 to 5. Ann changed Sheet1!B2 at 14:02 (revision 41)
> after you began; the sheet shows Ann's now. Make your change again to
> replace it.

Edits made after the refused one, still waiting to be sent, were made on
top of it and are dropped with it; the line says how many.

**Undo in a team workbook** takes back the person's own last kept edit, as
a new op whose lines are that edit's lines swapped (`put` becomes `del`). The
same compare-and-set applies: if someone has since changed one of those
cells, the undo is refused and says whose change it would overwrite. Redo
sends the edit again. (The app's own undo stack holds whole-workbook steps
that include other people's edits, so it is not used in team mode.)

**Hearing others' edits: long-poll.** `GET /api/books/<b>/ops?since=N&wait=10`
answers at once with every op after N, or holds the request until one
arrives or ten seconds pass. Measured before choosing (section 8 has the
numbers): Aether's HTTP server runs each request on a worker thread (cores ×
2, `AETHER_HTTP_WORKERS` to raise it), so a held poll costs a thread and
nothing else, and SSE would cost the same; sae's page `http` delivers a
response whole, so SSE would need a new streaming page API, while a
long-poll is an ordinary request. sae ran each request on one of four fetcher
actors, and a blocking call on an actor holds its scheduler core, so a held
poll delayed other requests for up to its wait; sae is changed to give each
request's call a thread of its own (section 8).

A remote op is applied to the person's workbook by patching its text and
reading the workbook back from the text (`file.importText`'s path: 46 ms at
2,000 cells on the Mac), keeping the sheet they are on and their selection;
an op that only sets plain cell inputs is applied to those cells directly,
without the reread.

## 4. Data at rest: a directory per workbook, an append-only log as the record

```
<data>/
  people                 one line per person: name, a tab, the SHA-256 of their token
  books/<book>/
    ops.log              the record: one JSON object per line, one op per line, append-only
    snapshots/r00000500.sheet.txt   the workbook's text at revision 500 (every 500 ops, and at shutdown)
    head.sheet.txt       the workbook's text now (rewritten with each snapshot)
    checkpoints/<label>.sheet.txt   named checkpoints: the text at a revision, named by a person
```

- **The log is the primary record.** Every op ever accepted, with who, when,
  against which revision and the command in words. It gives the audit trail
  ("who changed this cell, and when": every op that touched its key), time
  travel (the workbook at any revision: the nearest snapshot at or before it,
  plus the ops after), undo across sessions, and replay as a test (the
  server's text after folding the log must equal what each sae shows).
  Revision 1 is the workbook's creation: a `put` of every line of its
  starting text (empty, or a text the creator uploads).
- **Crash safety.** An op is written and `fsync`ed before it is
  acknowledged, so an acknowledged edit survives a power cut. A line is
  written whole with its trailing newline; at open, a last line with no
  newline or that does not parse is a torn write from a crash: it is cut off
  (the file truncated to the last whole line) and the server says so on its
  console. Nothing torn was ever acknowledged, so nothing acknowledged is
  lost. Snapshots and `head.sheet.txt` are written to a temporary file,
  `fsync`ed and renamed into place, so a crash leaves the old one or the new
  one, never half; and they are only a cache: deleted, they are rebuilt from
  the log.
- **Never trapped.** The snapshots are the app's documented text form,
  which the app opens (`IMPORTTEXT`, `file.importText`) and `git diff` reads
  cell by cell; the log is JSON lines, which `jq`, `grep` and any script
  read. Exports in other tools' formats (an `.xlsx` through the app's own
  writer, a CSV of values per sheet for BI tools) need the model, which the
  server does not run: they are a later slice, made by a headless sae the
  server drives over its control channel (sae's `docs/app-backend.md`), not
  by a second implementation of the spreadsheet.
- **Plain enough to back up and version.** A workbook is a directory of
  text files: `rsync` it, put it in git (the snapshots diff by cell), or
  copy one out to take a workbook to another server. No database to
  administer.
- **Named checkpoints** ("before the Q3 numbers") are the text at a
  revision, kept under the name; the log still has everything after.

Not done in v1: attachments (the app keeps pictures inside the workbook's
text as base64; content-addressing them is the obvious next step if
workbooks with pictures get large), compaction of a long log (the
snapshots keep opening fast; the log is kept whole on purpose).

## 5. Who are you

No ambient credentials. The server's `people` file has a line per person:
their name and the SHA-256 of a token. `sheet-server add-person <data> ann`
makes a random token, writes the line and prints the token once; the person
pastes it into the team line's sign-in field the first time, and the page
keeps it in its origin's `storage` and sends it as `Authorization: Bearer`
on each request it makes. The server knows who made each op from the
token, never from what the client says. Removing the line revokes it.

Presence comes cheaply with the long-poll: each poll carries the cell the
person is on, and each answer lists who polled in the last 30 seconds and
where they are, which the team line shows ("Ann on B4, Bob on Costs!C9").

v1 serves plain `http:` on the team's network. Aether's server terminates
TLS (`server_set_tls`), and the server takes a certificate and key when
given them; sae refuses an `https:` page that fetches `http:`, so the two
are never mixed.

## 6. The server

`server/sheetserver.ae`, one Aether program built with `ae build` like
sae's tools: `sheet-server <data-dir> <app-dir> [port]`. It serves the
app's pages (`/sheet`, its modules) from `<app-dir>` and the workbook API
under `/api` from one origin.

No global mutable state: a `Library` holds the data directory and the open
workbooks; each `Book` holds its own revision, keyed text, per-key last
writer, and log file. Handlers run on several worker threads at once, and
Aether has no mutex in std, so all of a library's state is owned by one
actor (`LibraryActor`) and handlers ask it (`library ? Append {...}`); the
actor serialises every change to every workbook, which for an in-house
team is far below its capacity (an op costs an `fsync`). Long-polls wait
outside the actor, watching the book's revision through a `std.sync` atomic
cell, and ask the actor only when there is something new.

API (JSON; every `/api` request but `GET /api/info` needs the bearer token):

| Request | Answer |
|---|---|
| `GET /api/me` | `{name}` |
| `GET /api/books` | `[{name, rev}]` |
| `POST /api/books/<b>` (body: a workbook text, or empty) | `{rev}`; `409` if it exists |
| `GET /api/books/<b>` | `{rev, text}`: the workbook now |
| `GET /api/books/<b>/ops?since=N&wait=S&cell=K` | `{rev, ops, people}`, held up to S seconds (at most 10) when nothing is new |
| `POST /api/books/<b>/ops` (an op) | `{rev}`, or `409 {conflict: {key, who, at, rev, cmd}}` |
| `GET /api/books/<b>/text?rev=R` | the workbook's text at revision R |
| `GET /api/books/<b>/history?key=K` | every op that changed key K: `[{rev, who, at, cmd, was, now}]` |
| `POST /api/books/<b>/checkpoints/<label>` | `{rev}` |

## 7. Decisions for Paul

- **P1. A bigger heap for a team page.** Browser pages get 32 MB; the app
  asks 512. Team-sized workbooks fit; a 100,000-cell one would not. sae's
  roadmap lists "a per-origin memory budget" as to do; a team server is the
  first page that wants one. Options: the person grants an origin more
  (asked once, kept in sae's own settings), or a server's answer asks for it
  (refused above a cap). Not built.
- **P2. Plain http on the team's network for v1**, TLS when given a
  certificate. A team that wants TLS without managing certificates would
  need something more (a self-signed certificate pinned on first use is the
  web's mistake; better not).
- **P3. Undo in team mode is "take back my last kept edit"**, refused when
  someone has changed those cells since. The alternative (undo applies
  over others' later edits) loses their work silently.
- **P4. Structural conflicts are conservative**: a sort or row insert and
  any concurrent edit to that sheet, from the same starting revision,
  conflict, even when they would have composed (an edit far below an
  inserted row could be shifted down instead). Shifting is possible later
  (the app already rewrites addresses for inserts in `shift.ts`); refusing is
  easy to explain and never wrong.
- **P5. Exports** (xlsx, values CSV) made by a headless sae the server
  drives, rather than any second implementation; a later slice.
- **P6. Tokens per person in a file**, no passwords, no accounts service.
  Enough for in-house; a directory (LDAP, OIDC) is a later question.

## 8. As built (2026-11-19)

Slices, each committed with its tests:

1. **The server** (`server/sheetserver.ae`, about 1,000 lines of Aether): the
   directory layout of section 4, the API of section 6, compare-and-set and
   the shape check of section 3, ops kept once by id, presence from the polls.
   `tests/server/server.test.js` runs the real binary over real HTTP on real
   files; each guard (compare-and-set, the torn-line cut, de-duplication,
   the shape check) was taken out in turn and its test failed.
2. **The page** (`app/sheet/team.ts`, `teamtext.ts`, `view/teambar.ts`, and
   `Commands.team` in `commands.ts`): browser mode, as section 1 chose; no
   page or model forked. `tests/server/team.test.js` runs two people's real
   Workbook and Commands on node through the real server. sae grew one thing
   (below).
3. **Two sae on one workbook** (`tests/specs/spec_team.ae`, run by
   `spec_team.sh`): two sae-drivers on cachyos (Xvfb), signing in, edits
   both ways, presence, a true conflict and what the person sees, undo, a row
   insert, history and an earlier revision; screenshots of both windows at
   each step under `target/spec-team/shots/`.
4. **History in the page**: "Who changed this cell" (every kept change to the
   selected address, oldest first) and "Look at revision" (the workbook at any
   revision, opened beside the team one in the read view).

Measured:

- Long-poll: an op reached the other person's poll 12 to 26 ms after it was
  kept (Mac, server tests), and the other person's workbook 105 to 128 ms
  after the edit (node, two models through the server: the poll, the
  re-read of the workbook from its text, the recalculation).
- Writing the workbook's text, which each local edit costs: 8 ms at 500
  cells, 44 ms at 2,000, 120 ms at 5,000 (Mac, sae's QuickJS); reading it
  back, which each of someone else's edits costs: 10, 46, 124 ms.
- Durability under SIGKILL: 5,751 to 7,054 ops acknowledged in 0.7 s by two
  writers, every one of them in the log after a restart (Mac). macOS's
  `fsync` does not force the drive's cache (that needs `F_FULLFSYNC`); on
  Linux it does. A power cut on a Mac server could lose the last
  acknowledged ops; Aether's `std.fs` offers no full sync today.
- sae ran each page request's blocking call on one of four fetcher actors.
  A blocking call on an actor holds its scheduler core, and every actor
  sharing it: on cachyos (4 cores), eight actors each fetching once beside a
  3 s request, one quick request took the full 3001 ms; sae's own spec
  ("a request held open does not hold up the next ones", five quick requests
  beside a 3 s one) showed one taking 3012 ms, and choosing an idle fetcher
  did not help (3013 ms: the core, not the fetcher, was the queue). In
  spec_team, Ann's long-poll answers came only after her own next edit. sae
  now runs each request's call on a thread of its own (std.worker's
  `run_detached`): the spec passes on cachyos (49 of 49 in spec_nav; the
  build without the change fails that one, 3012 ms), and spec_team passes.

Differences from the design:

- **A refusal drops only the refused op.** The design dropped the person's
  later waiting edits with it; that lost a burst of typing for one conflict.
  The later ones are sent as they are, and the server refuses any whose
  lines the refused one had changed.
- **The server re-reads the whole log at open** (the snapshots are for people,
  tools and time travel, not yet for opening). Fine for a team's workbooks;
  a log of millions of ops would want the snapshot plus its tail.
- **History is by address.** A row insert or a sort moves cells, and the
  history of B2 lists what was put at B2, whichever cell that was then.
- **The team line is two lines**: the workbook and its revision (and who is
  here) on one, a refusal on its own line until the person's next edit.
- **Another person's edit re-reads the workbook from its text** (about as long
  as writing it); the plain-cell fast path of section 3 is not built.
- **A command's commands**: a script or macro is one op (the outermost
  command's line, everything it changed).
- **Exports** (P5), **TLS** (P2) and **a bigger heap** (P1) are not built.

Verified on cachyos (Linux, GTK4 under Xvfb, 4 cores): spec_team 8 of 8, on
sae before and after rebasing on main; spraedsheet's whole run-tests.sh with
it (spec_sheet, the app-mode specs, the MCP and CLI tests). With the
long-poll answered promptly, the conflict step had to press both Returns
within a few milliseconds: Bob otherwise heard Ann's B2 before sending his,
and his was rightly kept as a later edit.

Found on the way:

- Aether: `std.http`'s `http_get_query_param` hands a value over as sent,
  still percent-encoded; the server decodes it with `std.url`.
- Aether: `string.string_pad_start(s, 8, "0")` type-checks (the pad is an
  `int` in the extern) and fails in the C compiler.
- macOS kills (SIGKILL) a new run of a binary rewritten in place while an
  older copy of it runs; the build scripts build aside and move it in.
