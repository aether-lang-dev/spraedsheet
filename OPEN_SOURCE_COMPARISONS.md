# Spreadsheet comparisons for Paul and Nic

Reviewed 18 October 2026 against Spraedsheet `cfc906f` and the local checkouts
listed below. These are source/documentation reviews, not executed performance
benchmarks or full product acceptance tests.

**Direction:** keep the Aether/AetherUI/sae workbook and its two formula languages.
Use these comparisons to improve reliability, repeatable automation, and editable
record workflows. The React/Tauri/DOM implementation is abandoned; it is not the
implementation baseline for any recommendation here.

## Read these comparisons

| Technology | Local revision | Main lesson | Comparison |
|---|---|---|---|
| LibreOffice Calc | `1fd35eecd` | Interacting workbook features and mature operational behaviour | [Calc](LIBREOFFICE_COMPARISON.md) |
| Grist | `edab14f` | Records, relations, linked views, incremental imports | [Grist](docs/comparisons/GRIST.md) |
| HyperFormula | `187dfa1` | Range-aware incremental calculation and explicit engine contracts | [HyperFormula](docs/comparisons/HYPERFORMULA.md) |
| Univer | `062b8ba` | Commands, persistent mutations, transient UI operations, extension boundaries | [Univer](docs/comparisons/UNIVER.md) |
| IronCalc | `297553d` | Reversible edit data and workbook-based compatibility checks | [IronCalc](docs/comparisons/IRONCALC.md) |
| Gnumeric | `e00c927` | Numerical quality and visible test maturity | [Gnumeric](docs/comparisons/GNUMERIC.md) |
| Gnumeric website | `7567afd` | Documentation and explicit implementation/test status | Included in the Gnumeric comparison |
| VisiData | `f9f3dbe` | Recordable commands, replay, source-aware data exploration | [VisiData](docs/comparisons/VISIDATA.md) |
| Airtable | Official documentation checked 18 October 2026 | Make structured data feel like an application | [Should we envy Airtable?](docs/comparisons/AIRTABLE.md) |

Gnumeric's two repositories serve different purposes: `gnumeric` contains the
application; `gnumeric-web` contains its website and supporting documentation.
The latter is not a second browser-based calculation engine. Airtable is a
proprietary product reference, separate from the open source checkout studies.

## The baseline has changed

[README.md](README.md), [the implementation tracker](docs/gridcraft_parity.md),
and [TODO.md](TODO.md) describe the current native app. The original
`ENHANCEMENT_PLAN.md` and `COMPARISON.md` belong to the earlier implementation and
are available in repository history. Several of their proposed features now exist.

The current app already has Excel and RexxJS formula modes, compiled references,
cross-sheet recalculation and cycle detection, tables and structured references,
RexxJS range queries, auto IDs, contextual validation, grouped workbook undo,
auditing, pivots, charts, analysis tools, and several file formats. Commands reach
the same model from panels, REXX, the control bus, and MCP. There is no reason to
schedule these again as absent features.

Source observations refine the next steps:

- `app/sheet/address.ts` still sets a sheet to 26 columns by 100 rows. Larger-file
  work requires changes to addressing, model behaviour, and the grid together.
- `language.ts` describes tables through ranges, names, column names and metadata.
  That is a foundation for records, but it is not yet a general relation model
  with immutable record and field identities.
- `query.ts` materialises rows and groups using textual keys. Stronger type/null
  contracts and source lineage matter before editable joins or saved views.
- `commands.ts` journals operations; workbook snapshots group undo; scripts can
  execute within an all-or-nothing workbook boundary. Recording those actions as
  editable REXX remains explicit work in TODO 7.1.
- `sheet.ts`, `workbook.ts`, and `language.ts` already separate calculation from
  UI and await evaluation. The next question is correctness and cost at scale,
  rather than introducing that separation.

These findings come from reading the named files and existing test definitions;
they are not new runtime verification results.

## What to build next

| Order | Deliverable | Why now | Original acceptance workload |
|---|---|---|---|
| 1 | Behaviour and scale harness | Gives the existing broad feature set a trustworthy depth measure | Cross-sheet edit → undo → redo → save/open; running totals, dynamic references, spills, controlled volatile functions |
| 2 | Record UI commands as REXX | Directly extends the command architecture and existing TODO 7.1 | Record import/filter/edit/export; replay in a fresh workbook; compare semantic state |
| 3 | Stable table/record/field identities and keyed reconciliation | Makes the object↔sheet direction safe | Import twice by key; show insert/update/delete preview; sort and rename; apply an edit to the intended record |
| 4 | Saved views and a record detail form | Gives non-programmers a useful application over the same workbook | Two views edit one record; reopen preserves view settings; form and grid enforce the chosen validation policy |
| 5 | Relations, lookups and rollups | Adds capability beyond range-based tables | Customer rename preserves order links; aggregate refreshes; deleting a referenced record follows a declared policy |
| 6 | Triggered REXX with run history | Builds on recording, record identity, and committed changes | Condition transition runs once; self-trigger loop stops; retries do not duplicate an external action |

The order is a proposal, not a replacement for the existing tracker. Numerical
checks, metadata completeness, keyboard accessibility, and import-loss reporting
can improve alongside these milestones. Make headless and interactive workflows
agree before claiming scale or compatibility.

## How to use the references

Each comparison separates observed capabilities from independently proposed work
and lists local evidence paths. No third-party source, fixtures, assets, or
algorithms are transplanted. Source-informed design is not a formal clean-room
process; if that distinction matters for later implementation, write behavioural
requirements and have an independent implementation process established first.

The GridCraft restriction in `AGENTS.md` remains in force. Only this repository's
existing feature reference and tracker were consulted; its external repository,
documentation, and screenshots were not read. No upstream checkout was changed.
