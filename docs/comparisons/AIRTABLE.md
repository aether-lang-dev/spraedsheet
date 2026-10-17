# Should Spraedsheet envy Airtable?

**Yes: Airtable makes structured data usable as an application.** Its most useful
lesson is how record identity, field types, relationships, saved views, forms,
and automation fit together. For Spraedsheet, the opportunity is to offer that
coherence alongside free-form worksheets and editable REXX automation.

This is a product assessment using Airtable's official documentation checked
18 October 2026. Airtable is proprietary; there is no local source comparison.
Availability varies by feature and plan. No pricing or quota assumptions are
needed for these recommendations. Spraedsheet baseline:
[comparison index](../../OPEN_SOURCE_COMPARISONS.md).

## What is worth envying

| Airtable capability | Why users care | Independently proposed Spraedsheet direction |
|---|---|---|
| Typed records and fields | A customer, task, or order has an understandable shape | Strengthen selected tables with stable record/field identities and schema-backed editors |
| Linked records, lookups, rollups | Relationships survive changes to display names and ordering | Reference fields between tables, explicit cardinality, and dependency-aware aggregates |
| Several views over one table | People can organise their work without copying the data | Saved filter/sort/group/column settings, then cards, board and calendar views |
| Record details and forms | Editing does not require navigating a wide spreadsheet | A generated record form using the existing native field widgets |
| Interfaces for a particular job | An operator sees the fields and actions needed for the task | Small native workbenches bound to table views and registered commands |
| Trigger/action automation and history | Routine work happens consistently and failures can be inspected | Recorded REXX, explicit triggers, test mode, and a durable run log |

Airtable documents linked-record fields and related lookup/rollup fields. These
give relationships a visible place in the product rather than making every user
write matching formulas. Its rollup documentation also specifies that filtering a
table view does not, by default, change a rollup over linked records. That is an
example of semantics we should declare explicitly.
[Linking records](https://support.airtable.com/articles/3370222027-linking-records-in-airtable),
[rollups](https://support.airtable.com/articles/7497685062-rollup-field-overview).

Its saved views configure presentations over table data, while Interface Designer
adds task-focused pages such as lists, dashboards, record review and forms. This
reduces the amount a user must understand before completing a job.
[Views](https://support.airtable.com/articles/5189551686-getting-started-with-airtable-views),
[Interface Designer](https://support.airtable.com/articles/8078126534-getting-started-with-airtable-interface-designer),
[form views](https://support.airtable.com/articles/8588280258-getting-started-with-airtable-form-views).

## How this differs from the app we already have

Spraedsheet has tables, column metadata, auto IDs, calculated columns, validation,
queries, filters, slicers, pivots, charts, and a large command surface. Those are
substantial building blocks. They do not yet amount to a record application:

- A visible auto-ID column is not by itself an immutable identity contract for
  records imported, renamed, sorted, deleted, or edited through a view.
- Sorting/filtering a grid is not yet several independently saved table views
  with a shared underlying record model.
- A generated record form needs a schema and field policies, not merely a set of
  adjacent cell editors.
- Executing REXX on demand is not yet a trigger scheduler with retained execution
  history, retries, and explicit effect handling.

Airtable's computed-field model is also different from arbitrary cell formulas.
Its formula-field documentation treats formulas as field computations. Preserve
Spraedsheet's free-form Excel/RexxJS worksheets; a record-oriented table can add
schema and calculated-field behaviour as an explicit mode.
[Formula fields](https://support.airtable.com/articles/8772429326-getting-started-with-the-airtable-formula-field).

## A concrete product slice

Build a small **customers and orders workbook** before a general interface
designer:

1. Import customer objects with a business key. Preview changes on the next
   import and keep internal record identities stable. Persist field identities
   separately from names and positions.
2. Define Orders → Customer as a reference. Pick a customer by a readable label;
   store the identity. Define what happens when that customer is deleted.
3. Offer saved views: all orders, unpaid orders, and orders grouped by status.
   Add a record detail pane. Edits in either place call the same commands and
   enforce the chosen validation contract.
4. Show customer totals over related orders. State whether a total uses all
   related records, the selected view, or an explicit predicate. Make the choice
   inspectable in formulas and commands.
5. Record a workflow as REXX, then parameterise it. Add a condition-transition
   trigger for a committed order changing to Ready, with a preview and run log.

This would demonstrate the object↔sheet direction with a recognisable user
outcome. A native sae record pane can use AetherUI components and existing grants;
missing platform primitives belong in sae/AetherUI with platform tests.

## Automation is more than a trigger menu

Airtable documents record, condition, view, form, scheduled, and webhook triggers,
actions, run history, and server-side execution while nobody has the base open.
It also explains that condition/view triggers have transition semantics and
existing matching records are not automatically backfilled when an automation is
enabled. These distinctions prevent surprising execution.
[Automations](https://support.airtable.com/articles/3669392397-getting-started-with-airtable-automations).

For our implementation, define whether a trigger observes individual writes or
the final committed transaction; start with final committed state. Retain input
record identities, script version, outcomes, errors, and effect attempts. Prevent
self-trigger loops and make retries idempotent where an external service permits
it. Do not promise exactly-once external effects or pretend workbook undo can
recall a sent message.

A local app cannot execute while its process is stopped. Start with triggers
while the workbook is open; an always-running scheduler/worker is a separately
specified deployment feature. Scheduled work needs a declared clock/timezone and
restart policy. The existing control bus is a useful execution entry point, not
an always-on scheduler by itself.

## What would require a larger commitment

Multi-user sharing, public forms, identity management, conflict handling, access
control, and hosted execution are a separate product investment. A saved view's
hidden columns should not be presented as access enforcement. SaaS-style sharing
would require checks at every data/command boundary, including REXX and MCP.
Airtable distinguishes access permissions from field/table editing permissions;
our design would need equally clear contracts.
[Permissions](https://support.airtable.com/articles/3994336578-airtable-permissions-overview),
[field/table editing](https://support.airtable.com/articles/2296468758-airtable-field-and-table-editing-permissions).

We should not treat a canvas designer, every integration connector, or cloud
collaboration as a prerequisite for useful record views. Start with one local
workbook, a schema-driven form, and reproducible REXX workflows. Preserve portable
data and scripts, and give every view action a command equivalent.

## Recommendation

Envy the ease of turning structured data into a useful working application.
Build record identity, reconciliation, saved views, and a detail form first;
follow with relations and observable REXX triggers. Grist is the closest open
source reference for that model, and VisiData is particularly useful for making
the workflow recordable. Airtable supplies a product-quality target rather than
an implementation template.
