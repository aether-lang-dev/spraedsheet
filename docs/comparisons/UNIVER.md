# Univer compared with Spraedsheet

**Most useful lesson: distinguish a user's command, the durable workbook changes
it causes, and temporary view operations.** Our existing command architecture is
well placed to make that distinction useful for recording and extensions.

Review: `/home/paul/scm/univer`, revision `062b8ba`. Baseline and limits:
[comparison index](../../OPEN_SOURCE_COMPARISONS.md).

## What is present

Univer is an office SDK with plugins, services, a canvas renderer, a formula
engine, and Facade APIs for browser and Node usage. Its command service explicitly
separates commands that orchestrate work, mutations of persisted data, and
operations on transient state. Undo items contain mutation sequences for undo
and redo, associated with a document/unit. There are grouped undo scopes and
before/after command hooks.

Its sheets package has interceptors and distinct command/mutation/operation
modules. These provide extension points around an operation such as row deletion,
where tables, styles, references, and other features must all participate.

The README clearly separates open source plugins from Pro capabilities. In this
checkout it places collaboration, various advanced product surfaces, and several
import/export/reporting capabilities in the commercial layer. A hosted showcase
is not evidence that all those features are in the public repository.

## Where we stand comparatively

Spraedsheet already routes panels, REXX, control-bus calls, and MCP through
`commands.ts`. It has labels, parameters, journaling and grouped workbook undo.
`workbook.ts` and the view modules already separate model and presentation.
Univer validates this direction; it does not justify replacing AetherUI with
canvas or adopting its JavaScript UI framework.

The next opportunity is to tighten contracts around the command boundary as the
command collection grows. A script recording should capture what the user meant;
an undo entry must capture everything changed; cursor movement should be recorded
only where it is required to reproduce that meaning.

## Things to build independently

1. **Classify existing commands.** Distinguish durable edits, view changes,
   external effects, and queries. Keep the current `journaled`/`unstepped`
   behaviour explicit in the registry. Recording, permissions, help, and MCP
   discovery should consume this same information.
2. **Resolve context at execution.** If a toolbar action operates on the current
   selection, record its resolved sheet/range or record key. Replaying should
   not silently target whatever is selected later.
3. **Define feature participation in structural edits.** Establish a small
   lifecycle for validating an edit, transforming related metadata, calculating,
   committing history, and notifying views. Require new features to demonstrate
   their behaviour through insertion, deletion, copy, undo, and reopen.
4. **Define extension lifetimes.** A registered command, listener, or panel should
   be removable when its owning workbook or extension closes. Avoid retaining
   stale workbook state in global callbacks.

A mutation journal may eventually reduce snapshot costs, but a large plugin/DI
framework would be premature. Start with the registry and model services we have.

## First proof and priority

Record a command invoked from a panel, replay it over the bus with a different
selection, and confirm it affects the originally resolved target. Insert a row
through a named table with a chart, validation, and conditional format attached;
undo and reopen must restore all of them. A view-only action must follow its
declared persistence/history policy.

**Priority: medium**, alongside REXX recording and command hardening. Rendering
improvements that need native primitives belong in AetherUI/sae, with their own
tests, and are then consumed here.

## Local evidence

Under `/home/paul/scm/univer/`: `README.md` (including Open Source and Pro),
`packages/core/src/services/command/command.service.ts`,
`packages/core/src/services/undoredo/undoredo.service.ts`,
`packages/core/src/facade/`, `packages/sheets/src/commands/`, and
`packages/sheets/src/services/sheet-interceptor/`.

Spraedsheet counterparts: `app/sheet/{commands,workbook,described,menus}.ts`,
`app/sheet/view/`, `mcp/`. No third-party UI or command implementation is proposed
for reuse.
