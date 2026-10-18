# IronCalc material in this repository

Upstream: <https://github.com/ironcalc/IronCalc>
Commit: `297553dc3a37ec6b0f4821055220e3d26c4ac6a1`
Licence: IronCalc is MIT OR Apache-2.0 at the recipient's option; we take it
under **MIT**. Its `LICENSE-MIT` (copyright EqualTo GmbH and Nicolas Hatcher,
with the full permission and warranty notice) is [LICENSE-MIT](LICENSE-MIT)
here, unmodified, and applies to everything listed below.

Why it is here: Paul approved copying and porting selected IronCalc material
into Spraedsheet under that licence (an exception to the usual not-ours rule,
for IronCalc only). The plan is `docs/comparisons/IRONCALC.md`.

Rules this file keeps:

- Every upstream path copied, its local destination, and what was adapted is
  recorded here, in the commit that brings it in.
- Upstream workbooks are copied byte for byte into `third_party/ironcalc/`
  (never mixed with our own cases); everything we translated lives in our own
  test files, each case naming the upstream test it came from.
- Upstream expected values are not an oracle. Where ours differ, Excel's
  documented behaviour decides, and an intentional difference is recorded
  under "Deliberate differences" below.
- No workbook is assumed to have the provenance of the repository as a whole;
  each fixture's own notes are under its entry.

## Copied and adapted material

| Upstream path | Local destination | What was adapted |
|---|---|---|
| `LICENSE-MIT` | `third_party/ironcalc/LICENSE-MIT` | nothing |

## Deliberate differences

None recorded yet.
