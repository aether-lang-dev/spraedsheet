# Comparisons: actions by project

Each file here is a short action list, with the project's licence and what we
may take. Earlier prose is in git history. GridCraft is summary-only
(`docs/comparisons/gridcraft_feature_reference.md`); its repository is never read.

TODO.md is fully struck. Done 2026-10-24: IronCalc 4 and 5, Univer 1, HyperFormula 1, 4
and 5, Gnumeric 1 to 5 (1 in part), VisiData 2, LibreOffice 2 to 5 (4 and 6 in part). Left:
the column below; the record model (Grist, Airtable, VisiData 1 and 4, LibreOffice 7) waits
for Paul.

| Project | Licence / reuse | Top actions |
|---|---|---|
| [IronCalc](IRONCALC.md) | MIT OR Apache-2.0: copy or port with attribution (`third_party/ironcalc/PROVENANCE.md`) | the 48 calc_tests workbooks with more differences; small formatter ports where a test shows a gap |
| [Grist](GRIST.md) | Apache-2.0, ideas only | record identity; keyed JSON reconcile; reference fields |
| [Airtable](AIRTABLE.md) | proprietary, public docs ideas only | saved views and record form; trigger run log |
| [VisiData](VISIDATA.md) | GPL-3.0, ideas only | record by identity and deferred patches (the record model); lineage for group/pivot results (typed group keys first) |
| [Univer](UNIVER.md) | Apache-2.0, ideas only | a mutation journal (cuts snapshot cost; larger change); recording and MCP discovery to read `effect` |
| [HyperFormula](HYPERFORMULA.md) | GPLv3 or commercial, ideas only | range-aware aggregates (a rounding decision for Paul); coercion and array lifting declared in the function record |
| [Gnumeric](GNUMERIC.md) | GPL, ideas only | reference examples and boundary tests per function; `LOGEST` |
| [LibreOffice Calc](LIBREOFFICE.md) | MPL-2.0, ideas only | keyboard-only pass, print/layout, larger-grid interaction (aether-ui/sae); row lineage; the record model |

Gnumeric's website repository is covered in Gnumeric.md.
