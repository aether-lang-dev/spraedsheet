# Gnumeric and gnumeric-web compared with Spraedsheet

**Most useful lesson: numerical accuracy and evidence of correctness are product
features.** Having a statistical function or Solver command is the beginning of
support, not the end of validation.

Review: `/home/paul/scm/gnumeric` at `e00c927` and
`/home/paul/scm/gnumeric-web` at `7567afd`. Baseline and limits:
[comparison index](../../OPEN_SOURCE_COMPARISONS.md).

## The two repositories

`gnumeric` is the C/GTK spreadsheet application. It contains expressions,
dependency handling, numerical functions, analysis tools, conversion tools,
plugins, a manual, and tests. Parts of its README explicitly label their old
feature descriptions as outdated, so source organisation and focused tests are
more useful than treating that list as a current feature matrix.

`gnumeric-web` is the website: HTML pages, announcements, function-status tables,
screenshots, and technical explanations. It is not a WebAssembly or JavaScript
spreadsheet implementation. Its numerical notes and status reporting provide a
second perspective on how Gnumeric presents its engineering choices.

## What is worth learning

The application has separate metadata for a function's implementation and testing
status. The website's `functions.shtml` exposes those two dimensions. This is a
better signal than claiming compatibility solely because a function name exists.
Treat website counts as claims in that checkout, not independently verified
coverage measurements.

`src/README.gnm_float` discusses maintaining precision and choosing operations
that behave better near numerically difficult inputs. The test directory contains
function workbooks, recalculation checks, statistical-analysis checks, and named
nonlinear-solver cases. `src/dependent.h` explicitly supports volatile and dynamic
dependencies. `ssconvert` provides a conversion-oriented application surface.

Website material is still documentation, not a mathematical oracle. Numerical
claims and expected values should be checked against independent references and
the actual configured precision before we use them as acceptance criteria.

## Where we stand comparatively

Spraedsheet already implements distributions, financial functions, regression,
analysis reports, Goal Seek, and a linear/integer Solver. Nonlinear/evolutionary
Solver work remains in TODO 6.8b. It also has native conversion/import paths and
a growing function catalogue. The opportunity is stronger numerical testing,
diagnostics, and public maturity reporting rather than simply increasing counts.

## Things to build independently

1. **Function maturity metadata.** Label implemented behaviour, known deviations,
   reference examples, boundary tests, and numerical stress tests separately.
   Generate help/status documentation from the app's own function definitions.
2. **A numerical acceptance suite.** Test distributions in extreme tails, inverse
   functions near boundaries, nearly constant datasets, ill-conditioned
   regressions, financial roots, convergence failure, and large/small magnitudes.
   Use independent published reference datasets or high-precision expected values;
   write original fixtures and check provenance.
3. **Declared tolerances and error policy.** Use exact checks for counts and
   identities, and justified absolute/relative error bounds for approximations.
   Specify NaN/infinity, domain errors, singular inputs, and non-convergence.
   Excel-compatible quirks and mathematical accuracy are distinct goals.
4. **Analysis reports that explain themselves.** Show method, inputs, assumptions,
   convergence status, and useful diagnostics. A numeric result should not mask a
   failed fit or an infeasible solve.
5. **Headless inspection/conversion workflows.** Extend the planned CLI over the
   existing command interface with machine-readable failures and loss reports.

## First proof and priority

For each selected function family, pair ordinary examples with difficult cases
and expected error cases. Check invariants such as inverse/distribution round
trips, but also compare with independently derived expected values: two mutually
wrong functions can round-trip. Exercise both formula modes where the same
operation is offered, keeping language/coercion differences explicit.

**Priority: high for numerical hardening.** Implementing more Solver algorithms
should follow a defined accuracy/convergence contract. Aether's number semantics
and sae lowering must be covered where they affect results.

## Local evidence

Under `gnumeric/`: `src/{func.h,dependent.h,mathfunc.c,ssconvert.c,sstest.c}`,
`src/README.gnm_float`, `src/tools/`, `plugins/`, `test/t1000-statfuns.pl`,
`test/t1001-mathfuns.pl`, `test/t2005-recalc.pl`,
`test/t7203-descriptive-statistics.pl`, and the `test/t71*-solver-*` cases.

Under `gnumeric-web/`: `index.html`, `functions.shtml`,
`numerical-issues.html`, and `date-time-formats.shtml`.

Spraedsheet counterparts: `app/sheet/{distributions,finance,bonds,analysis,solver,dates,described}.ts`.
No Gnumeric numerical routine, test workbook, or website text is copied.
