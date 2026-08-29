# Contributing to KPR

Thank you for helping improve KPR, an Excel/VBA library for financial analytics
and instrument pricing. Contributions are welcome across code, tests,
documentation, numerical references, examples, and repository infrastructure.

By participating, you agree to follow the
[Code of Conduct](CODE_OF_CONDUCT.md). Security-sensitive reports should follow
[SECURITY.md](SECURITY.md) rather than being disclosed in a public issue.

## Before you start

For a defect, open an issue with a minimal reproduction. For a material feature,
new pricing model, public-API change, dependency, or architectural change, open
an issue before implementation so that scope and conventions can be agreed.

Small documentation corrections and similarly narrow changes may be submitted
directly in a focused pull request.

## Project scope

KPR is intended to provide reusable financial calculations and instrument
pricing components that are:

- explicit about financial conventions and assumptions;
- numerically defensible;
- deterministic and testable;
- usable from supported versions of Excel on 32-bit and 64-bit Office;
- conservative with caller-owned workbook and application state; and
- documented well enough to be used without reading the implementation.

A proposal may be declined when it falls outside that scope, duplicates an
existing capability, introduces disproportionate maintenance risk, or cannot be
validated independently.

## Repository model

KPR uses exported, reviewable source files as its development record.

| Path | Purpose |
|---|---|
| `src/` | Exported VBA modules, classes, forms, and Ribbon source |
| `test/` | Regression tests, numerical contracts, and validation support |
| `demo/` | Example and demonstration source |
| `dist/` | Release-artifact guidance and approved distributable outputs |
| `assets/`, `images/` | Documentation and repository media |

Keep form `.frm` and `.frx` companions together. Do not treat a binary workbook
as a substitute for exported source. Generated workbooks or add-ins should be
committed only when they are deliberate distribution or test artifacts and
their relationship to source is documented.

## Financial contract first

Every public financial function or pricing model must define its contract before
its implementation is reviewed. State whichever of the following are relevant:

- instrument and cash-flow definition;
- valuation and settlement dates;
- currency and units;
- day-count basis and accrual rules;
- business-day convention and holiday calendar;
- compounding, frequency, and rate quotation;
- clean, dirty, present-value, or yield output convention;
- payer/receiver, long/short, and cash-flow sign conventions;
- curve, interpolation, extrapolation, and missing-data behavior;
- treatment of past, same-day, and ex-coupon cash flows;
- admissible input domain and invalid-input behavior; and
- accuracy, rounding, and convergence expectations.

Do not hide a material convention behind an unexplained constant or an Excel
regional default. If the market admits multiple conventions, require or document
the selected convention explicitly.

## Numerical engineering requirements

Financial code must be supported by more than a few happy-path examples.

### Reference evidence

Use an independent and attributable reference where practical: a published
formula, regulator or exchange specification, trusted analytical library,
vendor output, or separately implemented high-precision calculation. Record the
source, input conventions, precision, and comparison tolerance.

Do not use the implementation under test to generate its own expected values.
If independently generated reference data is committed, make its provenance and
generation method reproducible.

### Test coverage

Tests should cover, as applicable:

- ordinary market examples;
- zero, near-zero, negative, and extreme inputs;
- boundary dates and leap years;
- invalid domains and missing inputs;
- convergence limits and non-convergence behavior;
- monotonicity, symmetry, parity, conservation, or other model invariants;
- round trips such as price-to-yield-to-price;
- comparison with independent reference values; and
- regression cases for every corrected defect.

Choose tolerances from the numerical contract and reference precision. Do not
relax a tolerance merely to make a test pass, and do not rely on exact binary
equality for floating-point results unless exactness is itself guaranteed.

### Numerical robustness

Prefer algorithms that remain stable over the documented input domain. Review
cancellation, overflow, underflow, loss of significance, discontinuities, tail
behavior, iteration limits, and initial guesses where relevant. A function must
fail predictably when it cannot produce a valid result; it must not silently
return a plausible-looking number.

## VBA and Excel standards

Contributed VBA source should:

- use `Option Explicit`;
- declare variables and parameters with the narrowest practical types;
- make `ByVal` and `ByRef` intent explicit;
- avoid `Select`, `Activate`, `ActiveWorkbook`, `ActiveSheet`, and implicit
  default members unless the public contract specifically requires them;
- qualify workbook, worksheet, range, and application references;
- keep public entry points small and move reusable logic into focused modules or
  classes;
- use descriptive names and comments that explain financial or numerical intent
  rather than restating the code;
- centralize shared constants and avoid unexplained numeric literals;
- handle errors deliberately and preserve useful diagnostic context;
- restore any Excel application state it changes, but only when ownership of
  that change is established;
- clean up objects and transient state on both success and failure paths; and
- avoid new external references or dependencies unless they have been discussed
  and accepted.

Any WinAPI declaration must be compatible with supported 32-bit and 64-bit
Office configurations, normally through conditional compilation and
`LongPtr`-safe declarations.

Code must not depend unintentionally on locale-sensitive parsing, date formats,
decimal separators, worksheet names, the active selection, calculation mode, or
the caller's global Excel state.

## Public API and compatibility

Treat a documented public function, type, enum, parameter, return convention,
default, and error behavior as a compatibility contract.

A contribution that changes that contract must:

1. identify the affected callers;
2. explain the compatibility impact;
3. add or update contract tests;
4. update user-facing documentation and examples; and
5. state the expected Semantic Versioning impact.

Do not expose an internal helper as `Public` merely for convenience. Some VBA
callbacks may need public visibility for Excel or RibbonX; document such members
as infrastructure rather than supported consumer API.

## Workflow

1. Start from the current `main` branch.
2. Keep the change focused on one coherent purpose.
3. Preserve unrelated code and formatting.
4. Export all changed VBA components in a reviewable form.
5. Compile the imported VBA project with `Debug > Compile VBAProject`.
6. Run the relevant regression tests and numerical comparisons.
7. Review the diff for accidental binary, generated, confidential, or
   environment-specific content.
8. Submit a pull request that links the relevant issue.

Use clear commits that describe the behavior changed. Avoid mixing refactors,
formatting, new functionality, and unrelated fixes when they can be reviewed
separately.

## Pull-request evidence

A pull request should explain:

- the problem or capability;
- the intended financial and API contract;
- the implementation approach and alternatives considered;
- compatibility and performance implications;
- tests added or changed;
- the Excel, Windows, and Office bitness used for validation;
- the exact validation result, including tolerances for numerical comparisons;
  and
- documentation or release-note changes.

Screenshots can support UI or workbook-output changes, but they do not replace
source tests or numerical evidence.

## Validation baseline

Until a task-specific automated harness is documented, the minimum validation
for a code change is:

- import the exported components into a clean test workbook or add-in;
- compile the complete VBA project;
- run every available relevant test;
- exercise the changed public path and its principal failure paths;
- confirm behavior in a fresh Excel process; and
- record the environment and results in the pull request.

Changes involving platform declarations, Excel host behavior, date systems, or
locale-sensitive inputs should be tested on each affected configuration. If a
configuration cannot be tested, state that limitation explicitly.

Performance claims must include a reproducible workload, environment, warm-up
policy, timing method, sample size, and comparison baseline.

## Documentation and changelog

Update documentation in the same change whenever behavior, supported inputs,
outputs, defaults, error handling, installation, or compatibility changes.

Add user-visible changes to the `[Unreleased]` section of
[CHANGELOG.md](CHANGELOG.md). Do not record unverified claims or validation that
has not actually been run.

## Data, privacy, and provenance

Use synthetic or properly licensed test data. Do not commit:

- client, employer, counterparty, student, or personal data;
- credentials, connection strings, signing material, or internal URLs;
- proprietary models, spreadsheets, market data, or vendor outputs that cannot
  legally be redistributed; or
- copied code or generated material whose origin and license cannot be
  established.

Contributors remain responsible for reviewing, testing, and licensing any
AI-assisted material they submit. AI assistance does not replace numerical
evidence or authorship responsibility.

## License

By contributing, you agree that your contribution will be licensed under KPR's
[MIT License](LICENSE). Identify any adapted source and its license in the pull
request.

KPR is maintained by **Daniele Penza**.
