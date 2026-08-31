<div align="center">

# 🤝 Contributing to KPR

### Engineering guidance for defensible Excel/VBA financial analytics and pricing

[![Conduct](https://img.shields.io/badge/Read_first-Code_of_conduct-6f42c1?style=flat-square)](CODE_OF_CONDUCT.md)
[![License](https://img.shields.io/badge/License-MIT-217346?style=flat-square)](LICENSE)
[![Source](https://img.shields.io/badge/Model-Source--first-0969da?style=flat-square)](#source-first-development)
[![Numerics](https://img.shields.io/badge/Numerics-Evidence_required-d97706?style=flat-square)](#numerical-engineering)
[![Office](https://img.shields.io/badge/Office-32_%2F_64--bit-217346?style=flat-square)](#vba-and-excel-standards)

<br>

**Explicit conventions · Stable numerics · Reviewable source · Reproducible evidence · Honest boundaries**

<br>

[Before you start](#before-you-start)
&nbsp;·&nbsp;
[Financial contracts](#financial-contract-first)
&nbsp;·&nbsp;
[Numerical evidence](#numerical-engineering)
&nbsp;·&nbsp;
[VBA standards](#vba-and-excel-standards)
&nbsp;·&nbsp;
[Validation](#validation-baseline)
&nbsp;·&nbsp;
[Pull requests](#pull-requests)

</div>

---

Thank you for helping improve **KPR**, an Excel/VBA library for financial
analytics and instrument pricing.

Contributions are welcome across code, tests, numerical reference sets,
documentation, examples, and repository infrastructure. Every contribution is
reviewed against the same question:

> Does this make the library more correct, explicit, reproducible, and safe for
> a caller to use?

By participating, you agree to follow the
[Code of Conduct](CODE_OF_CONDUCT.md). Security-sensitive reports should follow
[SECURITY.md](SECURITY.md) rather than being disclosed in a public issue.

---

<a id="before-you-start"></a>

## 🧭 Before you start

Read:

- [README.md](README.md)
- [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md)
- [SECURITY.md](SECURITY.md)
- [CHANGELOG.md](CHANGELOG.md)

For a defect, open an issue with a minimal reproduction. For a material feature,
new pricing model, public-API change, dependency, or architectural change, open
an issue before implementation so that scope and conventions can be agreed.

Small documentation corrections and similarly narrow changes may be submitted
directly in a focused pull request.

> [!IMPORTANT]
> KPR is not a generic VBA snippet collection. A result can be numerically close
> and still be financially wrong because of a date, quote, compounding, calendar,
> interpolation, settlement, or sign convention. Define the contract before
> optimizing the implementation.

### Engineering priorities

| Priority | Why it matters |
|---|---|
| 📐 **Explicit financial contract** | A formula without its conventions is not a reusable pricing function. |
| 🧮 **Numerical defensibility** | Plausible-looking output is not evidence of accuracy or stability. |
| 🧪 **Independent validation** | An implementation must not generate its own expected results. |
| 🎯 **Deterministic behavior** | The same supported inputs and environment should produce the same outcome. |
| 🔒 **Caller-owned Excel state** | A library must not silently take ownership of workbook or application state. |
| ⚙️ **32-bit and 64-bit compatibility** | Platform-specific defects remain defects even when they occur only on an untested host. |
| 🧱 **Stable public API** | Workbook formulas and VBA callers are expensive to migrate silently. |
| 📖 **Readable exported source** | The repository diff is the review artifact; the VBE is only an editing host. |
| 🔐 **Data and license integrity** | Financial workbooks and market data can carry confidential or restricted content. |
| 🧾 **Honest evidence** | Record what was actually tested and what remains unverified. |

---

## ⚡ Quick reference

| Before submitting | Minimum expectation |
|---|---|
| 🧭 **Contract** | Define inputs, outputs, conventions, domain, errors, and compatibility impact. |
| 🛠️ **Compile** | Import the exported source and run `Debug > Compile VBAProject`. |
| 🧪 **Test** | Run relevant regressions, edge cases, invariants, and independent comparisons. |
| 🎯 **Tolerances** | State absolute/relative rules and justify them from the contract. |
| 🖥️ **Environment** | Record Excel, Windows, Office bitness, locale, and date system where relevant. |
| 📚 **Documentation** | Update user-facing behavior and `[Unreleased]` notes in the same change. |
| 🔍 **Diff review** | Remove accidental binaries, generated files, secrets, and unrelated formatting. |

---

## 🎯 Project scope

KPR is intended to provide reusable financial calculations and instrument
pricing components that are:

- explicit about assumptions and market conventions;
- numerically defensible over a documented input domain;
- deterministic, testable, and diagnosable;
- usable from supported Excel environments on 32-bit and 64-bit Office;
- conservative with caller-owned workbook and application state; and
- documented well enough to use without reading the implementation.

A proposal may be declined when it falls outside that scope, duplicates an
existing capability, introduces disproportionate maintenance risk, or cannot be
validated independently.

---

<a id="source-first-development"></a>

## 📂 Source-first development

KPR uses exported, reviewable source files as its development record.

| Path | Purpose | Review status |
|---|---|---|
| `src/` | Exported VBA modules, classes, forms, and Ribbon source | Authoritative source |
| `test/` | Regression tests, numerical contracts, and validation support | Authoritative evidence code |
| `demo/` | Example and demonstration source | Reviewable examples |
| `dist/` | Release-artifact guidance and approved distributable outputs | Distribution boundary |
| `assets/`, `images/` | Documentation and repository media | Supporting material |

Every tracked VBA file conforms to the repository's Visual Basic Editor export
format and carries an `Attribute VB_Name` header that matches its file name.
Follow
[docs/VBE_EXPORT.md](docs/VBE_EXPORT.md) when exporting a component or importing
one back into Excel.

Keep form `.frm` and `.frx` companions together. Do not treat a binary workbook
as a substitute for exported source. Generated workbooks or add-ins should be
committed only when they are deliberate distribution or test artifacts and
their relationship to source is documented.

> [!CAUTION]
> Never normalize `.frx` files as text. Never use a binary workbook as the only
> record of a source change.

---

<a id="financial-contract-first"></a>

## 📐 Financial contract first

Every public financial function or pricing model must define its contract before
its implementation is reviewed.

The v0.0.2 date layer already has one. Its observable behaviour is frozen in
[docs/DATE_LAYER_CONTRACT.md](docs/DATE_LAYER_CONTRACT.md), which governs where
it and other documentation disagree about behaviour.

### Contract map

State whichever elements are relevant:

| Dimension | Questions the contract must answer |
|---|---|
| 🧾 **Instrument** | What cash flows, rights, obligations, or payoff are represented? |
| 📅 **Dates** | What are valuation, trade, settlement, payment, fixing, and maturity dates? |
| 💱 **Units** | Which currency, notional, price scale, rate scale, and output units apply? |
| 🗓️ **Calendars** | Which holidays, weekends, business-day rules, and end-of-month rules apply? |
| ⏱️ **Accrual** | Which day-count basis, frequency, stub, and accrual rules apply? |
| 📈 **Quotation** | Is the input/output a price, yield, spread, volatility, discount factor, or rate? |
| 🔁 **Compounding** | Which compounding convention and frequency apply? |
| ➕ **Signs** | Which payer/receiver, long/short, asset/liability, and cash-flow signs apply? |
| 📉 **Curves** | Which interpolation, extrapolation, bootstrap, and missing-data rules apply? |
| 🧮 **Output** | Is the result clean, dirty, accrued, present value, yield, sensitivity, or probability? |
| 🚧 **Domain** | Which inputs are valid, invalid, unsupported, or ambiguous? |
| ⚠️ **Failure** | How are invalid inputs, non-convergence, and unavailable results reported? |
| 🎯 **Accuracy** | Which precision, rounding, and tolerance contract applies? |

Do not hide a material convention behind an unexplained constant, a worksheet
format, or an Excel regional default. If the market admits multiple conventions,
require or document the selected convention explicitly.

### Compatibility-sensitive conventions

A change in any of the following may be a public behavior change even when the
VBA signature is unchanged:

```text
default settlement rule
day-count interpretation
cash-flow inclusion boundary
curve interpolation or extrapolation
quote or sign convention
rounding or convergence policy
invalid-input result
```

Assess and document the version impact.

---

<a id="numerical-engineering"></a>

## 🧮 Numerical engineering

Financial code must be supported by more than a few happy-path examples.

### 🔬 Independent reference evidence

Use an independent and attributable reference where practical:

- a published formula or technical paper;
- a regulator, central bank, exchange, or industry specification;
- a trusted analytical library;
- a separately implemented high-precision calculation;
- a licensed vendor result that may legally be used as evidence; or
- an independently derived invariant or limiting case.

Record the source, input conventions, precision, and comparison rule.

> [!WARNING]
> Do not use the implementation under test to generate its own expected values.
> If reference data is committed, make its provenance and generation process
> reproducible.

### 🧪 Test matrix

Tests should cover, as applicable:

| Class | Examples |
|---|---|
| ✅ **Ordinary** | Representative market inputs and documented examples |
| 0️⃣ **Degenerate** | Zero, near-zero, empty, coincident-date, and no-cash-flow cases |
| ↔️ **Boundary** | Leap years, month ends, cutoff dates, domain limits, and discontinuities |
| ➖ **Signed** | Negative rates, negative prices where admissible, long/short, payer/receiver |
| 🔭 **Extreme** | Large notionals, long maturities, deep tails, high/low volatility, stress inputs |
| 🚫 **Invalid** | Missing, malformed, inconsistent, non-finite, and out-of-domain inputs |
| 🔁 **Round trip** | Price → yield → price, discount factor → rate → discount factor |
| ⚖️ **Invariant** | Monotonicity, symmetry, parity, conservation, bounds, and limiting behavior |
| 📚 **Reference** | Independent values across the supported domain |
| 🐛 **Regression** | A permanent case for every corrected defect |

### 🎯 Tolerance contract

A numerical assertion should make its comparison policy explicit:

```text
absolute error
relative error
combined absolute/relative rule
units or basis-point interpretation
reference precision
domain over which the tolerance is claimed
```

Choose tolerances from the numerical contract and reference quality. Do not
relax a tolerance merely to make a test pass, and do not require exact binary
equality unless exactness is guaranteed by the algorithm and representation.

### 🛡️ Numerical robustness

Review, where relevant:

- cancellation and loss of significance;
- overflow, underflow, and non-finite intermediates;
- discontinuities and piecewise definitions;
- tail behavior and asymptotic approximations;
- iteration limits, initial guesses, bracketing, and convergence criteria;
- ill-conditioned inputs and sensitivity to perturbations; and
- reproducibility across VBA and worksheet-call paths.

A function must fail predictably when it cannot produce a valid result. It must
not silently return a plausible-looking number.

---

<a id="vba-and-excel-standards"></a>

## ⚙️ VBA and Excel standards

### Source rules

Contributed VBA source should:

- use `Option Explicit`;
- declare variables and parameters with the narrowest practical types;
- make `ByVal` and `ByRef` intent explicit;
- avoid `Select`, `Activate`, `ActiveWorkbook`, `ActiveSheet`, and implicit
  default members unless the public contract specifically requires them;
- qualify workbook, worksheet, range, and application references;
- keep public entry points small and move reusable logic into focused modules or
  classes;
- use comments to explain financial intent, numerical safeguards, and invariants
  rather than restating syntax;
- centralize shared constants and avoid unexplained numeric literals;
- handle errors deliberately and preserve useful diagnostic context;
- clean up transient state on both success and failure paths; and
- avoid new external references or dependencies unless discussed and accepted.

### Caller-owned state

Treat these as owned by Excel, the host workbook, or another component unless
KPR has explicitly established ownership:

```text
Application.Calculation
Application.EnableEvents
Application.ScreenUpdating
Application.DisplayAlerts
Application.StatusBar
active workbook / worksheet / selection
names, links, connections, and workbook structure
```

Restore only state that KPR successfully changed and can safely claim. Error
handling and cleanup must not conceal the original failure.

### Platform compatibility

Any WinAPI declaration must be compatible with supported 32-bit and 64-bit
Office configurations, normally through conditional compilation and
`LongPtr`-safe declarations.

Code must not depend unintentionally on locale-sensitive parsing, date formats,
decimal separators, worksheet names, the active selection, calculation mode, or
the caller's global Excel state.

Do not claim cross-bitness, cross-version, or cross-locale compatibility unless
the relevant environments were actually tested.

---

## 🧱 Public API and compatibility

Treat a documented public function, type, enum, parameter, return convention,
default, and error behavior as a compatibility contract.

A contribution that changes that contract must:

1. identify the affected callers;
2. explain the compatibility impact;
3. add or update contract tests;
4. update user-facing documentation and examples; and
5. state the expected Semantic Versioning impact.

Do not expose an internal helper as `Public` merely for convenience. Some Excel
or RibbonX callbacks may need public visibility; document those members as
infrastructure rather than supported consumer API.

### Semantic Versioning guide

| Impact | Typical change |
|---|---|
| 🩹 **Patch** | Backward-compatible defect correction, internal hardening, or documentation fix |
| ✨ **Minor** | New backward-compatible model, function, option, or supported convention |
| 💥 **Major** | Removed/renamed API, incompatible default, redefined convention, or changed result contract |

Version the supported behavior, not only the VBA signature.

---

## 🔄 Development workflow

1. Start from the current `main` branch.
2. Keep the change focused on one coherent purpose.
3. Preserve unrelated code and formatting.
4. Export all changed VBA components in reviewable form.
5. Compile the imported VBA project with `Debug > Compile VBAProject`.
6. Run the relevant regression tests and numerical comparisons.
7. Review the diff for accidental binary, generated, confidential, or
   environment-specific content.
8. Submit a pull request that links the relevant issue.

Use clear commits that describe the behavior changed. Avoid mixing refactors,
formatting, new functionality, and unrelated fixes when they can be reviewed
separately.

---

<a id="validation-baseline"></a>

## 🧪 Validation baseline

Until a task-specific automated harness is documented, the minimum validation
for a code change is:

- import the exported components into a clean test workbook or add-in;
- compile the complete VBA project;
- run every available relevant test;
- exercise the changed public path and its principal failure paths;
- compare numerical output with independent evidence where applicable;
- confirm behavior in a fresh Excel process; and
- record the environment, reference, tolerance, and results.

Changes involving platform declarations, Excel host behavior, date systems, or
locale-sensitive inputs should be tested on each affected configuration. If a
configuration cannot be tested, state that limitation explicitly.

### Performance evidence

Performance claims must include:

```text
representative workload
Excel / Windows / Office bitness
calculation and application-state setup
warm-up policy
timing method
sample size and dispersion
comparison baseline
```

Correctness and numerical stability outrank micro-optimization.

---

<a id="pull-requests"></a>

## 🚀 Pull requests

A reviewer should be able to answer:

```text
What problem does this solve?
What financial and API contract applies?
What behavior changes?
What behavior remains unchanged?
What independent evidence supports the result?
What remains unverified?
```

### Pull-request checklist

```text
[ ] Related issue linked for non-trivial work
[ ] Scope is focused
[ ] Financial conventions and input domain are explicit
[ ] Public API and Semantic Versioning impact assessed
[ ] Numerical method and stability risks assessed
[ ] Independent reference and tolerance recorded
[ ] Boundary, invalid-input, invariant, and regression cases covered
[ ] Caller-owned Excel state preserved
[ ] Error and non-convergence behavior tested
[ ] 32-bit / 64-bit impact assessed where relevant
[ ] Debug > Compile VBAProject passed
[ ] Relevant automated and manual results recorded
[ ] Performance evidence included for performance claims
[ ] Documentation and CHANGELOG updated
[ ] No confidential, restricted, generated, or accidental binary content added
```

### Suggested evidence block

```text
Environment
-----------
KPR commit:
Excel:
Office bitness:
Windows:
Locale / date system:
Deployment: embedded workbook / add-in

Financial contract
------------------
Instrument / function:
Conventions:
Input domain:
Expected failure behavior:

Numerical evidence
------------------
Independent reference:
Reference precision:
Tolerance rule:
Worst observed error:

Validation
----------
Compile:
Regression tests:
Manual checks:

Known boundary
--------------
<what this evidence does not prove>
```

Record only tests and environments actually completed.

---

## 📖 Documentation and changelog

Documentation belongs in the same change as the behavior it describes.

| Change | Documentation impact |
|---|---|
| Public function or model | API, conventions, inputs, outputs, examples, and changelog |
| Numerical method | Method, domain, reference, tolerance, and known limitations |
| Default or convention | Compatibility note, migration impact, tests, and changelog |
| Error behavior | Public contract, examples, and regression evidence |
| Installation or packaging | README/install guidance and release notes |
| Validation infrastructure | Contributor guidance and evidence instructions |

Add user-visible changes to the `[Unreleased]` section of
[CHANGELOG.md](CHANGELOG.md). Do not record unverified claims or validation that
has not actually been run.

---

## 🔐 Data, privacy, and provenance

Use synthetic or properly licensed test data. Do not commit:

- client, employer, counterparty, student, or personal data;
- credentials, connection strings, signing material, or internal URLs;
- proprietary models, spreadsheets, market data, or vendor outputs that cannot
  legally be redistributed;
- copied code whose origin or license cannot be established; or
- generated content represented as independently validated evidence.

Contributors remain responsible for reviewing, testing, and licensing any
AI-assisted material they submit. AI assistance does not replace numerical
evidence, provenance, or authorship responsibility.

---

## 🤝 Review culture

A strong review comment states the location, contract, risk, evidence, and
whether the change is required or optional.

Useful:

> `BondPrice`: the new branch excludes a cash flow on the settlement date, but
> the public contract does not define that boundary. Please state the inclusion
> rule, add both sides of the boundary to the tests, and compare them with the
> independent reference.

Less useful:

> The bond calculation is wrong.

Review the software precisely and the contributor respectfully. See the
[Code of Conduct](CODE_OF_CONDUCT.md).

---

## 📄 License

By contributing, you agree that your contribution will be licensed under KPR's
[MIT License](LICENSE). Identify any adapted source and its license in the pull
request.

---

<div align="center">

### Contribution principle

**Define the convention · Preserve caller state · Prove the number · Test the boundary · State what remains unknown**

<br>

Maintained by **Daniele Penza**

</div>
