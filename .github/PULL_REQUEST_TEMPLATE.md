<!--
  KPR pull request template

  Keep the PR focused on one coherent purpose. Complete the core sections and
  delete optional sections that do not apply; do not convert unperformed work
  into "PASS" or fill the template with unexplained "N/A" entries.

  Record only tests, references, environments, and compatibility claims that
  were actually verified. Evidence matters more than checkbox volume.

  SECURITY: do not disclose a suspected vulnerability, credential, confidential
  workbook, restricted market data, or exploitable trust-boundary issue here.
  Follow SECURITY.md and report privately.
-->

<div align="center">

# 🔀 KPR Pull Request

### Focused change · Explicit contract · Reproducible evidence · Honest boundaries

[![Contract](https://img.shields.io/badge/contract-explicit-217346?style=flat-square)](../CONTRIBUTING.md#financial-contract-first)
[![Numerics](https://img.shields.io/badge/numerics-evidence_required-d97706?style=flat-square)](../CONTRIBUTING.md#numerical-engineering)
[![Compatibility](https://img.shields.io/badge/API-SemVer_assessed-6f42c1?style=flat-square)](../CONTRIBUTING.md)
[![Data](https://img.shields.io/badge/data-synthetic_only-d73a49?style=flat-square)](../SECURITY.md#financial-data-and-confidentiality)

</div>

---

## 📌 Summary

<!-- What changed and why? Prefer one precise outcome over a list of edits. -->

## 🎯 Problem or motivation

<!--
State the defect, limitation, analytical need, maintenance problem, or user
workflow. Describe the pre-change behavior where useful.
-->

## 🔗 Related issue

```text
Closes #
```

<!-- Use "Refs #" when the PR contributes to but does not close the issue. -->

---

## 🏷️ Type of change

- [ ] 🐛 Defect correction
- [ ] ✨ Backward-compatible feature or model
- [ ] 💥 Breaking contract or API change
- [ ] 🧮 Numerical-method or accuracy change
- [ ] 📅 Date, calendar, day-count, or schedule change
- [ ] 📈 Rate, curve, cash-flow, or instrument change
- [ ] 🔒 Excel-state, data-integrity, or security hardening
- [ ] ♻️ Internal refactor with no intended supported-behavior change
- [ ] 🧪 Test, reference-data, or validation change
- [ ] ⚡ Performance change
- [ ] 📦 Packaging, release, or repository-infrastructure change
- [ ] 📖 Documentation-only change

## 🎚️ Affected layer

- [ ] 📊 Excel UDF or VBA caller surface
- [ ] 🧱 Supported public API or result contract
- [ ] 🧮 Pricing or analytics engine
- [ ] 🧾 Cash-flow or instrument representation
- [ ] 📉 Curve, interpolation, extrapolation, or calibration
- [ ] 📅 Date, calendar, convention, or schedule foundation
- [ ] 🔢 Numerical primitive, solver, or approximation
- [ ] ⚙️ Excel host integration or application state
- [ ] 🧪 Regression harness or independent reference set
- [ ] 📦 Demo, workbook/add-in artifact, or release process
- [ ] 📚 Documentation or repository governance

---

## 📐 Contract snapshot

Complete the lines relevant to this PR. A reviewer should understand the
behavioral boundary before reading the implementation.

```text
Supported function / component:
Inputs and admissible domain:
Output and units:
Financial conventions:
Defaults:
Invalid-input behavior:
Non-convergence / unavailable-result behavior:
Caller-owned Excel state preserved:
Known limitation introduced or retained:
```

### Financial conventions

Check and explain every convention this change affects.

- [ ] 📅 Valuation, trade, settlement, fixing, payment, or maturity date
- [ ] 🗓️ Calendar, weekend, roll, stub, or end-of-month rule
- [ ] ⏱️ Day count, accrual, frequency, or compounding
- [ ] 💱 Currency, notional, units, or scaling
- [ ] 📈 Price, yield, rate, spread, volatility, or probability quotation
- [ ] ➕ Long/short, payer/receiver, asset/liability, or cash-flow sign
- [ ] 📉 Curve input, interpolation, extrapolation, or missing-data rule
- [ ] 🧮 Precision, rounding, tolerance, iteration, or convergence
- [ ] None — no financial convention is affected

```text
Convention detail:
```

> [!IMPORTANT]
> A formula can be numerically close and still be financially wrong. Do not
> leave a material convention implicit in workbook formatting, locale, a magic
> constant, or the implementation itself.

---

## 🧱 Public API and Semantic Versioning

```text
Supported behavior changed:     Yes / No
Backward compatible:           Yes / No / Uncertain
Suggested release impact:      patch / minor / major / none / uncertain
New supported members:
Removed or renamed members:
Changed defaults:
Changed result or error contract:
Migration required:
```

Assess compatibility against documented behavior, not only the VBA keyword
`Public`. Infrastructure members may need public visibility for Excel, RibbonX,
callbacks, `Application.Run`, packaging, or tests without becoming supported API.

If no supported behavior changes, state:

```text
No supported behavior change.
```

---

## 🛠️ Implementation approach

<!--
Explain the design, important alternatives, invariants, and why this approach was
chosen. Focus on decisions a future maintainer cannot infer safely from the diff.
-->

```text
Approach:
Alternatives considered:
Key invariant:
New dependency or reference:
```

<details>
<summary><strong>🧮 Numerical-method detail</strong></summary>

<!-- Complete for algorithms, solvers, approximations, curves, or accuracy changes. -->

```text
Method / equation:
Source or derivation:
Supported domain:
Stability risks:
Initial guess / bracket:
Iteration limit:
Convergence criterion:
Tail / boundary treatment:
Failure behavior:
```

</details>

<details>
<summary><strong>⚙️ Excel host-state detail</strong></summary>

<!-- Complete when the change reads or mutates Excel objects or application state. -->

```text
Workbook / worksheet / range scope:
Application state read:
Application state changed:
Ownership established how:
Success-path cleanup:
Failure-path cleanup:
Formula / name / link / connection impact:
```

</details>

---

## ✅ Validation

### Environment

Record environments actually used. Do not claim compatibility for an untested
configuration.

```text
KPR base commit:
Excel version / build:
Office bitness:             32-bit / 64-bit
Windows version:
Locale / decimal separator:
Workbook date system:       1900 / 1904
Deployment:                 source workbook / add-in / other
```

### Validation performed

- [ ] 🛠️ `Debug > Compile VBAProject`
- [ ] 🔬 Focused unit or component tests
- [ ] 🐛 Complete relevant regression pack
- [ ] 📚 Independent reference comparison
- [ ] ⚖️ Invariant, parity, monotonicity, symmetry, or round-trip tests
- [ ] ↔️ Boundary and limiting cases
- [ ] 🚫 Invalid-input and failure-path tests
- [ ] 🔁 Non-convergence or recovery tests
- [ ] 📊 Worksheet UDF and VBA caller-path comparison
- [ ] 🆕 Fresh Excel-process check
- [ ] ⚙️ 32-bit / 64-bit check where relevant
- [ ] 🌍 Locale or date-system check where relevant
- [ ] 📦 Packaged-artifact check where relevant

```text
Compile result:
Regression result:
Focused/manual result:
```

> [!WARNING]
> Record only validation that was actually run. An operational GitHub workflow,
> a screenshot, or a successful compile is not numerical-reference evidence.

### Numerical evidence

Complete for changes that produce or affect numerical output.

```text
Independent reference:
Reference provenance / version:
Reference precision:
Tested domain:
Comparison rule:            absolute / relative / combined / other
Tolerance:
Worst observed error:
Worst-error input:
Unverified boundary:
```

### Test cases added or changed

| Test | Contract or defect covered | Expected result |
|---|---|---|
| | | |

---

## ⚡ Performance

```text
Performance impact:         improved / neutral / regressed / not measured
Representative workload:
Timing method:
Warm-up policy:
Sample size:
Baseline:
Observed result:
```

Do not make a performance claim without a reproducible workload and measurement
boundary. Correctness and stability outrank micro-optimization.

---

## 🔐 Security, data, and provenance

- [ ] No client, employer, counterparty, student, or personal data is included
- [ ] No credential, connection string, internal URL, or signing material is included
- [ ] Test inputs are synthetic or legally redistributable
- [ ] Market data and vendor outputs are licensed for the way they are used
- [ ] Adapted algorithms, code, and datasets identify their source and license
- [ ] No unexpected formula, command, path, or external-content injection surface is introduced
- [ ] No security-sensitive detail requires private disclosure instead of this PR

```text
Source / dataset provenance:
Security impact:
```

---

## 📖 Documentation and release impact

- [ ] README updated
- [ ] Installation guidance updated
- [ ] API / model documentation updated
- [ ] Examples or demo updated
- [ ] `[Unreleased]` changelog updated
- [ ] Security policy updated
- [ ] No documentation change required — reason stated below

```text
Documentation impact:
Release-artifact impact:
```

---

## ⚠️ Known boundaries

<!-- State what this PR and its evidence do not prove. Do not leave implicit. -->

```text
Untested environment:
Unsupported domain:
Deferred follow-up:
Residual risk:
```

---

## 📋 Author checklist

- [ ] The PR has one coherent purpose
- [ ] The related issue is linked
- [ ] The financial and API contracts are explicit
- [ ] Public API and Semantic Versioning impact are assessed
- [ ] Numerical evidence is independent and attributable where required
- [ ] Tolerances follow the contract rather than the observed implementation
- [ ] Caller-owned Excel state and failure cleanup are preserved
- [ ] A regression case covers each corrected defect
- [ ] Documentation reflects current behavior, not intended future behavior
- [ ] Evidence records only tests and environments actually completed
- [ ] Confidential, restricted, generated, and binary content was reviewed
- [ ] Known limitations and unverified boundaries are stated

---

## 👀 Reviewer focus

<!-- Point reviewers to the highest-risk files, contracts, or numerical decisions. -->

```text
Primary review risk:
Files / procedures to inspect first:
Evidence to challenge:
```
