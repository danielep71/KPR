---
name: ✨ Feature or model proposal
about: Propose a KPR financial model, analytical capability, convention, API, validation, or deployment improvement
title: "[Feature]: "
labels: enhancement
---

<!--
  KPR feature or model proposal

  Begin with the real workflow and desired behavior. You do not need to design
  the implementation before the contract is clear.

  Use the bug template when current documented behavior is reproducibly wrong.
  Follow SECURITY.md privately for vulnerabilities, credentials, confidential
  workbooks, malicious artifacts, or exploitable trust-boundary issues.
-->

<div align="center">

# ✨ KPR Feature or Model Proposal

[![Use case](https://img.shields.io/badge/Start_with-Real_use_case-0969da?style=flat-square)](#problem-and-use-case)
[![Contract](https://img.shields.io/badge/Model-Contract-explicit-217346?style=flat-square)](#financial-contract)
[![Evidence](https://img.shields.io/badge/Validation-Independent-d97706?style=flat-square)](#validation-strategy)
[![API](https://img.shields.io/badge/API-Compatibility_assessed-6f42c1?style=flat-square)](#api-and-compatibility)

</div>

<a id="problem-and-use-case"></a>

## 🎯 Problem and real use case

Describe the workflow, analytical need, limitation, or recurring friction.

```text
Primary user:
Typical workflow:
Frequency of need:
Current limitation:
Practical consequence:
```

<!--
Prefer:
"I need bond accrued interest and clean/dirty price to use an explicit ex-coupon
rule so the same contract can be used from worksheet formulas and VBA."

Avoid:
"Add bond functions."
-->

## 💡 Desired behavior

Describe the outcome before prescribing the implementation.

```text
Requested capability:
Expected caller surface:
Expected default:
Expected output:
Expected invalid-input behavior:
Expected non-convergence / unavailable-result behavior:
```

## 👤 Who benefits?

- [ ] Financial analyst or risk practitioner
- [ ] Workbook author
- [ ] VBA developer
- [ ] Model validator or reviewer
- [ ] Academic or teaching user
- [ ] Add-in or controlled-deployment user
- [ ] Maintainer, tester, or release engineer
- [ ] Other

---

## 🏷️ Proposal category

- [ ] 📅 Date, calendar, roll, day-count, or schedule capability
- [ ] 📈 Rate, discounting, compounding, or time-value capability
- [ ] 📉 Curve, interpolation, extrapolation, or calibration capability
- [ ] 🧾 Cash-flow or instrument representation
- [ ] 💰 Pricing model or valuation function
- [ ] 🎯 Risk measure, sensitivity, scenario, or aggregation
- [ ] 🔢 Numerical method, solver, approximation, or accuracy improvement
- [ ] 📊 Worksheet UDF or VBA public API
- [ ] ⚙️ Excel integration or caller-state behavior
- [ ] 🧪 Test harness, reference data, or validation tooling
- [ ] ⚡ Performance or scalability improvement
- [ ] 📦 Installation, packaging, release, or provenance improvement
- [ ] 📖 Documentation or example
- [ ] Other

---

<a id="financial-contract"></a>

## 📐 Financial contract

Complete this section for a calculation, convention, model, instrument, or
analytical feature. Delete fields that genuinely do not apply.

```text
Instrument / calculation:
Inputs and admissible domain:
Valuation and settlement dates:
Currency / notional / units:
Calendar and business-day rule:
Day-count / accrual rule:
Frequency / stub / end-of-month rule:
Compounding convention:
Quote type and scaling:
Long/short or payer/receiver sign:
Curve inputs and interpolation/extrapolation:
Cash-flow inclusion boundary:
Output definition and units:
Rounding / precision:
```

### Market variants

Which alternative conventions exist, and which should KPR support?

| Variant | Required now | Possible later | Explicitly out of scope |
|---|:---:|:---:|:---:|
| | | | |

> [!IMPORTANT]
> Do not ask the implementation to infer a material market convention from
> workbook formatting, locale, or an unexplained default.

---

## ✅ Acceptance criteria

Use observable outcomes rather than internal helper names.

1.
2.
3.

```text
Accuracy / tolerance criterion:
Compatibility criterion:
Failure criterion:
Documentation criterion:
```

## 🚫 Non-goals

State what this proposal should not attempt to solve.

```text

```

This keeps a useful capability from silently becoming an architectural rewrite.

---

## 🔀 Current workaround and alternatives

```text
Current workaround:
Why it is insufficient:
Alternative implementation or library considered:
Reason KPR should own this capability:
```

- [ ] Existing KPR primitives can express part of the behavior
- [ ] A workbook formula is used today
- [ ] A wrapper VBA procedure is used today
- [ ] A third-party library or vendor tool is used today
- [ ] The result is calculated manually
- [ ] No practical workaround exists

---

## 🧮 Method and reference basis

Complete for algorithms, models, solvers, curves, or numerical changes.

```text
Proposed method / equation:
Primary specification or reference:
Reference link / citation / version:
Why this method is appropriate:
Known alternatives:
Supported domain:
Stability or conditioning risks:
Tail / boundary behavior:
Convergence approach:
```

### Reference availability

- [ ] Published formula or paper
- [ ] Regulator, central bank, exchange, or industry specification
- [ ] Independent trusted analytical library
- [ ] Separately implemented high-precision reference
- [ ] Licensed vendor output that may be used as evidence
- [ ] Independently derived invariant or limiting case
- [ ] No independent reference identified yet

Describe any licensing or redistribution limits on reference data.

---

<a id="api-and-compatibility"></a>

## 🧱 API and compatibility

An exact final signature is not required, but the intended caller contract should
be clear.

```text
Worksheet UDF, VBA API, or both:
New supported member proposed:
Existing supported member changed:
Parameter and type concept:
Return / result concept:
Error or unavailable-result concept:
Backward compatible:              Yes / No / Uncertain
Migration required:
Suggested release impact:         patch / minor / major / uncertain
```

### Illustrative use

```vb
' Illustrative only — the reviewed public API may differ.
```

```excel
=IllustrativeFormula(...)
```

Delete either block if it does not apply.

> [!NOTE]
> A VBA member may need public visibility for Excel, RibbonX, callbacks,
> `Application.Run`, packaging, or tests without becoming supported consumer API.

---

<a id="validation-strategy"></a>

## 🧪 Validation strategy

How can the capability be proved correct independently?

### Proposed evidence

```text
Independent reference:
Reference precision:
Comparison rule:              absolute / relative / combined / other
Tolerance rationale:
Tested domain:
Expected worst-case region:
```

### Required test classes

- [ ] ✅ Representative market examples
- [ ] 0️⃣ Zero, near-zero, empty, or degenerate cases
- [ ] ↔️ Date, domain, discontinuity, and limiting boundaries
- [ ] ➖ Negative rates, prices, signs, or notionals where admissible
- [ ] 🔭 Extreme maturities, notionals, rates, volatilities, or tails
- [ ] 🚫 Invalid, inconsistent, missing, and non-finite inputs
- [ ] 🔁 Price/yield, rate/discount-factor, or other round trips
- [ ] ⚖️ Bounds, parity, monotonicity, symmetry, or conservation
- [ ] 📚 Independent reference values across the supported domain
- [ ] 🐛 Permanent regressions for defects found during implementation
- [ ] 📊 Worksheet and VBA caller-path equivalence

### Evidence boundary

```text
Environment(s) required:
Data or tool required:
What the proposed evidence would not prove:
```

---

## ⚙️ Excel and deployment behavior

Complete when the feature touches Excel objects, global application state, or
distribution.

```text
Workbook / worksheet / range scope:
Formula / name / link / connection impact:
Application state read or changed:
32-bit / 64-bit impact:
Locale or date-system impact:
External reference / dependency:
Embedded-source behavior:
Add-in behavior:
```

KPR must preserve caller-owned state unless the public contract establishes a
specific, bounded mutation.

---

## ⚡ Performance expectations

```text
Representative workload:
Expected scale:
Acceptable latency:
Memory / workbook-size concern:
Measurement approach:
```

Do not trade away correctness, stability, or contract clarity for an unmeasured
micro-optimization.

---

## 🔐 Security, data, and licensing

- [ ] The proposal can be demonstrated with synthetic or redistributable data
- [ ] No client, employer, counterparty, student, or personal data is required
- [ ] Market data or vendor references can be used within their license
- [ ] No credential, secret, signing material, or internal endpoint is required
- [ ] Formula, command, path, and external-content injection risks were considered
- [ ] Adapted code, algorithms, and data can be attributed and licensed compatibly
- [ ] The proposal does not require private vulnerability disclosure

```text
Data / license constraints:
Security considerations:
```

---

## 📖 Documentation and release impact

- [ ] README or target-scope documentation
- [ ] Installation or deployment guidance
- [ ] Public API and financial-contract documentation
- [ ] Formula / VBA examples
- [ ] Numerical method, reference, and tolerance documentation
- [ ] Demo workbook or source
- [ ] Changelog and release notes
- [ ] Security policy
- [ ] No documentation impact

```text
Documentation needed:
Release-artifact impact:
```

---

## 🗺️ Dependencies and sequencing

What must exist first, and what future capabilities would depend on this work?

```text
Prerequisite issue / module / contract:
Downstream consumer:
Can be delivered incrementally:
Suggested first slice:
```

---

## ✅ Proposer checklist

- [ ] I searched existing issues and proposals
- [ ] I described a real use case rather than only an implementation idea
- [ ] I made relevant financial conventions explicit
- [ ] I identified acceptance criteria and non-goals
- [ ] I considered API compatibility and failure behavior
- [ ] I identified an independent reference or stated that one is missing
- [ ] I outlined boundary, invalid-input, invariant, and regression testing
- [ ] I considered Excel state, platform, locale, performance, and deployment impact
- [ ] I removed confidential, restricted, personal, and security-sensitive content
- [ ] I identified dependencies and the smallest defensible delivery slice

## ➕ Additional context

<!-- Add diagrams, synthetic examples, or references that clarify the proposal. -->
