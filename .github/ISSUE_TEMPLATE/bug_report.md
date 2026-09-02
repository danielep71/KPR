---
name: 🐞 Bug report
about: Report a reproducible KPR financial, numerical, Excel/VBA, compatibility, or packaging defect
title: "[Bug]: "
labels: bug
assignees: danielep71
---

<!--
  KPR bug report

  Thank you for helping improve KPR. Precision is more useful than volume:
  report only behavior, environments, and evidence you actually observed.

  SECURITY: do not use this public issue for a suspected vulnerability,
  credential, confidential workbook, restricted market data, malicious artifact,
  or exploitable trust-boundary issue. Follow SECURITY.md and report privately.
-->

<div align="center">

# 🐞 KPR Bug Report

[![Reproduction](https://img.shields.io/badge/Reproduction-Minimal-0969da?style=flat-square)](#minimal-reproduction)
[![Contract](https://img.shields.io/badge/Contract-Explicit-217346?style=flat-square)](#financial-contract)
[![Evidence](https://img.shields.io/badge/Evidence-Independent-d97706?style=flat-square)](#numerical-comparison)
[![Security](https://img.shields.io/badge/Security-Report_privately-d73a49?style=flat-square)](../../SECURITY.md#reporting-a-vulnerability)

</div>

> [!IMPORTANT]
> If the issue may disclose confidential data, expose a credential, execute
> unintended code, cross a workbook trust boundary, or compromise an official
> artifact, stop here and follow the private process in `SECURITY.md`.

## 📝 Summary

<!-- Describe the defect in one or two precise paragraphs. -->

## 💥 Practical impact

```text
Impact on calculation or workflow:
Severity from your perspective:     low / medium / high / blocking
Affected users or workbooks:
```

Examples of useful impact descriptions include an incorrect valuation,
unsupported valid input, corrupted workbook state, lost formula, inconsistent
UDF/VBA output, non-convergence, or an unusable release artifact.

---

## 🔖 Exact source state

Do not report only `latest`, `current`, or `main`.

```text
Release tag:           none / exact tag
Commit SHA:            full 40-character SHA if known
Branch:
Source obtained from:  official repository / GitHub Release / other
Deployment:            source workbook / add-in / development workbook / other
Artifact filename:
Artifact SHA-256:
```

---

## 🎚️ Affected area

Check all that apply.

- [ ] 📊 Worksheet UDF
- [ ] 🧱 VBA public API or result contract
- [ ] 📅 Date parsing, calendar, roll, day count, or schedule
- [ ] 📈 Rate, discount factor, compounding, or time-value calculation
- [ ] 📉 Curve, interpolation, extrapolation, or calibration
- [ ] 🧾 Cash flow or instrument definition
- [ ] 🧮 Pricing or analytics result
- [ ] 🔢 Solver, approximation, convergence, or numerical stability
- [ ] ⚙️ Excel workbook, worksheet, range, formula, or application state
- [ ] 🖥️ Office bitness, Windows, locale, or date-system compatibility
- [ ] 🧪 Test harness or numerical reference data
- [ ] 📦 Installation, packaging, release asset, or provenance
- [ ] 📖 Documentation
- [ ] Other

---

<a id="financial-contract"></a>

## 📐 Financial contract

Complete the fields relevant to the defect. A number cannot be assessed without
the conventions that give it meaning.

```text
Function / component:
Instrument or calculation:
Valuation date:
Settlement date:
Currency / notional / units:
Day-count basis:
Calendar and business-day rule:
Frequency / stub / end-of-month rule:
Compounding convention:
Quote type and scaling:
Long/short or payer/receiver sign:
Curve inputs and interpolation:
Expected output convention:
```

If a field is not relevant, delete it. Do not guess a convention that was not
actually used.

---

<a id="minimal-reproduction"></a>

## 🔁 Minimal reproduction

Reduce the problem to the smallest synthetic input and sequence that still
fails.

### Steps

1.
2.
3.

```text
Reproducibility:       always / often / intermittent / once
Approximate frequency:
First observed after:
```

### Smallest exact input

```text
Input name / field        Value
------------------        -----

```

### Smallest exact call or formula

```vb
Option Explicit

Public Sub ReproduceKPRIssue()

    ' Replace this comment with the smallest exact KPR call.

End Sub
```

```excel
=ReplaceWithExactWorksheetFormulaAndInputs
```

Delete the VBA or worksheet block if it does not apply.

---

## ✅ Expected behavior

<!-- State the required result and the contract or reference supporting it. -->

```text
Expected value or outcome:
Expected units:
Expected error / refusal behavior:
```

## ❌ Observed behavior

<!-- Include exact values and exact error text. -->

```text
Observed value or outcome:
Absolute error:
Relative error:
Runtime / worksheet error:
Diagnostic output:
```

---

<a id="numerical-comparison"></a>

## 🧮 Numerical comparison

Complete this section for an accuracy, pricing, analytics, curve, date, or
convergence defect.

```text
Independent reference:
Reference link / citation / version:
Reference conventions:
Reference precision:
Comparison rule:              absolute / relative / combined / other
Expected tolerance:
Observed error:
```

> [!WARNING]
> The KPR implementation under test is not an independent reference for its own
> expected result. A screenshot without inputs, conventions, and reference
> provenance is not enough to reproduce a numerical claim.

### Invariant or limiting-case evidence

- [ ] Monotonicity
- [ ] Symmetry or parity
- [ ] Bound or conservation rule
- [ ] Round trip
- [ ] Known limiting case
- [ ] Not applicable

```text
Invariant / limit and observed violation:
```

---

## 🖥️ Environment

```text
Excel version and build:
Office bitness:               32-bit / 64-bit
Windows version:
System locale:
Decimal separator:
Workbook date system:         1900 / 1904
Calculation mode:
Other loaded add-ins:
Fresh Excel process tested:   Yes / No
```

### Cross-environment result

| Environment | Result |
|---|---|
| Reported environment | fails / passes |
| Fresh Excel process | fails / passes / not tested |
| Other Office bitness | fails / passes / not tested |
| Other locale/date system | fails / passes / not tested |

---

## ⚙️ Workbook and Excel state

Complete if workbook content or application state affects the defect.

```text
Workbook type:
Target workbook / worksheet / range:
Protected sheet or workbook:
Formula / name / link / connection involved:
Application state before call:
Application state after call:
Events or calculation active:
```

State whether the issue persists after closing every Excel process and reopening
only the minimal reproduction.

---

## 🩹 Workaround

```text
Workaround available:         Yes / No
Workaround:
Why it is insufficient:
```

## 🎯 Proposed regression boundary

<!-- What permanent test would prove the defect stays fixed? -->

```text
Given:
When:
Then:
Tolerance / exact outcome:
```

---

## 🔐 Attachments, privacy, and provenance

- [ ] All examples use synthetic or legally redistributable data
- [ ] No client, employer, counterparty, student, or personal data is included
- [ ] No credential, connection string, internal URL, or signing material is included
- [ ] The workbook was inspected for hidden names, links, connections, VBA, metadata, and cached values
- [ ] Any external numerical reference can legally be cited or redistributed
- [ ] This report does not require private security disclosure

<!-- Drag sanitized screenshots or files here only when necessary. -->

---

## ✅ Reporter checklist

- [ ] I searched existing issues for the same behavior
- [ ] I identified the exact source state or artifact
- [ ] I reduced the issue to a minimal synthetic reproduction
- [ ] I stated the relevant financial conventions
- [ ] I separated expected behavior from observed behavior
- [ ] I identified an independent reference or invariant where numerical accuracy is involved
- [ ] I recorded only environments and tests actually observed
- [ ] I removed confidential, restricted, personal, and security-sensitive content

## ➕ Additional context

<!-- Add only information that helps reproduce, classify, or correct the defect. -->
