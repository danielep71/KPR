<div align="center">

# 🧭 KPR Code of Conduct

### Respectful, evidence-led collaboration for financial and numerical engineering

[![Applies to](https://img.shields.io/badge/Applies_to-Everyone-217346?style=for-the-badge)](#scope)
[![Spaces](https://img.shields.io/badge/Spaces-Code_%7C_Issues_%7C_PRs-0969da?style=for-the-badge)](#scope)
[![Standard](https://img.shields.io/badge/Standard-Respectful_%2B_Evidence--Led-6f42c1?style=for-the-badge)](#technical-discussion)
[![Domain](https://img.shields.io/badge/Domain-Finance_%2B_Numerics-d97706?style=for-the-badge)](#technical-discussion)

<br>

**Technical rigor · Respectful disagreement · Reproducible evidence · Privacy-aware collaboration**

<br>

[Our pledge](#our-pledge)
&nbsp;·&nbsp;
[Expected behavior](#expected-behavior)
&nbsp;·&nbsp;
[Technical discussion](#technical-discussion)
&nbsp;·&nbsp;
[Confidentiality](#data-and-confidentiality)
&nbsp;·&nbsp;
[Enforcement](#enforcement)

</div>

---

**KPR** is an open-source Excel/VBA library for financial analytics and
instrument pricing.

The project benefits from careful disagreement: models can be ambiguous,
financial conventions can differ, floating-point results require judgment, and
Excel behavior can depend on host state. Those challenges make respectful,
precise, and reproducible collaboration essential.

People should feel comfortable:

- asking basic or advanced questions;
- reporting a numerical or implementation defect;
- challenging a model, convention, tolerance, or design decision;
- proposing a safer or more accurate alternative;
- identifying an unsupported market convention or Excel environment; and
- correcting an earlier conclusion when new evidence emerges.

> [!IMPORTANT]
> Technical rigor and respectful interaction are complementary requirements.
> Neither excuses the absence of the other.

---

<a id="our-pledge"></a>

## 🤝 Our pledge

Everyone who participates — through code, issues, pull requests, reviews,
documentation, numerical references, examples, releases, or project discussion —
is expected to help create a harassment-free experience for all.

That expectation applies regardless of:

- age, body size, disability, ethnicity, gender identity or expression;
- level of professional, academic, financial, numerical, VBA, or Excel
  experience;
- nationality, personal appearance, race, religion, or socioeconomic status;
- sexual identity or orientation; or
- any other personal characteristic unrelated to the contribution.

We commit to acting and interacting in ways that support an open, welcoming,
diverse, inclusive, and healthy community.

---

<a id="expected-behavior"></a>

## ✅ Expected behavior

Participants are expected to:

| Principle | Expected practice |
|---|---|
| 🤝 **Respect** | Assume good faith and address the work rather than the person. |
| 🎯 **Precision** | Distinguish observed fact, convention, assumption, hypothesis, and opinion. |
| 🧪 **Evidence** | Provide reproducible examples and relevant environment details where practical. |
| 🧭 **Transparency** | State uncertainty, limitations, conflicts of interest, and material dependencies. |
| 🔄 **Correction** | Acknowledge mistakes openly when better evidence changes the conclusion. |
| 🌱 **Inclusion** | Welcome contributors with different levels and types of expertise. |
| 🔐 **Stewardship** | Respect privacy, confidentiality, licensing, and intellectual-property boundaries. |
| 🧱 **Coherence** | Accept that maintainers may adopt, revise, defer, split, or decline a contribution. |

### Useful disagreement

A useful technical disagreement can be investigated:

> “The two outputs differ outside the stated tolerance. The comparison uses the
> same settlement date, ACT/365F basis, compounding convention, and cash-flow
> sign. The minimal input and independent reference are attached.”

That statement identifies a testable claim.

### Unhelpful disagreement

A personal judgment cannot be tested:

> “This is wrong because the author does not understand pricing.”

Both statements may arise from frustration with the same result. Only the first
helps improve the library.

---

<a id="unacceptable-behavior"></a>

## 🚫 Unacceptable behavior

Unacceptable behavior includes:

- harassment, intimidation, discrimination, or personal attacks;
- trolling, insulting or derogatory comments, and deliberately inflammatory
  language;
- unwelcome sexual attention or advances;
- publishing private or confidential information without permission;
- deliberately misrepresenting results, sources, authorship, test evidence, or
  model behavior;
- pressuring others to disclose employer, client, counterparty, student, or
  proprietary information;
- using credentials, reputation, job title, or academic status to silence a
  technical challenge rather than answering its substance;
- repeated disruption after a maintainer has asked participants to stop; and
- other conduct that would reasonably be considered inappropriate in a
  professional setting.

---

<a id="technical-discussion"></a>

## 🧪 Technical discussion standards

KPR discussions should make the evidence boundary visible.

### Evidence capsule

When reporting or debating behavior, provide the relevant items where practical:

| Evidence | Examples |
|---|---|
| 🧾 **Identity** | KPR release, commit, module, and public entry point |
| 🖥️ **Environment** | Excel and Windows versions, Office bitness, locale, and date system |
| 📐 **Contract** | Instrument definition, units, dates, signs, quote type, and expected output |
| 📅 **Conventions** | Day count, calendar, business-day rule, settlement, frequency, and compounding |
| 🔬 **Reproduction** | Smallest deterministic input and exact steps |
| 🎯 **Comparison** | Expected and observed values, absolute/relative error, and tolerance |
| 📚 **Reference** | Independent derivation, specification, library, vendor result, or published source |
| ⚠️ **Boundary** | What was not tested or cannot yet be concluded |

> [!NOTE]
> A screenshot may illustrate a result, but it does not replace the inputs,
> conventions, reference, and tolerance needed to reproduce a numerical claim.

### Evidence language

Prefer explicit classifications:

```text
Observed       reproduced directly
Derived        follows from stated inputs and equations
Inferred       best explanation, not independently proven
Expected       required by the documented contract
Unverified     plausible, but evidence is incomplete
Environment    specific to a tested Excel/Windows configuration
```

Do not present a plausible inference as a verified fact.

---

<a id="data-and-confidentiality"></a>

## 🔐 Data and confidentiality

Financial work often involves sensitive information. Sample portfolios,
workbooks, screenshots, market data, trade details, logs, and error reports must
not expose:

- client, employer, counterparty, student, or personal data;
- credentials, connection strings, internal URLs, or signing material;
- proprietary models, workbooks, curves, market data, or vendor content; or
- confidential business assumptions or positions.

Replace sensitive material with the smallest synthetic example that preserves
the behavior. Contributors must have the right to submit every code fragment,
document, image, dataset, and numerical reference they provide.

> [!CAUTION]
> Excel workbooks can contain hidden names, connections, cached values, VBA,
> metadata, links, and other information that is not visible on the active sheet.
> Sanitize and inspect a reproduction before uploading it.

---

<a id="scope"></a>

## 🌐 Scope

This Code of Conduct applies in all KPR project spaces, including:

- source code and committed artifacts;
- issues, pull requests, reviews, comments, and discussions;
- documentation, examples, release activity, and the project Wiki; and
- other public spaces where someone is representing KPR or its community.

It applies to maintainers, contributors, reviewers, users, and visitors alike.

---

<a id="enforcement"></a>

## ⚖️ Enforcement

The maintainer is responsible for clarifying and enforcing this Code of Conduct
and may remove, edit, or reject comments, commits, code, issues, and other
contributions that are inconsistent with it.

Responses may include:

1. clarification or a private reminder;
2. a formal warning;
3. editing or removing project content;
4. closing or locking a discussion;
5. rejecting or reverting a contribution;
6. temporary restriction from project participation;
7. permanent exclusion from project spaces; or
8. escalation to GitHub or another relevant platform.

Enforcement aims to be proportionate, consistent, protective of participants,
and protective of the technical record. Retaliation against anyone who reports a
concern or participates in its review is itself a violation.

Ordinary conduct concerns may be raised in the relevant issue or pull request.
Do not publish sensitive details. If a public report would expose personal or
confidential information, contact the maintainer through an established private
channel or use GitHub's reporting facilities.

---

## 🧩 Conflicts of interest

Disclose a material interest when it could reasonably affect technical review.
Examples include ownership of a competing implementation, commercial interest
in a proposed dependency, employer or client restrictions, or uncertainty about
the origin or license of submitted material.

A conflict is not automatically disqualifying. Undisclosed material influence
is the concern.

---

## 📜 Attribution

This Code of Conduct is informed by the
[Contributor Covenant, version 2.1](https://www.contributor-covenant.org/version/2/1/code_of_conduct.html)
and adapted for KPR's evidence-led financial and numerical engineering context.

---

<div align="center">

### Practical principle

**Be precise about the model · Be generous toward the person · Show the evidence · State the boundary · Protect the data**

<br>

Maintained by **Daniele Penza**

</div>
