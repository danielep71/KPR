<div align="center">

# 📜 Changelog

### Release history for Excel/VBA financial analytics and pricing

[![Format](https://img.shields.io/badge/Format-Keep_a_Changelog-0969da?style=flat-square)](https://keepachangelog.com/en/1.1.0/)
[![Versioning](https://img.shields.io/badge/Versioning-SemVer-6f42c1?style=flat-square)](https://semver.org/spec/v2.0.0.html)
[![Dates](https://img.shields.io/badge/Dates-YYYY--MM--DD-217346?style=flat-square)](#date-and-version-rules)
[![Staging](https://img.shields.io/badge/Staging-Unreleased_first-d97706?style=flat-square)](#unreleased)
[![Contributing](https://img.shields.io/badge/Changes-Contribution_guide-2ea44f?style=flat-square)](CONTRIBUTING.md)

<br>

**User-visible history · Explicit compatibility · Reproducible evidence · Immutable releases**

</div>

---

All notable changes to **KPR** are documented here.

This changelog follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
and [Semantic Versioning](https://semver.org/spec/v2.0.0.html). It records
released behavior and material unreleased changes; it is not a commit log, issue
tracker, or substitute for release evidence.

Versioning covers the documented worksheet/VBA API, financial and date conventions, numerical behavior, error contracts, compatibility, and supported environments.

---

## 🧭 Maintenance policy

- Add material changes under **Unreleased** in the same pull request as the
  behavior or documentation they describe.
- Write from the user's perspective: describe the observable result, contract,
  compatibility impact, and migration need.
- Link the owning issue or pull request when it contains useful engineering
  detail.
- Keep entries concise; do not duplicate implementation notes already preserved
  in source, issues, or technical documentation.
- Record only validation actually performed. State skipped environments and
  known limitations plainly.
- Move Unreleased entries into a dated version section during release.
- Do not edit a published release entry except to correct a demonstrable factual
  or link error; annotate material corrections instead of rewriting history.
- Never claim that a tag, binary, workbook, hash, test run, or environment was
  certified unless the evidence binds it to the released source.

See [CONTRIBUTING.md](CONTRIBUTING.md) for change and evidence requirements and
[SECURITY.md](SECURITY.md) for private vulnerability reporting.

<a id="date-and-version-rules"></a>

### Date and version rules

| Rule | Standard |
|---|---|
| Version | `MAJOR.MINOR.PATCH`, without the leading `v` in headings |
| Release heading | `## [X.Y.Z] - YYYY-MM-DD` |
| Date | Gregorian calendar date in ISO `YYYY-MM-DD` format |
| Ordering | Unreleased first; released versions newest to oldest |
| Comparison | Unreleased → latest tag; each release → preceding tag |
| Patch | Backward-compatible correction or hardening |
| Minor | Backward-compatible capability |
| Major | Incompatible public-contract change |
| Pre-release | State maturity and compatibility boundaries explicitly |

A repository may remain below `1.0.0` while its supported surface is still
forming. Pre-release status does not excuse undocumented breaking changes.

<details>
<summary><strong>Entry categories</strong></summary>

<br>

| Category | Use for |
|---|---|
| **Added** | New supported capabilities, APIs, files, or tests |
| **Changed** | Changes to existing behavior, contracts, tooling, or documentation |
| **Deprecated** | Supported behavior scheduled for removal |
| **Removed** | Removed capabilities or compatibility |
| **Fixed** | Corrected defects |
| **Security** | Safely disclosed security corrections |
| **Documentation** | Material documentation-only changes |
| **Validation** | Evidence actually produced |
| **Compatibility** | Upgrade or migration effects |
| **Known limitations** | Deliberate, unresolved boundaries |

Use only the categories needed by a release.

</details>

---

<a id="unreleased"></a>

## [Unreleased]

### Added

- Added a standardized installation and maintainer release documentation set with project-specific deployment, certification, provenance, recovery, and post-publication controls.

- Multi-cell evaluation on the single 22-name surface. Every value-taking
  function now accepts a scalar, a single-cell Range or 1x1 wrapper, or a
  multi-element Range or array at each value position, expands scalars,
  requires exact shape agreement between non-scalar arguments, and returns a
  1-based 2-D array of the resolved shape evaluated row-major, with each
  element resolved independently. An all-scalar call still returns a scalar.
  Multi-cell results are claimed on dynamic-array Excel only.
- `TryClassifyShape`, a classification-only preflight, so every wrapper
  applies the contract's call-level stages in order: host guard, classify
  every argument, controls and broadcast, cap, materialize, traverse.

- The array engine services in `KPR_Core_Array`: classification and
  one-read materialization of Ranges and VBA arrays, exact-shape broadcast
  resolution, a 100,000-element capacity gate decided from dimensions before
  any content is read, row-major element access, output allocation, and
  control unwrapping that reports `CONTROL_NOT_SCALAR` without reading a
  multi-element control. The 21 value-taking facade wrappers use these services
  while retaining their shared element implementations.
- A static purity rule forbidding Excel state, host classification,
  function-pointer dispatch and date intrinsics in the engine.

- The complete 22-name worksheet surface: `AddDays`, `BeginOfQuarter`,
  `EndOfQuarter`, `BeginOfYear` and `EndOfYear`, with the four calendar
  boundaries provided by `KPR_Core_Dates`.
- A private element implementation behind every value-taking public
  function, so the scalar call is the 1x1 case of the same code the array
  engine loops. Elements resolve their own value arguments in signature
  order and never run the host guard.
- An exact-inventory static rule: the facade must declare precisely the 22
  supported names, so a missing member, an extra `KPR_Dates_*` name, the
  legacy plural and any `_Spill` twin fail the build until #26's manifest
  replaces the hard-coded list.

- `Opt_Rounding` on `KPR_Dates_PillarFromDates`, accepting `NEAREST`,
  `FLOOR` and `CEILING` over one uniform candidate set. The `3W`/`1M`
  boundary under `NEAREST` is now derived from calendar-day distance rather
  than pinned at 25 days, and an anchor outside the supported window is
  excluded rather than approximated. `CEILING` returns `#NUM!` when no
  in-window anchor reaches the end date.
- Condition identifiers for every pillar grammar rejection, so a duplicate
  unit, a signed alias, a malformed token and a non-text payload are
  distinguishable to a caller.
- `KPR_Core_Dates` may now depend on `KPR_Core_Err` for the condition
  vocabulary; it still never constructs a worksheet error value.

- The worksheet/VBA date-system policy: every value-taking function inspects
  its caller once before touching any argument, refuses an identified 1904
  worksheet host with one call-level `#N/A`, and proceeds under the documented
  1900 serial contract when no worksheet host can be identified. No
  active-workbook fallback exists, and the static gate enforces guard
  placement, volatility scope and the absence of a fallback.
- `KPR_Dates_HostDateSystem`, a deliberately volatile scalar diagnostic that
  reports `1900`, `1904` or `#N/A` for the current caller, discriminating a
  library-produced `#N/A` from a propagated one.
- A macro-only `KPR_Tests_RunHost` runner that exercises the real
  worksheet-`Range` caller path under both date systems in a scratch
  workbook, alongside a pure `host` suite in the dispatcher.

- Strict scalar input parsing: ISO-8601 `YYYY-MM-DD` text only, validated
  component by component, with locale-formatted and numeric-looking text
  rejected rather than reinterpreted. Numeric-looking text is recognized by a
  fixed rule rather than by `IsNumeric`, which consults the host locale and
  reads `31.12.2026` as a number in some regions.
- A condition vocabulary mirroring the contract's registry identifiers, so a
  failure is classified rather than collapsed into a single error value.
- Deterministic tests for every accepted and rejected scalar input class, and
  static rules forbidding locale-sensitive date conversions in production and
  pinning the two representations of the supported window against each other.

- A five-module calculation architecture: `KPR_Core_Err`, `KPR_Core_Parse`,
  `KPR_Core_Dates` and `KPR_Core_Array` behind the worksheet facade
  `KPR_DATES_DAYS`, replacing the monolithic `KPR_Dates_Days` module. The
  migration is structural: the sixteen existing functions keep their behaviour.
- Static rules for module visibility, the allowed-dependency matrix, and public
  worksheet functions being declared only in a non-private facade module.
- A documented VBE export format for tracked VBA source, with a static rule that
  requires a unique `Attribute VB_Name` header matching each file name and
  rejects procedure-level description attributes.
- A stable condition-identifier registry in the date-layer contract, so fixtures
  and certification evidence can cite an originating condition that the returned
  Excel error value cannot express.

### Changed

- Reconciled the README, installation, contribution, security, contract, and
  implementation-plan status with the implemented v0.0.2 date surface while
  preserving #29 as the only final exact-source certification gate.
- Corrected the v0.0.1 history to describe repository setup only, not a
  functional date or business-day baseline.

- A multi-cell optional control now reports `CONTROL_NOT_SCALAR` rather than
  the generic shape rejection, from its dimensions alone.

- `KPR_Dates_DaysInYear` and `KPR_Dates_IsLeapYear` take `YearIn`, a
  calendar year, as the contract specifies; a date is now rejected rather
  than silently read as a serial.
- `KPR_Dates_DatesFromPillar` is renamed to the singular
  `KPR_Dates_DateFromPillar`. No alias remains.
- User-facing error documentation no longer lists an unexpected runtime
  error as a `#VALUE!` condition; the containment handlers are a defect if
  reached, not a documented outcome.

- An incoming Excel error at the `Pillar` argument of `DateFromPillar` now
  propagates verbatim instead of being rejected as a non-text payload.
- Pillar text and rounding tokens are normalized by ASCII-only case folding
  and trimming rather than `UCase$`/`Trim$`, so the result is identical under
  every locale, including Turkish casing.

- Restored the declared defaults on the existing optional Boolean controls,
  avoiding any reliance on forwarding VBA's missing-argument marker through a
  non-optional helper, and completed the scalar type matrix for all native
  numeric Variant subtypes (including 64-bit `LongLong`) and rejected objects.
- Value, integer and control arguments are declared `Variant`, so a fractional
  shift count or a numeric stand-in for a Boolean control is rejected instead of
  being silently coerced by Excel before the function is entered.
- An out-of-window date now returns `#NUM!` rather than `#VALUE!`, and an
  incoming Excel error propagates verbatim instead of being collapsed into a
  parse failure.

- Component-name uniqueness is now compared case-insensitively, matching VBA's
  case-insensitive component namespace. File-stem matching stays case-sensitive
  for export fidelity.
- Aligned the VBE source-format documentation with the current
  `KPR_DATES_DAYS` facade and clarified that static conformance does not prove a
  Windows import or normalized round trip.

- Refined the KPR social preview by removing the maintainer footer and adding
  clearer financial-market imagery and pricing formulas.
- Reconciled the README release status and repository tree with the published
  `v0.0.1` pre-release and the tracked tooling and version file.
- Hardened the required-file gate around the label synchronizer, social
  preview, version marker, and certified KPR-native source baseline.
- Moved the README banner to an archive-safe asset path so release source
  archives remain self-contained.
- Removed stale repository-policy references to nonexistent distribution
  directories and aligned visual-asset policy with tracked content.
- Changed `KPR_Dates_DaysInYear` and `KPR_Dates_IsLeapYear` to take a calendar
  year rather than a date, removing an ambiguity in which a bare year such as
  `2024` was a valid serial and silently answered for 1905. A date argument is
  now rejected; year 1900 is in the domain because neither function constructs
  a date.
- Tightened the date-layer contract: pillar tokens reject a repeated unit
  instead of accumulating it, pillar parsing trims outer whitespace while
  rejecting internal whitespace, the emitted pillar grammar is stated as a
  narrower canonical subset, and an omitted or blank optional control selects
  its documented default.
- Made the plan and the date-layer contract cite each other with an explicit
  authority split, and required the contract in the static file inventory.
- Clarified that defensive internal catch-alls are unreachable containment,
  not a supported error condition; any activation is a contract defect and a
  regression/certification failure.
- Reconciled the v0.0.2 architecture plan with the `KPR_DATES_DAYS` facade,
  case-insensitive VBA component identity, and #11's structural-only migration
  boundary while retaining facade/core casing as a documentation convention.

## [0.0.1] - 2026-08-29

### Added

- A KPR-native repository identity and source-first project structure, without
  a supported analytical or business-day API.
- Premium KPR governance, contribution, installation, security, pull-request,
  bug-report, and feature-request documentation.
- A canonical 23-label manifest with a self-healing GitHub label workflow.
- Hosted static repository checks covering repository identity, structured data,
  Markdown links, text integrity, forbidden artifacts, file policy, label
  definitions, immutable workflow actions, whitespace, and VBA declarations.
- Shared editor, line-ending, encoding, attribute, and ignore policies.
- KPR repository metadata, security reporting, branch protection, release-tag
  protection, and an explicit suite-wide merge policy.
- A suite-aligned social preview, also used as the README banner, with an honest
  under-development status.

### Changed

- Removed the inherited template implementation and all stale identity,
  routing, links, and claims.
- Reset the repository documentation and issue surfaces around KPR's intended
  financial-analytics and instrument-pricing scope.

### Release status

- This is a **repository-setup pre-release only**. It certifies the KPR project
  baseline and its governance and automation surfaces; it does not provide a
  supported analytical API, installation package, Excel execution CI,
  numerical accuracy certification, production support, or a stable functional
  release.

[Unreleased]: https://github.com/danielep71/KPR/compare/v0.0.1...HEAD
[0.0.1]: https://github.com/danielep71/KPR/releases/tag/v0.0.1
