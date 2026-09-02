# Changelog

All notable changes to **KPR** will be documented in this file.

The project intends to follow [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
and [Semantic Versioning](https://semver.org/) from its first published release.

## [Unreleased]

### Added

- The array engine services in `KPR_Core_Array`: classification and
  one-read materialization of Ranges and VBA arrays, exact-shape broadcast
  resolution, a 100,000-element capacity gate decided from dimensions before
  any content is read, row-major element access, output allocation, and
  control unwrapping that reports `CONTROL_NOT_SCALAR` without reading a
  multi-element control. No public behaviour changes; #17 wires the facade.
- A static purity rule forbidding Excel state, host classification,
  function-pointer dispatch and date intrinsics in the engine.

- The complete 22-name worksheet surface: `AddDays`, `BeginOfQuarter`,
  `EndOfQuarter`, `BeginOfYear` and `EndOfYear`, with the four calendar
  boundaries provided by `KPR_Core_Dates`.
- A private element implementation behind every value-taking public
  function, so the scalar call is the 1x1 case of the same code the array
  engine will loop. Elements resolve their own value arguments in signature
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

- `KPR_Dates_DaysInYear` and `KPR_Dates_IsLeapYear` take `YearIn`, a
  calendar year, as the contract specifies; a date is now rejected rather
  than silently read as a serial.
- `KPR_Dates_DatesFromPillar` is renamed to the singular
  `KPR_Dates_DateFromPillar`. No alias remains.
- User-facing error documentation no longer lists an unexpected runtime
  error as a `#VALUE!` condition; the containment handlers are a defect if
  reached, not a documented outcome.

- An incoming Excel error at the `Pillar` argument of `DatesFromPillar` now
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

- A KPR-native scalar date and business-day source baseline with native Excel
  error returns.
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
