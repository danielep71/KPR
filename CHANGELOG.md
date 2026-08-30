# Changelog

All notable changes to **KPR** will be documented in this file.

The project intends to follow [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
and [Semantic Versioning](https://semver.org/) from its first published release.

## [Unreleased]

### Added

- A documented VBE export format for tracked VBA source, with a static rule that
  requires a unique `Attribute VB_Name` header matching each file name and
  rejects procedure-level description attributes.

### Changed

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
