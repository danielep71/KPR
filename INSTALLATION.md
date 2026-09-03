# 📦 KPR Installation

[![Status: development only](https://img.shields.io/badge/status-development%20only-6f42c1)](#support-status)
[![Version: 0.0.1](https://img.shields.io/badge/version-0.0.1-blue)](VERSION)
[![Security policy](https://img.shields.io/badge/security-policy-success)](SECURITY.md)
[![License: MIT](https://img.shields.io/badge/license-MIT-yellow.svg)](LICENSE)

KPR does not yet publish a supported workbook, add-in, or production-ready source package. This guide defines the safe developer-evaluation path and the gate that a future installable release must satisfy.

> [!IMPORTANT]
> Do not present the current `main` branch as an end-user installation. It is development material, and `VERSION` remains the authoritative release marker.

<a id="support-status"></a>
## 🧭 Support status

| Route | Status | Intended use |
| --- | --- | --- |
| Packaged workbook or add-in | Not published | No supported binary artifact exists |
| Production source import | Not published | Public contracts and dependencies are not yet frozen |
| Developer evaluation | Available | Review and controlled testing of the repository sources |
| Snippet-by-snippet copying | Unsupported | Loses dependencies, ordering, and traceability |

## 🧰 Prerequisites

- Windows desktop Excel with access to the Visual Basic Editor.
- Macro execution allowed only for code you have reviewed and trust.
- A disposable macro-enabled workbook for evaluation.
- A local clone or a source archive from an exact commit.
- Familiarity with KPR's development contracts and [VBE export guidance](docs/VBE_EXPORT.md).

## 🧪 Developer evaluation

1. Clone the repository or download an archive for an exact commit.
2. Record the commit SHA before importing anything.
3. Read [README.md](README.md), [CONTRIBUTING.md](CONTRIBUTING.md), and the relevant files under `docs/`.
4. Review every VBA source file before enabling macros.
5. Open a new disposable `.xlsm` workbook.
6. Import only the modules required by the scenario being evaluated, preserving documented module names and dependencies.
7. In the Visual Basic Editor, run **Debug → Compile VBAProject**.
8. Exercise the narrow scenario under review and retain the commit SHA with the result.

Developer evaluation is evidence gathering, not installation. Do not distribute the resulting workbook as an official KPR release.

## 🚦 First-installable-release gate

Before KPR can claim a supported installation path, the repository must provide all of the following:

- A frozen public API and explicit compatibility statement.
- A complete, ordered production-source manifest.
- Supported Excel, Windows, and Office-bitness boundaries.
- A clean compile and project-specific automated regression evidence.
- Independent numerical or financial validation for calculation-sensitive behavior.
- A reproducible package built from the exact tagged source.
- Hashes and provenance that bind each published artifact to that tag.
- Upgrade, removal, rollback, and known-limit documentation.
- A release prepared according to [RELEASING.md](RELEASING.md).

## 🔐 Safety and trust

VBA runs with the permissions of the current Office process. Review source, use trusted locations deliberately, and never enable macros solely because a workbook displays a prompt. Report suspected vulnerabilities privately through [SECURITY.md](SECURITY.md).

## 🧯 Troubleshooting

### The code does not compile

Confirm that all required companion modules were imported, module names were preserved, and no host-specific references were added accidentally.

### A document appears to describe a finished product

Treat `VERSION`, the [changelog](CHANGELOG.md), and a published GitHub Release as the release boundary. Development documentation does not turn an unreleased branch into a supported installation.

### I need a distributable workbook now

No official KPR artifact is currently available. Building one locally creates an unsupported derivative that must not be represented as an upstream release.

## 🔄 Removal

Delete the disposable evaluation workbook and any locally exported copies. If you imported KPR code into another workbook, remove those modules manually and compile the remaining project before saving.

## 📚 Related documents

- [README.md](README.md) — project purpose and current status
- [CONTRIBUTING.md](CONTRIBUTING.md) — contribution workflow
- [RELEASING.md](RELEASING.md) — maintainer release controls
- [CHANGELOG.md](CHANGELOG.md) — release history
- [SECURITY.md](SECURITY.md) — private vulnerability reporting
- [VERSION](VERSION) — authoritative project version
- [LICENSE](LICENSE) — MIT license
- [VBE export guidance](docs/VBE_EXPORT.md) — source-export conventions

---

**Installation principle:** KPR becomes installable only when a reproducible, tested artifact is published from an exact tagged source revision.
