#!/usr/bin/env python3
"""KPR repository-integrity gate.

The checker intentionally uses only the Python standard library so the hosted
gate can run on a clean GitHub Ubuntu runner without installing dependencies.
It validates repository structure and source hygiene; it does not compile VBA
or replace Excel regression and release certification.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import os
from pathlib import Path, PurePosixPath
import re
from datetime import date, timedelta
import shutil
import subprocess
import sys
import tempfile
from typing import Callable, Iterable
from urllib.parse import unquote, urlsplit
import xml.etree.ElementTree as ET


SCHEMA_VERSION = 1
TOOL_NAME = "KPR repository static checks"
SCOPE_NOTE = (
    "Hosted static checks complement, but do not replace, Excel compilation "
    "and regression certification."
)

REQUIRED_FILES = (
    ".editorconfig",
    ".gitattributes",
    ".gitignore",
    ".github/ISSUE_TEMPLATE/bug_report.md",
    ".github/ISSUE_TEMPLATE/config.yml",
    ".github/ISSUE_TEMPLATE/feature_request.md",
    ".github/PULL_REQUEST_TEMPLATE.md",
    ".github/labels.json",
    ".github/scripts/labels-sync.mjs",
    ".github/workflows/labels-sync.yml",
    ".github/workflows/static-checks.yml",
    "CHANGELOG.md",
    "CODE_OF_CONDUCT.md",
    "CONTRIBUTING.md",
    "INSTALLATION.md",
    "LICENSE",
    "README.md",
    "SECURITY.md",
    "VERSION",
    "assets/social-preview.png",
    "docs/DATE_LAYER_CONTRACT.md",
    "docs/VBE_EXPORT.md",
    "src/modules/KPR_Core_Array.bas",
    "src/modules/KPR_Core_Dates.bas",
    "src/modules/KPR_Core_Err.bas",
    "src/modules/KPR_Core_Parse.bas",
    "src/modules/KPR_DATES_DAYS.bas",
    "test/modules/KPR_REGRESSION_TESTS.bas",
    "tools/check_repo.py",
)

VBA_SOURCE_SUFFIXES = {".bas", ".cls", ".frm"}
VBA_COMPONENT_NAME_LIMIT = 31
VBA_HEADER_SCAN_LINES = 12
VBA_PROCEDURE_ATTRIBUTE_PREFIXES = (
    "Attribute VB_Description",
    "Attribute VB_ProcData",
)

VBA_PRIVATE_MODULE_DIRECTIVE = "Option Private Module"
VBA_FACADE_STEM_PREFIX = "kpr_dates_"
VBA_INTERNAL_MODULE_PREFIX = "kpr_core_"
VBA_PUBLIC_FUNCTION_PREFIX = "kpr_dates_"

# Allowed production dependencies, keyed by component name folded to lower case.
# Mirrors the matrix in docs/IMPLEMENTATION_PLAN.md; a module may reference only
# the public members of the components listed for it.
VBA_ALLOWED_DEPENDENCIES: dict[str, frozenset[str]] = {
    "kpr_core_err": frozenset(),
    "kpr_core_dates": frozenset(),
    "kpr_core_parse": frozenset({"kpr_core_err"}),
    "kpr_core_array": frozenset({"kpr_core_err"}),
    "kpr_dates_days": frozenset(
        {"kpr_core_err", "kpr_core_parse", "kpr_core_dates", "kpr_core_array"}
    ),
    # The regression suites assert exact condition classification, so they call
    # the parser and the error vocabulary directly. Everything else they exercise
    # through the facade. They are deliberately not granted access to every core.
    "kpr_regression_tests": frozenset(
        {"kpr_core_err", "kpr_core_parse", "kpr_dates_days"}
    ),
}

# Locale-sensitive conversions have no place in production parsing. IsNumeric is
# on the list because it consults the host: under a locale that groups with ".",
# it reads "31.12.2026" as a number. CDate is not listed: it is still used to
# convert an already-validated numeric serial to a Date, which is deterministic.
# What must never appear is a conversion that asks the host how to read text.
VBA_FORBIDDEN_PARSE_CALLS = ("IsDate", "DateValue", "CVDate", "IsNumeric")

# The supported window is owned by KPR_Core_Dates as Date constants. KPR_Core_Parse
# may not depend on that module, so it restates the same bounds as serials and
# years. These two representations are pinned against each other here.
VBA_WINDOW_PINS = {
    "kpr_core_dates": {"KPR_MIN_DATE": "1900-03-01", "KPR_MAX_DATE": "9999-12-31"},
    "kpr_core_parse": {
        "KPR_MIN_SERIAL": "1900-03-01",
        "KPR_MAX_SERIAL": "9999-12-31",
        "KPR_MIN_YEAR": "1900",
        "KPR_MAX_YEAR": "9999",
    },
}

# Public members each architecture component must keep declaring, keyed by
# component name folded to lower case. A rename or a deletion here breaks
# callers that this gate cannot compile, so the internal surface is pinned by
# name rather than left to Windows-only verification.
VBA_REQUIRED_MEMBERS: dict[str, frozenset[str]] = {
    "kpr_core_err": frozenset({"ErrValue", "ErrNum", "ErrNA", "ErrForCondition"}),
    "kpr_core_array": frozenset({"Array_Rank", "TryUnwrapScalar"}),
    "kpr_core_parse": frozenset(
        {"TryParseDateScalar", "TryParseLongScalar", "TryParseBoolControl"}
    ),
    "kpr_core_dates": frozenset(
        {
            "KPR_MIN_DATE",
            "KPR_MAX_DATE",
            "IsDateInWindow",
            "IsLeapYear",
            "DaysInMonth",
            "EndOfMonth",
            "TryAddMonths",
            "TryPillar_Parse",
            "Pillar_Format_Nearest",
        }
    ),
    "kpr_dates_days": frozenset({"KPR_Dates_HostDateSystem"}),
}

# Host date-system policy: names the facade rules pin to. The guard must appear
# exactly once in every value-taking public function, before any argument is
# resolved; the diagnostic must call the classifier directly instead.
VBA_HOST_GUARD = "PassHostGuard"
VBA_HOST_CLASSIFIER = "TryResolveHostDateSystem"
VBA_HOST_DIAGNOSTIC = "KPR_Dates_HostDateSystem"
VBA_VOLATILE_CALL = "Application.Volatile True"
VBA_ARGUMENT_RESOLVERS = ("TryResolveDate", "TryResolveLong", "TryResolveBool")

# Date-system authority belongs to the caller's own workbook and nothing else.
# The prohibition is scoped to the facade's host-resolution path rather than
# all of src/, because later registration, UI or demo code may have legitimate
# explicit workbook operations. Comments and string literals are ignored.
VBA_WORKBOOK_FALLBACKS = ("ActiveWorkbook", "ThisWorkbook", "ActiveSheet")
VBA_HOST_PATH_PROCEDURES = (VBA_HOST_CLASSIFIER, VBA_HOST_GUARD, VBA_HOST_DIAGNOSTIC)

# DateSerial with a day argument of zero names the day before the first of the
# given month. The idiom is compact but not total: it raises error 5 whenever
# the implied date leaves the VBA Date range, which is the case for month 13 of
# year 9999 and for month 1 of the floor year. Behind a defensive handler that
# surfaces as a worksheet error indistinguishable from a genuine rejection, so
# month length is read from KPR_Core_Dates.DaysInMonth instead.
VBA_DAY_ZERO_FUNCTION = "DateSerial"
VBA_DAY_ZERO_ARGUMENTS = 3

# Every supported worksheet function returns Variant, because a native Excel
# error value cannot be carried by any narrower return type.
VBA_FACADE_RETURN_TYPE = "variant"

TEXT_SUFFIXES = {
    ".bas",
    ".bat",
    ".cfg",
    ".cjs",
    ".cls",
    ".cmd",
    ".csv",
    ".frm",
    ".ini",
    ".js",
    ".json",
    ".jsonc",
    ".md",
    ".mjs",
    ".ps1",
    ".psd1",
    ".psm1",
    ".py",
    ".pyw",
    ".reg",
    ".sh",
    ".svg",
    ".toml",
    ".tsv",
    ".txt",
    ".vbs",
    ".xml",
    ".yaml",
    ".yml",
}
TEXT_NAMES = {
    ".editorconfig",
    ".gitattributes",
    ".gitignore",
    ".gitkeep",
    "CITATION.cff",
    "LICENSE",
}
CROSS_PLATFORM_SUFFIXES = {
    ".cfg",
    ".cjs",
    ".csv",
    ".js",
    ".json",
    ".jsonc",
    ".md",
    ".mjs",
    ".py",
    ".pyw",
    ".sh",
    ".svg",
    ".toml",
    ".tsv",
    ".txt",
    ".xml",
    ".yaml",
    ".yml",
}
WINDOWS_TEXT_SUFFIXES = {
    ".bas",
    ".bat",
    ".cls",
    ".cmd",
    ".frm",
    ".ini",
    ".ps1",
    ".psd1",
    ".psm1",
    ".reg",
    ".vbs",
}
OFFICE_BINARY_SUFFIXES = {
    ".accdb",
    ".accde",
    ".accdr",
    ".ade",
    ".doc",
    ".docm",
    ".docx",
    ".dotm",
    ".dotx",
    ".mdb",
    ".mde",
    ".potm",
    ".potx",
    ".ppsm",
    ".ppsx",
    ".ppt",
    ".pptm",
    ".pptx",
    ".xla",
    ".xlam",
    ".xls",
    ".xlsb",
    ".xlsm",
    ".xlsx",
    ".xlt",
    ".xltm",
    ".xltx",
    ".xlw",
}
SECRET_SUFFIXES = {".key", ".p12", ".pem", ".pfx", ".pvk"}


def finding(path: str, message: str, line: int | None = None) -> dict[str, object]:
    item: dict[str, object] = {"path": path, "message": message}
    if line is not None:
        item["line"] = line
    return item


def rule_result(
    rule_id: str,
    title: str,
    failures: list[dict[str, object]],
    success_summary: str,
) -> dict[str, object]:
    if failures:
        count = len(failures)
        summary = f"{count} finding{'s' if count != 1 else ''}"
        status = "fail"
    else:
        summary = success_summary
        status = "pass"
    return {
        "id": rule_id,
        "title": title,
        "status": status,
        "summary": summary,
        "findings": failures,
    }


class Repository:
    def __init__(self, root: Path):
        self.root = root.resolve()
        self.files = self._tracked_files()

    def _git(self, *arguments: str, check: bool = True) -> subprocess.CompletedProcess[str]:
        return subprocess.run(
            ["git", "-C", str(self.root), *arguments],
            check=check,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )

    def _tracked_files(self) -> tuple[str, ...]:
        completed = subprocess.run(
            ["git", "-C", str(self.root), "ls-files", "-z"],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        return tuple(
            sorted(
                item.decode("utf-8", errors="surrogateescape")
                for item in completed.stdout.split(b"\0")
                if item
            )
        )

    def path(self, relative: str) -> Path:
        return self.root / PurePosixPath(relative)

    def bytes(self, relative: str) -> bytes:
        return self.path(relative).read_bytes()

    def text(self, relative: str) -> str:
        data = self.bytes(relative)
        if Path(relative).suffix.lower() in {".bas", ".cls", ".frm"}:
            return data.decode("cp1252")
        return data.decode("utf-8")

    def commit(self) -> str | None:
        completed = self._git("rev-parse", "HEAD", check=False)
        return completed.stdout.strip() or None if completed.returncode == 0 else None


def is_text_file(path: str) -> bool:
    pure = PurePosixPath(path)
    return pure.suffix.lower() in TEXT_SUFFIXES or pure.name in TEXT_NAMES


def line_number(text: str, offset: int) -> int:
    return text.count("\n", 0, offset) + 1


def check_required_files(repo: Repository) -> dict[str, object]:
    tracked = set(repo.files)
    failures = [
        finding(path, "Required repository file is not tracked.")
        for path in REQUIRED_FILES
        if path not in tracked
    ]
    return rule_result(
        "required-files",
        "Required repository files",
        failures,
        f"All {len(REQUIRED_FILES)} required files are tracked",
    )


def check_stale_identity(repo: Repository) -> dict[str, object]:
    joined_name = "Date" + "Picker"
    tokens = (
        "vba-" + "datetime" + "picker",
        "datetime" + "picker",
        "date" + "picker",
        "date" + " picker",
    )
    failures: list[dict[str, object]] = []
    for path in repo.files:
        if not is_text_file(path):
            continue
        try:
            text = repo.text(path)
        except (OSError, UnicodeError):
            continue
        folded = text.casefold()
        for token in tokens:
            start = 0
            while True:
                offset = folded.find(token, start)
                if offset < 0:
                    break
                failures.append(
                    finding(
                        path,
                        f"Inherited {joined_name} identity token must be replaced with KPR identity.",
                        line_number(text, offset),
                    )
                )
                start = offset + len(token)
    return rule_result(
        "stale-identity",
        "Template identity",
        failures,
        "No inherited template identity remains",
    )


def _strip_yaml_comment(value: str) -> str:
    single = False
    double = False
    escaped = False
    for index, character in enumerate(value):
        if escaped:
            escaped = False
            continue
        if character == "\\" and double:
            escaped = True
        elif character == "'" and not double:
            single = not single
        elif character == '"' and not single:
            double = not double
        elif character == "#" and not single and not double:
            if index == 0 or value[index - 1].isspace():
                return value[:index].rstrip()
    return value.rstrip()


def _yaml_scalar_error(value: str) -> str | None:
    single = False
    double = False
    escaped = False
    stack: list[str] = []
    pairs = {"]": "[", "}": "{"}
    for character in value:
        if escaped:
            escaped = False
            continue
        if character == "\\" and double:
            escaped = True
        elif character == "'" and not double:
            single = not single
        elif character == '"' and not single:
            double = not double
        elif not single and not double and character in "[{":
            stack.append(character)
        elif not single and not double and character in "]}":
            if not stack or stack.pop() != pairs[character]:
                return "has unbalanced flow brackets"
    if single or double:
        return "has an unterminated quoted scalar"
    if stack:
        return "has unbalanced flow brackets"
    return None


def validate_yaml_subset(text: str) -> list[tuple[int, str]]:
    """Validate the conservative YAML dialect used by GitHub repository files."""

    errors: list[tuple[int, str]] = []
    block_parent_indent: int | None = None
    mapping = re.compile(r"^(?:[A-Za-z0-9_.${}\-]+|'[^']+'|\"[^\"]+\")\s*:(?:\s*(.*))?$")

    for number, raw_line in enumerate(text.splitlines(), start=1):
        if not raw_line.strip() or raw_line.lstrip().startswith("#"):
            continue
        leading = raw_line[: len(raw_line) - len(raw_line.lstrip(" \t"))]
        if "\t" in leading:
            errors.append((number, "tab indentation is not allowed"))
            continue
        indent = len(leading)
        if block_parent_indent is not None:
            if indent > block_parent_indent:
                continue
            block_parent_indent = None
        if indent % 2:
            errors.append((number, "indentation must use two-space levels"))
            continue

        content = _strip_yaml_comment(raw_line[indent:])
        if not content:
            continue
        candidate = content
        if candidate == "-":
            continue
        if candidate.startswith("- "):
            candidate = candidate[2:].strip()
            if not candidate:
                continue
            if not mapping.match(candidate):
                scalar_error = _yaml_scalar_error(candidate)
                if scalar_error:
                    errors.append((number, scalar_error))
                continue

        match = mapping.match(candidate)
        if not match:
            errors.append((number, "expected a mapping entry or sequence item"))
            continue
        scalar = (match.group(1) or "").strip()
        if scalar in {"|", "|-", "|+", ">", ">-", ">+"}:
            block_parent_indent = indent
            continue
        scalar_error = _yaml_scalar_error(scalar)
        if scalar_error:
            errors.append((number, scalar_error))
    return errors


def check_structured_data(repo: Repository) -> dict[str, object]:
    failures: list[dict[str, object]] = []
    json_count = 0
    yaml_count = 0
    xml_count = 0
    for path in repo.files:
        suffix = PurePosixPath(path).suffix.lower()
        if suffix == ".json":
            json_count += 1
            try:
                json.loads(repo.text(path))
            except (OSError, UnicodeError, json.JSONDecodeError) as error:
                line = getattr(error, "lineno", None)
                failures.append(finding(path, f"Invalid JSON: {error}", line))
        elif suffix in {".yml", ".yaml"}:
            yaml_count += 1
            try:
                text = repo.text(path)
            except (OSError, UnicodeError) as error:
                failures.append(finding(path, f"Cannot decode YAML as UTF-8: {error}"))
                continue
            for number, message in validate_yaml_subset(text):
                failures.append(finding(path, f"Invalid YAML structure: {message}.", number))
        elif suffix == ".xml":
            xml_count += 1
            try:
                ET.fromstring(repo.text(path))
            except (OSError, UnicodeError, ET.ParseError) as error:
                line = error.position[0] if isinstance(error, ET.ParseError) else None
                failures.append(finding(path, f"Invalid XML: {error}", line))
    summary = f"Parsed {json_count} JSON, {yaml_count} YAML, and {xml_count} XML files"
    return rule_result("structured-data", "JSON, YAML, and Ribbon XML", failures, summary)


def _markdown_destinations(text: str) -> Iterable[tuple[int, str]]:
    fenced = False
    fence = ""
    inline = re.compile(r"!?\[[^\]]*\]\(([^)]+)\)")
    reference = re.compile(r"^\s*\[[^\]]+\]:\s*(\S+)")
    for number, line in enumerate(text.splitlines(), start=1):
        stripped = line.lstrip()
        if stripped.startswith(("```", "~~~")):
            marker = stripped[:3]
            if not fenced:
                fenced = True
                fence = marker
            elif marker == fence:
                fenced = False
            continue
        if fenced:
            continue
        for match in inline.finditer(line):
            yield number, match.group(1).strip()
        match = reference.match(line)
        if match:
            yield number, match.group(1).strip()


def _split_destination(raw: str) -> tuple[str, str]:
    value = raw.strip()
    if value.startswith("<") and ">" in value:
        value = value[1 : value.index(">")]
    elif re.search(r"\s", value):
        value = value.split(None, 1)[0]
    parsed = urlsplit(value)
    return unquote(parsed.path), unquote(parsed.fragment)


def _github_slugs(text: str) -> set[str]:
    slugs: set[str] = set()
    counts: dict[str, int] = {}
    fenced = False
    fence = ""
    for line in text.splitlines():
        stripped = line.lstrip()
        if stripped.startswith(("```", "~~~")):
            marker = stripped[:3]
            if not fenced:
                fenced = True
                fence = marker
            elif marker == fence:
                fenced = False
            continue
        if fenced:
            continue
        for anchor in re.finditer(r"<(?:a\s+(?:id|name)|[A-Za-z][^>]*\s+id)=[\"']([^\"']+)[\"']", line, re.IGNORECASE):
            slugs.add(anchor.group(1).casefold())
        match = re.match(r"^\s{0,3}#{1,6}\s+(.+?)\s*#*\s*$", line)
        if not match:
            continue
        heading = re.sub(r"<[^>]+>", "", match.group(1)).casefold()
        heading = re.sub(r"[`*_~]", "", heading)
        heading = re.sub(r"[^\w\- ]", "", heading, flags=re.UNICODE)
        base = re.sub(r"\s+", "-", heading.strip())
        occurrence = counts.get(base, 0)
        counts[base] = occurrence + 1
        slugs.add(base if occurrence == 0 else f"{base}-{occurrence}")
    return slugs


def check_markdown_links(repo: Repository) -> dict[str, object]:
    failures: list[dict[str, object]] = []
    checked = 0
    slug_cache: dict[Path, set[str]] = {}
    for path in repo.files:
        if PurePosixPath(path).suffix.lower() != ".md":
            continue
        try:
            text = repo.text(path)
        except (OSError, UnicodeError):
            continue
        source = repo.path(path)
        for number, raw in _markdown_destinations(text):
            parsed = urlsplit(raw.strip("<>"))
            if parsed.scheme or raw.startswith("//"):
                continue
            target_text, fragment = _split_destination(raw)
            if not target_text and not fragment:
                continue
            checked += 1
            target = source if not target_text else (source.parent / target_text)
            try:
                target = target.resolve()
                target.relative_to(repo.root)
            except (OSError, ValueError):
                failures.append(finding(path, f"Relative link escapes the repository: {raw}", number))
                continue
            if not target.exists():
                failures.append(finding(path, f"Relative link target does not exist: {raw}", number))
                continue
            if fragment and target.is_file() and target.suffix.lower() == ".md":
                if target not in slug_cache:
                    try:
                        slug_cache[target] = _github_slugs(target.read_text(encoding="utf-8"))
                    except (OSError, UnicodeError):
                        slug_cache[target] = set()
                if fragment.casefold() not in slug_cache[target]:
                    failures.append(finding(path, f"Markdown heading does not exist: {raw}", number))
    return rule_result(
        "markdown-links",
        "Markdown relative links",
        failures,
        f"Resolved {checked} relative links and anchors",
    )


def check_text_integrity(repo: Repository) -> dict[str, object]:
    failures: list[dict[str, object]] = []
    checked = 0
    conflict = re.compile(r"^(?:<{7}|={7}|>{7})(?:\s|$)", re.MULTILINE)
    private_key_marker = "-----BEGIN " + "PRIVATE KEY-----"
    github_token_prefix = "gh" + "p_"
    aws_key = re.compile(r"AKIA[0-9A-Z]{16}")
    for path in repo.files:
        if not is_text_file(path):
            continue
        checked += 1
        try:
            data = repo.bytes(path)
        except OSError as error:
            failures.append(finding(path, f"Tracked text file cannot be read: {error}"))
            continue
        if b"\0" in data:
            failures.append(finding(path, "Tracked text file contains a NUL byte."))
            continue
        try:
            text = repo.text(path)
        except UnicodeError as error:
            failures.append(finding(path, f"Tracked text file has invalid encoding: {error}"))
            continue
        for match in conflict.finditer(text):
            failures.append(finding(path, "Merge-conflict marker is present.", line_number(text, match.start())))
        for marker, label in (
            (private_key_marker, "private-key material"),
            (github_token_prefix, "GitHub token material"),
        ):
            offset = text.find(marker)
            if offset >= 0:
                failures.append(finding(path, f"Possible {label} is tracked.", line_number(text, offset)))
        match = aws_key.search(text)
        if match:
            failures.append(finding(path, "Possible AWS access key is tracked.", line_number(text, match.start())))
    return rule_result(
        "text-integrity",
        "Text integrity and secret markers",
        failures,
        f"Validated {checked} tracked text files",
    )


def check_forbidden_artifacts(repo: Repository) -> dict[str, object]:
    failures: list[dict[str, object]] = []
    for path in repo.files:
        pure = PurePosixPath(path)
        lower = path.casefold()
        name = pure.name.casefold()
        suffix = pure.suffix.casefold()
        if name.startswith("~$") or suffix in {".laccdb", ".ldb"}:
            failures.append(finding(path, "Office lock file must not be tracked."))
        if suffix in OFFICE_BINARY_SUFFIXES:
            failures.append(finding(path, "Generated Office binary must be a release asset, not tracked source."))
        if name == ".env" or (name.startswith(".env.") and name != ".env.example"):
            failures.append(finding(path, "Local environment or secret file must not be tracked."))
        if suffix in SECRET_SUFFIXES:
            failures.append(finding(path, "Private key or certificate material must not be tracked."))
        components = {component.casefold() for component in pure.parts}
        if components.intersection({"private", "private-review", "review-private"}):
            failures.append(finding(path, "Private review material must not be tracked."))
        if any(token in lower for token in ("private_review", "confidential-review", "internal-review")):
            failures.append(finding(path, "Private review material must not be tracked."))
    return rule_result(
        "forbidden-artifacts",
        "Forbidden tracked artifacts",
        failures,
        "No Office locks, generated Office binaries, secret files, or private review material are tracked",
    )


def _has_bare_lf(data: bytes) -> bool:
    return bool(re.search(rb"(?<!\r)\n", data))


def _has_bare_cr(data: bytes) -> bool:
    return bool(re.search(rb"\r(?!\n)", data))


def check_line_endings(repo: Repository) -> dict[str, object]:
    failures: list[dict[str, object]] = []
    checked = 0
    for path in repo.files:
        if not is_text_file(path):
            continue
        pure = PurePosixPath(path)
        suffix = pure.suffix.casefold()
        if suffix not in CROSS_PLATFORM_SUFFIXES | WINDOWS_TEXT_SUFFIXES and pure.name not in TEXT_NAMES:
            continue
        checked += 1
        try:
            data = repo.bytes(path)
        except OSError:
            continue
        if data.startswith(b"\xef\xbb\xbf"):
            failures.append(finding(path, "UTF-8 BOM is not permitted by the repository encoding policy."))
        if data and not data.endswith(b"\n"):
            failures.append(finding(path, "Text file must end with a newline."))
        if suffix in WINDOWS_TEXT_SUFFIXES:
            if _has_bare_lf(data) or _has_bare_cr(data):
                failures.append(finding(path, "Windows/VBA source must use CRLF line endings only."))
        else:
            if b"\r" in data:
                failures.append(finding(path, "Cross-platform text must use LF line endings only."))
    return rule_result(
        "line-endings-encoding",
        "Line endings and encoding",
        failures,
        f"Validated policy for {checked} tracked text files",
    )


def check_label_manifest(repo: Repository) -> dict[str, object]:
    path = ".github/labels.json"
    failures: list[dict[str, object]] = []
    try:
        document = json.loads(repo.text(path))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        return rule_result(
            "label-manifest",
            "Canonical label manifest",
            [finding(path, f"Cannot load label manifest: {error}", getattr(error, "lineno", None))],
            "",
        )
    if not isinstance(document, dict):
        failures.append(finding(path, "Manifest root must be a JSON object."))
        labels: list[object] = []
    else:
        if document.get("version") != 1:
            failures.append(finding(path, "Manifest version must be 1."))
        labels_value = document.get("labels")
        if not isinstance(labels_value, list):
            failures.append(finding(path, "Manifest labels must be an array."))
            labels = []
        else:
            labels = labels_value
    if len(labels) != 23:
        failures.append(finding(path, "Manifest must contain exactly 23 labels."))
    seen: dict[str, int] = {}
    names: list[str] = []
    expected_keys = {"name", "color", "description"}
    for index, label in enumerate(labels):
        location = f"labels[{index}]"
        if not isinstance(label, dict):
            failures.append(finding(path, f"{location} must be an object."))
            continue
        if set(label) != expected_keys:
            failures.append(finding(path, f"{location} must contain exactly name, color, and description."))
        name = label.get("name")
        color = label.get("color")
        description = label.get("description")
        if not isinstance(name, str) or not name or name != name.strip() or len(name) > 50 or "\n" in name or "\r" in name:
            failures.append(finding(path, f"{location}.name must be a trimmed single-line string of 1-50 characters."))
        else:
            folded = name.casefold()
            if folded in seen:
                failures.append(finding(path, f"{location}.name duplicates labels[{seen[folded]}] case-insensitively."))
            else:
                seen[folded] = index
            names.append(name)
        if not isinstance(color, str) or not re.fullmatch(r"[0-9A-F]{6}", color):
            failures.append(finding(path, f"{location}.color must be six uppercase hexadecimal characters without #."))
        if not isinstance(description, str) or len(description) > 100 or "\n" in str(description) or "\r" in str(description):
            failures.append(finding(path, f"{location}.description must be a single-line string of at most 100 characters."))
    if names != sorted(names, key=lambda value: (value.casefold(), value)):
        failures.append(finding(path, "Labels must be sorted case-insensitively by name."))
    return rule_result(
        "label-manifest",
        "Canonical label manifest",
        failures,
        "Version 1 manifest contains 23 unique, sorted labels",
    )


def check_workflow_actions(repo: Repository) -> dict[str, object]:
    failures: list[dict[str, object]] = []
    checked = 0
    uses_line = re.compile(r"^\s*(?:-\s*)?uses:\s*([^\s#]+)(?:\s+#\s*(.+?))?\s*$")
    full_sha = re.compile(r"^[0-9a-f]{40}$")
    version_comment = re.compile(r"^v\d+\.\d+\.\d+(?:[-+][0-9A-Za-z.-]+)?$")
    for path in repo.files:
        pure = PurePosixPath(path)
        if not str(pure).startswith(".github/workflows/") or pure.suffix.lower() not in {".yml", ".yaml"}:
            continue
        try:
            text = repo.text(path)
        except (OSError, UnicodeError):
            continue
        for number, line in enumerate(text.splitlines(), start=1):
            if "uses:" not in line:
                continue
            match = uses_line.match(line)
            if not match:
                failures.append(finding(path, "Action reference cannot be parsed.", number))
                continue
            reference, comment = match.groups()
            if reference.startswith("./"):
                continue
            checked += 1
            if "@" not in reference:
                failures.append(finding(path, "External action must include an immutable revision.", number))
                continue
            action, revision = reference.rsplit("@", 1)
            if action.startswith("docker://") or not full_sha.fullmatch(revision):
                failures.append(finding(path, "External action must be pinned to a full lowercase 40-character commit SHA.", number))
            if not comment or not version_comment.fullmatch(comment.strip()):
                failures.append(finding(path, "Pinned action must include an audited semantic-version comment such as '# v1.2.3'.", number))
    return rule_result(
        "workflow-actions",
        "Immutable workflow actions",
        failures,
        f"Validated {checked} external action references",
    )


def check_git_diff(repo: Repository) -> dict[str, object]:
    completed = repo._git("diff", "--check", "HEAD", "--", check=False)
    failures: list[dict[str, object]] = []
    if completed.returncode != 0:
        detail = (completed.stdout + completed.stderr).strip() or "git diff --check failed"
        failures.append(finding(".", detail))
    return rule_result("git-diff-check", "Git whitespace check", failures, "git diff --check passes")


def check_vba_option_explicit(repo: Repository) -> dict[str, object]:
    failures: list[dict[str, object]] = []
    checked = 0
    declaration = re.compile(r"^\s*Option\s+Explicit\b", re.IGNORECASE | re.MULTILINE)
    for path in repo.files:
        if PurePosixPath(path).suffix.casefold() not in {".bas", ".cls", ".frm"}:
            continue
        checked += 1
        try:
            text = repo.text(path)
        except (OSError, UnicodeError):
            continue
        if not declaration.search(text):
            failures.append(finding(path, "VBA source must declare Option Explicit."))
    return rule_result(
        "vba-option-explicit",
        "VBA explicit declarations",
        failures,
        f"All {checked} VBA source files declare Option Explicit",
    )


def check_vba_export_header(repo: Repository) -> dict[str, object]:
    """Validate the VBE export header of every tracked VBA component.

    A VBE export names its component through `Attribute VB_Name`. A standard
    module carries that declaration on its first line; a class or form export
    carries a `VERSION`/`BEGIN` header block first, so the declaration is
    accepted anywhere in the leading header region for those suffixes. VBA
    components share one flat project namespace, so the declared name must be
    unique across the whole tracked tree, not merely within a directory.

    Procedure-level attributes are rejected. `Application.MacroOptions` is the
    single description mechanism for this project, and hidden procedure
    attributes would silently compete with it.
    """
    failures: list[dict[str, object]] = []
    declaration = re.compile(r'^Attribute VB_Name = "([^"]*)"$')
    identifier = re.compile(r"^[A-Za-z][A-Za-z0-9_]*$")
    declared: dict[str, str] = {}
    checked = 0
    for path in repo.files:
        pure = PurePosixPath(path)
        suffix = pure.suffix.casefold()
        if suffix not in VBA_SOURCE_SUFFIXES:
            continue
        checked += 1
        try:
            lines = repo.text(path).splitlines()
        except (OSError, UnicodeError) as error:
            failures.append(finding(path, f"VBA source cannot be read: {error}"))
            continue
        for number, line in enumerate(lines, start=1):
            if line.startswith(VBA_PROCEDURE_ATTRIBUTE_PREFIXES):
                failures.append(
                    finding(
                        path,
                        "Procedure-level VBA attribute is not permitted; function and "
                        "argument descriptions are owned by the MacroOptions manifest.",
                        number,
                    )
                )
        matches = [
            (number, match.group(1))
            for number, line in enumerate(lines, start=1)
            if (match := declaration.match(line))
        ]
        if not matches:
            failures.append(
                finding(
                    path,
                    'VBA export must declare Attribute VB_Name = "<ComponentName>" '
                    "in the canonical VBE form.",
                )
            )
            continue
        if len(matches) > 1:
            failures.append(
                finding(path, "VBA export declares Attribute VB_Name more than once.", matches[1][0])
            )
        number, name = matches[0]
        if suffix == ".bas":
            if number != 1:
                failures.append(
                    finding(path, "Attribute VB_Name must be the first line of a module export.", number)
                )
        elif number > VBA_HEADER_SCAN_LINES:
            failures.append(
                finding(
                    path,
                    "Attribute VB_Name must appear in the leading export header block.",
                    number,
                )
            )
        stem = pure.stem
        if name != stem:
            failures.append(
                finding(
                    path,
                    f"Declared component name {name!r} does not match the file name {stem!r}.",
                    number,
                )
            )
        if not identifier.match(name):
            failures.append(
                finding(path, f"Component name {name!r} is not a legal VBA identifier.", number)
            )
        elif len(name) > VBA_COMPONENT_NAME_LIMIT:
            failures.append(
                finding(
                    path,
                    f"Component name {name!r} exceeds the {VBA_COMPONENT_NAME_LIMIT}-character VBA limit.",
                    number,
                )
            )
        # VBA component names are case-insensitive, so uniqueness folds case even
        # though the file-stem match above stays case-sensitive for export fidelity.
        previous = declared.get(name.casefold())
        if previous is not None:
            failures.append(
                finding(
                    path,
                    f"Component name {name!r} collides with {previous}; VBA component names are "
                    "unique across the whole project and are compared without regard to case.",
                    number,
                )
            )
        else:
            declared[name.casefold()] = path
    return rule_result(
        "vba-export-header",
        "VBA export headers and component names",
        failures,
        f"All {checked} VBA source files carry a unique matching Attribute VB_Name",
    )


def _vba_sources(repo: Repository) -> list[tuple[str, str]]:
    """Return (path, text) for every tracked VBA source, skipping unreadable files."""
    sources: list[tuple[str, str]] = []
    for path in repo.files:
        if PurePosixPath(path).suffix.casefold() not in VBA_SOURCE_SUFFIXES:
            continue
        try:
            sources.append((path, repo.text(path)))
        except (OSError, UnicodeError):
            continue
    return sources


def _strip_vba_comments(text: str) -> str:
    """Drop whole-line VBA comments so documentation never satisfies a code rule."""
    return "\n".join(
        line for line in text.splitlines() if not line.lstrip().startswith("'")
    )


def _join_vba_continuations(text: str) -> list[tuple[int, str]]:
    """Return (first line number, logical line) for each continued VBA statement.

    A declaration split across physical lines with a trailing underscore is one
    statement, so a rule that reads a signature has to see it whole.
    """
    logical: list[tuple[int, str]] = []
    buffer = ""
    start = 0
    for number, line in enumerate(text.splitlines(), start=1):
        if not buffer:
            start = number
        stripped = line.rstrip()
        if stripped.endswith(" _"):
            buffer += stripped[:-1]
            continue
        logical.append((start, buffer + stripped))
        buffer = ""
    if buffer:
        logical.append((start, buffer))
    return logical


def _vba_call_arguments(code: str, name: str) -> Iterable[tuple[int, list[str]]]:
    """Yield (offset, arguments) for each call to `name` in already-stripped code.

    Arguments are split on commas at depth one, so a nested call such as
    Month(DateIn) stays inside the argument that contains it.
    """
    opening = re.compile(rf"\b{re.escape(name)}\s*\(", re.IGNORECASE)
    for match in opening.finditer(code):
        index = match.end()
        depth = 1
        argument = ""
        arguments: list[str] = []
        while index < len(code):
            char = code[index]
            if char == "(":
                depth += 1
            elif char == ")":
                depth -= 1
                if depth == 0:
                    break
            if depth == 1 and char == ",":
                arguments.append(argument.strip())
                argument = ""
            else:
                argument += char
            index += 1
        if depth == 0:
            arguments.append(argument.strip())
            yield match.start(), arguments


def check_vba_day_zero_idiom(repo: Repository) -> dict[str, object]:
    """No live code builds a date with a DateSerial day argument of zero.

    The idiom returns the last day of the preceding month, which is why it is
    tempting for month-end work, but it is not total over the supported window:
    it raises for month 13 of year 9999 and for month 1 of the floor year. The
    failure is invisible from a worksheet because a defensive handler turns it
    into an ordinary error value. Whole-line comments are stripped first, so the
    idiom may still be named in a module header that explains why it is absent.
    """
    failures: list[dict[str, object]] = []
    scanned = 0
    for path, text in _vba_sources(repo):
        code = _strip_vba_comments(text)
        scanned += 1
        for offset, arguments in _vba_call_arguments(code, VBA_DAY_ZERO_FUNCTION):
            if len(arguments) != VBA_DAY_ZERO_ARGUMENTS:
                continue
            if arguments[-1] != "0":
                continue
            failures.append(
                finding(
                    path,
                    "DateSerial with a day argument of zero is not total over the "
                    "supported window; read month length from "
                    "KPR_Core_Dates.DaysInMonth instead.",
                    line_number(code, offset),
                )
            )
    return rule_result(
        "vba-day-zero-idiom",
        "VBA day-zero DateSerial idiom",
        failures,
        f"None of the {scanned} VBA sources build a date with a zero day argument",
    )


def check_vba_facade_return_type(repo: Repository) -> dict[str, object]:
    """Every supported worksheet function is declared As Variant.

    A worksheet-facing function reports failure with a native Excel error value,
    and only Variant can carry one. A narrower return type compiles and then
    raises at run time on the first rejected input, which the caller sees as a
    generic failure rather than the documented error.
    """
    failures: list[dict[str, object]] = []
    declaration = re.compile(
        r"^\s*Public\s+Function\s+(KPR_Dates_\w+)\s*\(.*\)\s*As\s+(\w+)\s*$",
        re.IGNORECASE,
    )
    checked = 0
    for path, text in _vba_sources(repo):
        for number, statement in _join_vba_continuations(text):
            match = declaration.match(statement)
            if match is None:
                continue
            checked += 1
            if match.group(2).casefold() != VBA_FACADE_RETURN_TYPE:
                failures.append(
                    finding(
                        path,
                        f"Worksheet function {match.group(1)} is declared As "
                        f"{match.group(2)}; a native Excel error value needs Variant.",
                        number,
                    )
                )
    return rule_result(
        "vba-facade-return-type",
        "VBA facade return type",
        failures,
        f"All {checked} public worksheet functions are declared As Variant",
    )


def check_vba_required_members(repo: Repository) -> dict[str, object]:
    """Each architecture component still declares the members its callers use.

    VBA resolves these at compile time on Windows, which this gate cannot run.
    Pinning the names here means a rename or a deletion fails on the runner
    rather than at the first attempt to open the workbook.
    """
    failures: list[dict[str, object]] = []
    member = re.compile(r"^\s*Public\s+(?:Function|Sub|Const)\s+(\w+)", re.IGNORECASE)
    declared: dict[str, tuple[str, set[str]]] = {}
    for path, text in _vba_sources(repo):
        stem = PurePosixPath(path).stem.casefold()
        if stem not in VBA_REQUIRED_MEMBERS:
            continue
        names = set()
        for _, statement in _join_vba_continuations(text):
            match = member.match(statement)
            if match is not None:
                names.add(match.group(1).casefold())
        declared[stem] = (path, names)
    for stem, required in sorted(VBA_REQUIRED_MEMBERS.items()):
        if stem not in declared:
            failures.append(
                finding(
                    f"src/modules/{stem}.bas",
                    "Architecture component is absent; its required public members "
                    "cannot be verified.",
                )
            )
            continue
        path, names = declared[stem]
        missing = sorted(
            name for name in required if name.casefold() not in names
        )
        if missing:
            failures.append(
                finding(
                    path,
                    f"Component no longer declares required public member(s): "
                    f"{', '.join(missing)}.",
                )
            )
    total = sum(len(names) for names in VBA_REQUIRED_MEMBERS.values())
    return rule_result(
        "vba-required-members",
        "VBA required public members",
        failures,
        f"All {total} pinned public members are still declared",
    )


def check_vba_module_visibility(repo: Repository) -> dict[str, object]:
    """Internal core modules are project-private; the worksheet facade is not.

    `Option Private Module` is what keeps a core helper out of the worksheet, the
    Function Wizard and the macro list. Module role is decided by the component
    name prefix, which is a role marker rather than a casing convention: casing
    itself is deliberately not enforced.
    """
    failures: list[dict[str, object]] = []
    checked = 0
    for path, text in _vba_sources(repo):
        stem = PurePosixPath(path).stem.casefold()
        declares = any(
            line.strip().startswith(VBA_PRIVATE_MODULE_DIRECTIVE)
            for line in text.splitlines()
        )
        if stem.startswith(VBA_INTERNAL_MODULE_PREFIX):
            checked += 1
            if not declares:
                failures.append(
                    finding(path, "Internal core module must declare Option Private Module.")
                )
        elif stem.startswith(VBA_FACADE_STEM_PREFIX):
            checked += 1
            if declares:
                failures.append(
                    finding(
                        path,
                        "Worksheet facade must not declare Option Private Module; its members "
                        "have to reach the worksheet.",
                    )
                )
    return rule_result(
        "vba-module-visibility",
        "VBA module visibility",
        failures,
        f"All {checked} role-bearing VBA modules declare the expected visibility",
    )


def check_vba_public_surface(repo: Repository) -> dict[str, object]:
    """Supported worksheet functions live only in a non-private facade module.

    The test is by module role, not by file name casing: a `KPR_Dates_*` public
    function must sit in a module whose stem belongs to the facade family and
    which does not declare `Option Private Module`.
    """
    failures: list[dict[str, object]] = []
    declaration = re.compile(r"^\s*Public\s+Function\s+(KPR_Dates_\w+)", re.IGNORECASE)
    found = 0
    for path, text in _vba_sources(repo):
        stem = PurePosixPath(path).stem.casefold()
        is_private = any(
            line.strip().startswith(VBA_PRIVATE_MODULE_DIRECTIVE)
            for line in text.splitlines()
        )
        is_facade = stem.startswith(VBA_FACADE_STEM_PREFIX) and not is_private
        for number, line in enumerate(text.splitlines(), start=1):
            match = declaration.match(line)
            if match is None:
                continue
            found += 1
            if not is_facade:
                failures.append(
                    finding(
                        path,
                        f"Public worksheet function {match.group(1)} is declared outside the "
                        "facade; supported functions belong in a non-private KPR_Dates_* module.",
                        number,
                    )
                )
    return rule_result(
        "vba-public-surface",
        "VBA public worksheet surface",
        failures,
        f"All {found} public worksheet functions are declared in a facade module",
    )


def check_vba_module_dependencies(repo: Repository) -> dict[str, object]:
    """Enforce the allowed-dependency matrix recorded in the implementation plan.

    Public members of each known component are collected first, then every other
    component's code is scanned for references to them. Comments are stripped, so
    a module header may name a dependency it is not permitted to call.
    """
    failures: list[dict[str, object]] = []
    exported: dict[str, set[str]] = {}
    bodies: dict[str, tuple[str, str]] = {}
    member = re.compile(
        r"^\s*Public\s+(?:Function|Sub|Const|Enum)\s+(\w+)", re.IGNORECASE
    )
    enum_member = re.compile(r"^\s*(KPR_[A-Z0-9_]+)\s*=\s*-?\d+\s*$")
    for path, text in _vba_sources(repo):
        stem = PurePosixPath(path).stem.casefold()
        if stem not in VBA_ALLOWED_DEPENDENCIES:
            continue
        names: set[str] = set()
        inside_enum = False
        for line in text.splitlines():
            stripped = line.strip()
            if re.match(r"^Public\s+Enum\s+\w+", stripped, re.IGNORECASE):
                inside_enum = True
            elif re.match(r"^End\s+Enum\b", stripped, re.IGNORECASE):
                inside_enum = False
            elif inside_enum and (hit := enum_member.match(line)):
                # Enum members are referenced by name across modules exactly as
                # procedures are, so the matrix has to see them too.
                names.add(hit.group(1))
            if (match := member.match(line)) is not None:
                names.add(match.group(1))
        exported[stem] = names
        bodies[stem] = (path, _strip_vba_comments(text))
    for stem, (path, code) in sorted(bodies.items()):
        allowed = VBA_ALLOWED_DEPENDENCIES[stem]
        for other, names in sorted(exported.items()):
            if other == stem or other in allowed:
                continue
            hits = sorted(
                name for name in names if re.search(rf"\b{re.escape(name)}\b", code)
            )
            if hits:
                failures.append(
                    finding(
                        path,
                        f"Module may not depend on {other}; it references "
                        f"{', '.join(hits)}.",
                    )
                )
    return rule_result(
        "vba-module-dependencies",
        "VBA module dependency matrix",
        failures,
        f"All {len(bodies)} architecture modules respect the allowed-dependency matrix",
    )


def check_vba_locale_parsing(repo: Repository) -> dict[str, object]:
    """Production VBA must not ask the host how to read a date.

    `IsDate`, `DateValue` and `CVDate` all consult the host locale, so a build
    that passes in one region can fail in another. The contract admits exactly
    one text form, parsed character by character.
    """
    failures: list[dict[str, object]] = []
    checked = 0
    for path, text in _vba_sources(repo):
        if not path.startswith("src/"):
            continue
        checked += 1
        for number, line in enumerate(text.splitlines(), start=1):
            if line.lstrip().startswith("'"):
                continue
            for name in VBA_FORBIDDEN_PARSE_CALLS:
                if re.search(rf"\b{name}\s*\(", line):
                    failures.append(
                        finding(
                            path,
                            f"Locale-sensitive {name} is not permitted in production parsing.",
                            number,
                        )
                    )
    return rule_result(
        "vba-locale-parsing",
        "VBA locale-independent parsing",
        failures,
        f"None of the {checked} production VBA sources use a locale-sensitive date conversion",
    )


def check_vba_window_constants(repo: Repository) -> dict[str, object]:
    """Pin the two representations of the supported window against each other.

    `KPR_Core_Dates` declares the window as Date literals. `KPR_Core_Parse` may
    not depend on that module, so it restates the same bounds as serials and
    years. Without this rule the duplication could drift silently.
    """
    failures: list[dict[str, object]] = []
    epoch = date(1899, 12, 30)  # Excel serial 0 in the 1900 date system
    literal = re.compile(r"#(\d{1,2})/(\d{1,2})/(\d{4})#")
    number = re.compile(r"=\s*(-?\d+)#?\s*(?:'|$)")
    seen: dict[str, str] = {}
    for path, text in _vba_sources(repo):
        stem = PurePosixPath(path).stem.casefold()
        pins = VBA_WINDOW_PINS.get(stem)
        if pins is None:
            continue
        for line in text.splitlines():
            for name, expected in pins.items():
                if not re.search(rf"\bConst\s+{name}\b", line):
                    continue
                actual: str | None = None
                if (hit := literal.search(line)) is not None:
                    month, day, year = (int(part) for part in hit.groups())
                    actual = date(year, month, day).isoformat()
                elif (hit := number.search(line)) is not None:
                    value = int(hit.group(1))
                    actual = (
                        str(value)
                        if name.endswith("_YEAR")
                        else (epoch + timedelta(days=value)).isoformat()
                    )
                if actual is None:
                    failures.append(finding(path, f"Cannot read the value pinned to {name}."))
                elif actual != expected:
                    failures.append(
                        finding(
                            path,
                            f"{name} resolves to {actual}, but the supported window pins it "
                            f"to {expected}.",
                        )
                    )
                else:
                    seen[name] = actual
    missing = sorted(
        name for pins in VBA_WINDOW_PINS.values() for name in pins if name not in seen
    )
    failures.extend(
        finding("src/modules", f"Window constant {name} is not declared.") for name in missing
    )
    return rule_result(
        "vba-window-constants",
        "VBA supported-window constants",
        failures,
        f"All {len(seen)} window constants agree across the modules that declare them",
    )


def _vba_procedures(text: str) -> list[tuple[str, str, str]]:
    """Return (visibility, name, body) for every Function or Sub in a module."""
    out: list[tuple[str, str, str]] = []
    pattern = re.compile(
        r"^(Public|Private)\s+(?:Function|Sub)\s+(\w+)\b(.*?)^End\s+(?:Function|Sub)\s*$",
        re.IGNORECASE | re.MULTILINE | re.DOTALL,
    )
    for match in pattern.finditer(text):
        raw = match.group(3).splitlines()
        # Drop the remainder of the declaration, including any continuation
        # lines ending in " _", so the body starts at the first real line.
        index = 0
        while index < len(raw) and raw[index].rstrip().endswith("_"):
            index += 1
        body = "\n".join(raw[index + 1 :])
        out.append((match.group(1).casefold(), match.group(2), body))
    return out


def _strip_vba_strings(line: str) -> str:
    """Blank out double-quoted literals so a token inside a string is not a hit."""
    return re.sub(r'"[^"]*"', '""', line)


def _facade_sources(repo: Repository) -> list[tuple[str, str]]:
    return [
        (path, text)
        for path, text in _vba_sources(repo)
        if PurePosixPath(path).stem.casefold().startswith(VBA_FACADE_STEM_PREFIX)
        and not any(
            line.strip().startswith(VBA_PRIVATE_MODULE_DIRECTIVE)
            for line in text.splitlines()
        )
    ]


def check_vba_host_guard(repo: Repository) -> dict[str, object]:
    """Every value-taking public function runs the date-system guard exactly once.

    The guard must precede the first argument resolver or calculation, so a
    1904 host is refused before any serial is interpreted. The diagnostic is
    exempt from the guard but must call the shared classifier exactly once,
    because an identified 1904 workbook has to be reported there, not refused.
    """
    failures: list[dict[str, object]] = []
    checked = 0
    for path, text in _facade_sources(repo):
        for visibility, name, body in _vba_procedures(text):
            if visibility != "public" or not name.casefold().startswith(VBA_PUBLIC_FUNCTION_PREFIX):
                continue
            code = _strip_vba_comments(body)
            checked += 1
            if name == VBA_HOST_DIAGNOSTIC:
                calls = len(re.findall(rf"\b{VBA_HOST_CLASSIFIER}\s*\(", code))
                if calls != 1:
                    failures.append(
                        finding(path, f"{name} must call {VBA_HOST_CLASSIFIER} exactly once; found {calls}.")
                    )
                if re.search(rf"\b{VBA_HOST_GUARD}\s*\(", code):
                    failures.append(
                        finding(path, f"{name} must not call {VBA_HOST_GUARD}; it reports 1904 rather than refusing it.")
                    )
                continue
            guards = [m.start() for m in re.finditer(rf"\b{VBA_HOST_GUARD}\s*\(", code)]
            if len(guards) != 1:
                failures.append(
                    finding(path, f"{name} must call {VBA_HOST_GUARD} exactly once; found {len(guards)}.")
                )
                continue
            first_resolver = min(
                (m.start() for r in VBA_ARGUMENT_RESOLVERS for m in re.finditer(rf"\b{r}\s*\(", code)),
                default=None,
            )
            if first_resolver is not None and first_resolver < guards[0]:
                failures.append(
                    finding(path, f"{name} resolves an argument before {VBA_HOST_GUARD} runs.")
                )
    return rule_result(
        "vba-host-guard",
        "VBA date-system guard placement",
        failures,
        f"All {checked} public worksheet functions run the date-system guard as required",
    )


def check_vba_volatile_scope(repo: Repository) -> dict[str, object]:
    """Only the host diagnostic is volatile, and volatility is its first statement.

    A zero-argument function has no precedents and would otherwise be
    evaluated once and never again. Every date calculation must stay
    non-volatile, because volatility is contagious and would make a model
    recalculate constantly.
    """
    failures: list[dict[str, object]] = []
    total = 0
    for path, text in _vba_sources(repo):
        if not path.startswith("src/"):
            continue
        for visibility, name, body in _vba_procedures(text):
            # Declarations are not statements: Dim, Const and Static lines are
            # skipped when locating the first executable line.
            code_lines = [
                l
                for l in body.splitlines()
                if l.strip()
                and not l.lstrip().startswith("'")
                and not re.match(r"^\s*(Dim|Const|Static)\b", l, re.IGNORECASE)
            ]
            hits = [l for l in code_lines if VBA_VOLATILE_CALL in _strip_vba_strings(l)]
            total += len(hits)
            if name == VBA_HOST_DIAGNOSTIC:
                if len(hits) != 1:
                    failures.append(
                        finding(path, f"{name} must call {VBA_VOLATILE_CALL} exactly once; found {len(hits)}.")
                    )
                elif not code_lines or VBA_VOLATILE_CALL not in _strip_vba_strings(code_lines[0]):
                    failures.append(
                        finding(path, f"{VBA_VOLATILE_CALL} must be the first executable statement of {name}.")
                    )
            elif hits:
                failures.append(
                    finding(path, f"{name} must not be volatile; only {VBA_HOST_DIAGNOSTIC} may call {VBA_VOLATILE_CALL}.")
                )
    return rule_result(
        "vba-volatile-scope",
        "VBA volatility scope",
        failures,
        f"{VBA_VOLATILE_CALL} appears exactly once, first in {VBA_HOST_DIAGNOSTIC}",
    )


def check_vba_workbook_fallback(repo: Repository) -> dict[str, object]:
    """No active-workbook fallback on the facade's host-resolution path.

    The caller's own workbook is the only authority on its date system. The
    scope is the three host procedures, not all of src/.
    """
    failures: list[dict[str, object]] = []
    checked = 0
    for path, text in _facade_sources(repo):
        for visibility, name, body in _vba_procedures(text):
            if name not in VBA_HOST_PATH_PROCEDURES:
                continue
            checked += 1
            for number, line in enumerate(body.splitlines(), start=1):
                if line.lstrip().startswith("'"):
                    continue
                probe = _strip_vba_strings(line)
                for token in VBA_WORKBOOK_FALLBACKS:
                    if re.search(rf"\b{token}\b", probe):
                        failures.append(
                            finding(path, f"{name} references {token}; the caller's workbook is the only date-system authority.")
                        )
    return rule_result(
        "vba-no-workbook-fallback",
        "VBA no active-workbook fallback",
        failures,
        f"All {checked} host-resolution procedures avoid ActiveWorkbook, ThisWorkbook and ActiveSheet",
    )


CHECKS: tuple[Callable[[Repository], dict[str, object]], ...] = (
    check_required_files,
    check_stale_identity,
    check_structured_data,
    check_markdown_links,
    check_text_integrity,
    check_forbidden_artifacts,
    check_line_endings,
    check_label_manifest,
    check_workflow_actions,
    check_git_diff,
    check_vba_option_explicit,
    check_vba_export_header,
    check_vba_module_visibility,
    check_vba_public_surface,
    check_vba_module_dependencies,
    check_vba_locale_parsing,
    check_vba_window_constants,
    check_vba_host_guard,
    check_vba_volatile_scope,
    check_vba_workbook_fallback,
    check_vba_required_members,
    check_vba_facade_return_type,
    check_vba_day_zero_idiom,
)


def check_repository(root: Path) -> dict[str, object]:
    repo = Repository(root)
    rules = [check(repo) for check in CHECKS]
    failed = sum(rule["status"] == "fail" for rule in rules)
    passed = len(rules) - failed
    return {
        "schema_version": SCHEMA_VERSION,
        "tool": TOOL_NAME,
        "repository": os.environ.get("GITHUB_REPOSITORY", "danielep71/KPR"),
        "commit": os.environ.get("GITHUB_SHA") or repo.commit(),
        "generated_at_utc": dt.datetime.now(dt.timezone.utc).isoformat().replace("+00:00", "Z"),
        "scope_note": SCOPE_NOTE,
        "status": "fail" if failed else "pass",
        "counts": {"rules": len(rules), "passed": passed, "failed": failed},
        "rules": rules,
    }


def render_summary(report: dict[str, object]) -> str:
    status = str(report["status"]).upper()
    counts = report["counts"]
    lines = [
        "# KPR static repository checks",
        "",
        f"**Result: {status}** — {counts['passed']} passed, {counts['failed']} failed.",
        "",
        str(report["scope_note"]),
        "",
        "| Rule | Result | Summary |",
        "|---|---:|---|",
    ]
    for rule in report["rules"]:
        symbol = "✅" if rule["status"] == "pass" else "❌"
        lines.append(f"| `{rule['id']}` — {rule['title']} | {symbol} | {rule['summary']} |")
    for rule in report["rules"]:
        if rule["status"] != "fail":
            continue
        lines.extend(["", f"## Failed: `{rule['id']}`", ""])
        for item in rule["findings"]:
            location = str(item["path"])
            if "line" in item:
                location += f":{item['line']}"
            message = str(item["message"]).replace("\n", " ")
            lines.append(f"- `{location}` — {message}")
    return "\n".join(lines) + "\n"


def write_report(report: dict[str, object], output: Path, summary: Path | None) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(report, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    markdown = render_summary(report)
    if summary is not None:
        summary.parent.mkdir(parents=True, exist_ok=True)
        with summary.open("a", encoding="utf-8") as stream:
            stream.write(markdown)


def _run_git(root: Path, *arguments: str) -> None:
    subprocess.run(
        ["git", "-C", str(root), *arguments],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )


def _initialize_fixture(root: Path) -> None:
    _run_git(root, "init", "-q")
    _run_git(root, "config", "user.name", "KPR Static Checks")
    _run_git(root, "config", "user.email", "static-checks@example.invalid")
    _run_git(root, "config", "core.autocrlf", "false")
    _run_git(root, "add", "-A", "-f")
    _run_git(root, "commit", "-q", "-m", "fixture")


def _fixture(source: Path, target: Path) -> None:
    ignored = shutil.ignore_patterns(".git", "__pycache__", "test-results", "test-output")
    shutil.copytree(source, target, ignore=ignored)
    _initialize_fixture(target)


def _positive_fixture(root: Path) -> None:
    root.mkdir(parents=True)

    def write(relative: str, content: str, newline: str = "\n") -> None:
        path = root / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(content, encoding="utf-8", newline=newline)

    write(".editorconfig", "root = true\n")
    write(".gitattributes", "* text=auto\n*.bas text eol=crlf\n")
    write(".gitignore", "test-results/\n")
    for relative in (
        ".github/ISSUE_TEMPLATE/bug_report.md",
        ".github/ISSUE_TEMPLATE/feature_request.md",
        ".github/PULL_REQUEST_TEMPLATE.md",
        "CHANGELOG.md",
        "CODE_OF_CONDUCT.md",
        "CONTRIBUTING.md",
        "INSTALLATION.md",
        "SECURITY.md",
    ):
        write(relative, "# KPR\n")
    write("README.md", "# KPR\n\n[Conduct](CODE_OF_CONDUCT.md)\n")
    write("LICENSE", "MIT License\n")
    write(".github/ISSUE_TEMPLATE/config.yml", "blank_issues_enabled: true\n")
    write(".github/scripts/labels-sync.mjs", "// self-test placeholder\n")
    write("assets/social-preview.png", "self-test placeholder\n")
    labels = [
        {"name": f"label-{index:02d}", "color": "123ABC", "description": "self-test"}
        for index in range(1, 24)
    ]
    write(".github/labels.json", json.dumps({"version": 1, "labels": labels}, indent=2) + "\n")
    checkout = "actions/checkout@3d3c42e5aac5ba805825da76410c181273ba90b1 # v7.0.1"
    write(
        ".github/workflows/labels-sync.yml",
        "name: Labels\non:\n  push:\njobs:\n  check:\n    runs-on: ubuntu-latest\n"
        f"    steps:\n      - uses: {checkout}\n",
    )
    write(
        ".github/workflows/static-checks.yml",
        "name: Static\non:\n  push:\njobs:\n  check:\n    runs-on: ubuntu-latest\n"
        f"    steps:\n      - uses: {checkout}\n",
    )
    write("VERSION", "0.0.1\n")
    write("tools/check_repo.py", "# self-test placeholder\n")
    write("docs/DATE_LAYER_CONTRACT.md", "# Contract\n")
    write("docs/VBE_EXPORT.md", "# Export\n")
    for internal in ("KPR_Core_Err", "KPR_Core_Parse", "KPR_Core_Dates", "KPR_Core_Array"):
        body = ""
        if internal.casefold() == "kpr_core_dates":
            body += (
                "Public Const KPR_MIN_DATE As Date = #3/1/1900#\n"
                "Public Const KPR_MAX_DATE As Date = #12/31/9999#\n"
            )
        if internal.casefold() == "kpr_core_parse":
            body += (
                "Private Const KPR_MIN_SERIAL As Double = 61#\n"
                "Private Const KPR_MAX_SERIAL As Double = 2958465#\n"
                "Private Const KPR_MIN_YEAR As Long = 1900\n"
                "Private Const KPR_MAX_YEAR As Long = 9999\n"
            )
        for name in sorted(VBA_REQUIRED_MEMBERS[internal.casefold()]):
            if name.startswith("KPR_"):
                continue
            body += f"Public Function {name}() As Variant\nEnd Function\n"
        write(
            f"src/modules/{internal}.bas",
            f'Attribute VB_Name = "{internal}"\nOption Explicit\nOption Private Module\n'
            + body,
            newline="\r\n",
        )
    write(
        "test/modules/KPR_REGRESSION_TESTS.bas",
        'Attribute VB_Name = "KPR_REGRESSION_TESTS"\nOption Explicit\n',
        newline="\r\n",
    )
    write(
        "src/modules/KPR_DATES_DAYS.bas",
        'Attribute VB_Name = "KPR_DATES_DAYS"\nOption Explicit\n'
        "Public Function KPR_Dates_DayOfWeek(ByVal DateIn As Variant) As Variant\n"
        "    Dim FailErr As Variant\n"
        "    If Not PassHostGuard(FailErr) Then Exit Function\n"
        "    If Not TryResolveDate(DateIn, 0, FailErr) Then Exit Function\n"
        "End Function\n"
        "Public Function KPR_Dates_HostDateSystem() As Variant\n"
        "    Dim DateSystem As Long\n"
        "    Application.Volatile True\n"
        "    If TryResolveHostDateSystem(DateSystem, 0) Then KPR_Dates_HostDateSystem = DateSystem\n"
        "End Function\n"
        "Private Function TryResolveHostDateSystem(ByRef DateSystem As Long, ByRef Condition As Long) As Boolean\n"
        "    DateSystem = 1900\n"
        "    TryResolveHostDateSystem = True\n"
        "End Function\n"
        "Private Function PassHostGuard(ByRef ErrOut As Variant) As Boolean\n"
        "    PassHostGuard = True\n"
        "End Function\n"
        "Private Function TryResolveDate(ByVal DateIn As Variant, ByRef ParsedDate As Date, ByRef ErrOut As Variant) As Boolean\n"
        "    TryResolveDate = True\n"
        "End Function\n",
        newline="\r\n",
    )
    _initialize_fixture(root)


def run_self_tests(root: Path) -> None:
    static_workflow = ".github/workflows/static-checks.yml"
    identity_name = "Date" + "Picker"

    def stale_identity(case: Path) -> None:
        readme = case / "README.md"
        readme.write_text(readme.read_text(encoding="utf-8") + f"\nInherited {identity_name}\n", encoding="utf-8")

    def missing_required_dependency(case: Path) -> None:
        _run_git(case, "rm", ".github/scripts/labels-sync.mjs")

    def unpinned_action(case: Path) -> None:
        workflow = case / static_workflow
        text = workflow.read_text(encoding="utf-8")
        text = text.replace(
            "actions/checkout@3d3c42e5aac5ba805825da76410c181273ba90b1 # v7.0.1",
            "actions/checkout@v7 # v7.0.1",
            1,
        )
        workflow.write_text(text, encoding="utf-8")

    def invalid_label_json(case: Path) -> None:
        (case / ".github/labels.json").write_text("{\n", encoding="utf-8")

    def invalid_yaml(case: Path) -> None:
        path = case / ".github/workflows/invalid.yml"
        path.write_text("jobs\n", encoding="utf-8")
        _run_git(case, "add", "-f", str(path.relative_to(case)))

    def invalid_xml(case: Path) -> None:
        path = case / "src/ribbon/customUI14.xml"
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text("<customUI>\n", encoding="utf-8")
        _run_git(case, "add", "-f", str(path.relative_to(case)))

    def broken_markdown(case: Path) -> None:
        readme = case / "README.md"
        readme.write_text(readme.read_text(encoding="utf-8") + "\n[missing](docs/not-here.md)\n", encoding="utf-8")

    module_relative = "src/modules/KPR_DATES_DAYS.bas"

    def rewrite_module(case: Path, transform: Callable[[str], str]) -> None:
        module = case / module_relative
        module.write_text(
            transform(module.read_text(encoding="utf-8")),
            encoding="utf-8",
            newline="\r\n",
        )

    def missing_module_name(case: Path) -> None:
        rewrite_module(case, lambda text: text.split("\n", 1)[1])

    def mismatched_module_name(case: Path) -> None:
        rewrite_module(case, lambda text: text.replace('"KPR_DATES_DAYS"', '"KPR_DATES_DAY"', 1))

    def duplicate_module_name(case: Path) -> None:
        source = case / module_relative
        duplicate = case / "test/modules/KPR_DATES_DAYS.bas"
        duplicate.parent.mkdir(parents=True, exist_ok=True)
        duplicate.write_text(source.read_text(encoding="utf-8"), encoding="utf-8", newline="\r\n")
        _run_git(case, "add", "-f", str(duplicate.relative_to(case)))

    def case_only_duplicate_name(case: Path) -> None:
        # VBA component names are case-insensitive, so these are one component.
        duplicate = case / "test/modules/kpr_dates_days.bas"
        duplicate.parent.mkdir(parents=True, exist_ok=True)
        duplicate.write_text(
            'Attribute VB_Name = "kpr_dates_days"\nOption Explicit\n',
            encoding="utf-8",
            newline="\r\n",
        )
        _run_git(case, "add", "-f", str(duplicate.relative_to(case)))

    def locale_parsing(case: Path) -> None:
        core = case / "src/modules/KPR_Core_Parse.bas"
        core.write_text(
            core.read_text(encoding="utf-8")
            + "Public Function Probe(ByVal S As String) As Boolean\n"
            + "    Probe = IsNumeric(S)\n"
            + "End Function\n",
            encoding="utf-8",
            newline="\r\n",
        )

    def drifted_window_constant(case: Path) -> None:
        core = case / "src/modules/KPR_Core_Parse.bas"
        core.write_text(
            core.read_text(encoding="utf-8").replace("= 61#", "= 60#", 1),
            encoding="utf-8",
            newline="\r\n",
        )

    def missing_host_guard(case: Path) -> None:
        facade = case / "src/modules/KPR_DATES_DAYS.bas"
        facade.write_text(
            facade.read_text(encoding="utf-8").replace(
                "    If Not PassHostGuard(FailErr) Then Exit Function\n", "", 1
            ),
            encoding="utf-8",
            newline="\r\n",
        )

    def guard_after_resolver(case: Path) -> None:
        facade = case / "src/modules/KPR_DATES_DAYS.bas"
        text = facade.read_text(encoding="utf-8")
        guard = "    If Not PassHostGuard(FailErr) Then Exit Function\n"
        resolve = "    If Not TryResolveDate(DateIn, 0, FailErr) Then Exit Function\n"
        facade.write_text(
            text.replace(guard + resolve, resolve + guard, 1), encoding="utf-8", newline="\r\n"
        )

    def stray_volatile(case: Path) -> None:
        facade = case / "src/modules/KPR_DATES_DAYS.bas"
        facade.write_text(
            facade.read_text(encoding="utf-8").replace(
                "    Dim FailErr As Variant\n",
                "    Dim FailErr As Variant\n    Application.Volatile True\n",
                1,
            ),
            encoding="utf-8",
            newline="\r\n",
        )

    def workbook_fallback(case: Path) -> None:
        facade = case / "src/modules/KPR_DATES_DAYS.bas"
        facade.write_text(
            facade.read_text(encoding="utf-8").replace(
                "    DateSystem = 1900\n",
                "    DateSystem = IIf(ActiveWorkbook.Date1904, 1904, 1900)\n",
                1,
            ),
            encoding="utf-8",
            newline="\r\n",
        )

    def missing_private_module(case: Path) -> None:
        core = case / "src/modules/KPR_Core_Parse.bas"
        core.write_text(
            core.read_text(encoding="utf-8").replace("Option Private Module\n", ""),
            encoding="utf-8",
            newline="\r\n",
        )

    def public_function_outside_facade(case: Path) -> None:
        core = case / "src/modules/KPR_Core_Dates.bas"
        core.write_text(
            core.read_text(encoding="utf-8")
            + "Public Function KPR_Dates_EndOfYear() As Variant\nEnd Function\n",
            encoding="utf-8",
            newline="\r\n",
        )

    def forbidden_dependency(case: Path) -> None:
        core = case / "src/modules/KPR_Core_Err.bas"
        core.write_text(
            core.read_text(encoding="utf-8")
            + "Public Function Probe() As Variant\n"
            + "    Probe = TryParseDateScalar(0, 0)\n"
            + "End Function\n",
            encoding="utf-8",
            newline="\r\n",
        )
        parse = case / "src/modules/KPR_Core_Parse.bas"
        parse.write_text(
            parse.read_text(encoding="utf-8")
            + "Public Function TryParseDateScalar() As Variant\nEnd Function\n",
            encoding="utf-8",
            newline="\r\n",
        )

    def procedure_attribute(case: Path) -> None:
        rewrite_module(
            case,
            lambda text: text.replace(
                "Option Explicit\n",
                'Option Explicit\nAttribute VB_Description = "self-test"\n',
                1,
            ),
        )

    def day_zero_idiom(case: Path) -> None:
        core = case / "src/modules/KPR_Core_Dates.bas"
        core.write_text(
            core.read_text(encoding="utf-8")
            + "Public Function MonthEndProbe(ByVal Y As Long, ByVal M As Long) As Date\n"
            + "    MonthEndProbe = DateSerial(Y, M + 1, 0)\n"
            + "End Function\n",
            encoding="utf-8",
            newline="\r\n",
        )

    def narrow_facade_return(case: Path) -> None:
        rewrite_module(
            case,
            lambda text: text.replace(
                "Public Function KPR_Dates_DayOfWeek(ByVal DateIn As Variant) As Variant",
                "Public Function KPR_Dates_DayOfWeek(ByVal DateIn As Variant) As Long",
                1,
            ),
        )

    def missing_required_member(case: Path) -> None:
        core = case / "src/modules/KPR_Core_Dates.bas"
        core.write_text(
            core.read_text(encoding="utf-8").replace(
                "Public Function DaysInMonth() As Variant\nEnd Function\n", "", 1
            ),
            encoding="utf-8",
            newline="\r\n",
        )

    scenarios: tuple[tuple[str, str, Callable[[Path], None]], ...] = (
        ("stale identity", "stale-identity", stale_identity),
        ("missing required dependency", "required-files", missing_required_dependency),
        ("unpinned action", "workflow-actions", unpinned_action),
        ("invalid label JSON", "label-manifest", invalid_label_json),
        ("invalid YAML", "structured-data", invalid_yaml),
        ("invalid Ribbon XML", "structured-data", invalid_xml),
        ("broken Markdown link", "markdown-links", broken_markdown),
        ("missing module name", "vba-export-header", missing_module_name),
        ("mismatched module name", "vba-export-header", mismatched_module_name),
        ("duplicate module name", "vba-export-header", duplicate_module_name),
        ("case-only duplicate name", "vba-export-header", case_only_duplicate_name),
        ("procedure-level attribute", "vba-export-header", procedure_attribute),
        ("missing Option Private Module", "vba-module-visibility", missing_private_module),
        ("public function outside facade", "vba-public-surface", public_function_outside_facade),
        ("forbidden module dependency", "vba-module-dependencies", forbidden_dependency),
        ("locale-sensitive parsing", "vba-locale-parsing", locale_parsing),
        ("drifted window constant", "vba-window-constants", drifted_window_constant),
        ("missing host guard", "vba-host-guard", missing_host_guard),
        ("guard after resolver", "vba-host-guard", guard_after_resolver),
        ("stray volatile call", "vba-volatile-scope", stray_volatile),
        ("active-workbook fallback", "vba-no-workbook-fallback", workbook_fallback),
        ("day-zero DateSerial idiom", "vba-day-zero-idiom", day_zero_idiom),
        ("narrow facade return type", "vba-facade-return-type", narrow_facade_return),
        ("missing required member", "vba-required-members", missing_required_member),
    )

    with tempfile.TemporaryDirectory(prefix="kpr-static-self-test-") as temporary:
        base = Path(temporary)
        positive_root = base / "positive"
        _positive_fixture(positive_root)
        positive = check_repository(positive_root)
        if positive["status"] != "pass":
            failed = ", ".join(rule["id"] for rule in positive["rules"] if rule["status"] == "fail")
            raise RuntimeError(f"Positive self-test fixture failed: {failed}")
        for index, (name, expected_rule, degrade) in enumerate(scenarios):
            case = base / f"case-{index}"
            _fixture(positive_root, case)
            degrade(case)
            report = check_repository(case)
            failed_ids = {rule["id"] for rule in report["rules"] if rule["status"] == "fail"}
            if report["status"] != "fail" or expected_rule not in failed_ids:
                raise RuntimeError(f"Degraded self-test '{name}' did not fail {expected_rule}.")
            summary = render_summary(report)
            expected = next(rule for rule in report["rules"] if rule["id"] == expected_rule)
            if expected_rule not in summary or not expected["findings"]:
                raise RuntimeError(f"Degraded self-test '{name}' was not fully represented in the summary.")
            serialized = json.dumps(report)
            if expected_rule not in serialized:
                raise RuntimeError(f"Degraded self-test '{name}' was not represented in JSON.")
            print(f"PASS degraded self-test: {name} -> {expected_rule}")
    print(f"PASS positive self-test: {positive['counts']['rules']} rules")


def parse_arguments(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=TOOL_NAME)
    parser.add_argument("--root", type=Path, default=Path.cwd(), help="repository root")
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("test-results/static-checks.json"),
        help="machine-readable report path",
    )
    parser.add_argument("--summary", type=Path, help="Markdown summary file to append")
    parser.add_argument("--self-test", action="store_true", help="run positive and deliberately degraded self-tests")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    arguments = parse_arguments(sys.argv[1:] if argv is None else argv)
    root = arguments.root.resolve()
    try:
        if arguments.self_test:
            run_self_tests(root)
            return 0
        report = check_repository(root)
        output = arguments.output if arguments.output.is_absolute() else root / arguments.output
        summary = arguments.summary
        write_report(report, output, summary)
        print(render_summary(report), end="")
        print(f"JSON report: {output}")
        return 0 if report["status"] == "pass" else 1
    except (OSError, subprocess.SubprocessError, RuntimeError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
