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
    "src/modules/KPR_Dates_Days.bas",
    "tools/check_repo.py",
)

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
    write("src/modules/KPR_Dates_Days.bas", "Option Explicit\n", newline="\r\n")
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

    scenarios: tuple[tuple[str, str, Callable[[Path], None]], ...] = (
        ("stale identity", "stale-identity", stale_identity),
        ("missing required dependency", "required-files", missing_required_dependency),
        ("unpinned action", "workflow-actions", unpinned_action),
        ("invalid label JSON", "label-manifest", invalid_label_json),
        ("invalid YAML", "structured-data", invalid_yaml),
        ("invalid Ribbon XML", "structured-data", invalid_xml),
        ("broken Markdown link", "markdown-links", broken_markdown),
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
