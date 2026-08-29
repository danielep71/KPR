#!/usr/bin/env node

import assert from "node:assert/strict";
import { appendFile, readFile } from "node:fs/promises";
import process from "node:process";

const EXPECTED_VERSION = 1;
const EXPECTED_LABEL_COUNT = 23;
const DEFAULT_MANIFEST = ".github/labels.json";
const API_VERSION = "2022-11-28";

function compareNames(left, right) {
  const leftKey = left.toLowerCase();
  const rightKey = right.toLowerCase();
  if (leftKey < rightKey) return -1;
  if (leftKey > rightKey) return 1;
  return left < right ? -1 : left > right ? 1 : 0;
}

function normalizedLiveLabel(label) {
  return {
    name: String(label.name ?? ""),
    color: String(label.color ?? "").toUpperCase(),
    description: label.description == null ? "" : String(label.description)
  };
}

export function validateManifest(document) {
  const errors = [];

  if (!document || typeof document !== "object" || Array.isArray(document)) {
    throw new Error("Manifest must be a JSON object.");
  }

  if (document.version !== EXPECTED_VERSION) {
    errors.push(`version must be ${EXPECTED_VERSION}`);
  }

  if (!Array.isArray(document.labels)) {
    errors.push("labels must be an array");
  } else if (document.labels.length !== EXPECTED_LABEL_COUNT) {
    errors.push(`labels must contain exactly ${EXPECTED_LABEL_COUNT} entries`);
  }

  const labels = Array.isArray(document.labels) ? document.labels : [];
  const seen = new Map();

  labels.forEach((label, index) => {
    const location = `labels[${index}]`;
    if (!label || typeof label !== "object" || Array.isArray(label)) {
      errors.push(`${location} must be an object`);
      return;
    }

    const keys = Object.keys(label).sort();
    const expectedKeys = ["color", "description", "name"];
    if (JSON.stringify(keys) !== JSON.stringify(expectedKeys)) {
      errors.push(`${location} must contain exactly name, color, and description`);
    }

    if (typeof label.name !== "string") {
      errors.push(`${location}.name must be a string`);
    } else {
      if (label.name.length === 0 || label.name !== label.name.trim()) {
        errors.push(`${location}.name must be non-empty with no surrounding whitespace`);
      }
      if (label.name.length > 50) {
        errors.push(`${location}.name exceeds GitHub's 50-character limit`);
      }
      if (/\r|\n/.test(label.name)) {
        errors.push(`${location}.name must be single-line`);
      }

      const key = label.name.toLowerCase();
      if (seen.has(key)) {
        errors.push(`${location}.name duplicates ${seen.get(key)} case-insensitively`);
      } else {
        seen.set(key, `${location}.name`);
      }
    }

    if (typeof label.color !== "string" || !/^[0-9A-F]{6}$/.test(label.color)) {
      errors.push(`${location}.color must be six uppercase hexadecimal characters without #`);
    }

    if (typeof label.description !== "string") {
      errors.push(`${location}.description must be a string`);
    } else {
      if (label.description.length > 100) {
        errors.push(`${location}.description exceeds GitHub's 100-character limit`);
      }
      if (/\r|\n/.test(label.description)) {
        errors.push(`${location}.description must be single-line`);
      }
    }
  });

  for (let index = 1; index < labels.length; index += 1) {
    const previous = labels[index - 1]?.name;
    const current = labels[index]?.name;
    if (typeof previous === "string" && typeof current === "string" && compareNames(previous, current) >= 0) {
      errors.push(`labels must be sorted case-insensitively by name: ${previous} precedes ${current}`);
      break;
    }
  }

  if (errors.length > 0) {
    throw new Error(`Invalid label manifest:\n- ${errors.join("\n- ")}`);
  }

  return labels.map(label => ({ ...label }));
}

export function planChanges(desiredLabels, liveLabels, { prune = true } = {}) {
  const desired = new Map(desiredLabels.map(label => [label.name.toLowerCase(), label]));
  const live = new Map(liveLabels.map(label => {
    const normalized = normalizedLiveLabel(label);
    return [normalized.name.toLowerCase(), normalized];
  }));
  const changes = [];

  for (const [key, target] of desired) {
    const current = live.get(key);
    if (!current) {
      changes.push({ action: "create", name: target.name, target });
      continue;
    }

    const fields = [];
    if (current.name !== target.name) fields.push("name");
    if (current.color !== target.color) fields.push("color");
    if (current.description !== target.description) fields.push("description");
    if (fields.length > 0) {
      changes.push({
        action: "update",
        name: target.name,
        currentName: current.name,
        fields,
        target
      });
    }
  }

  if (prune) {
    for (const [key, current] of live) {
      if (!desired.has(key)) {
        changes.push({ action: "delete", name: current.name, currentName: current.name });
      }
    }
  }

  const order = { update: 0, create: 1, delete: 2 };
  changes.sort((left, right) => order[left.action] - order[right.action] || compareNames(left.name, right.name));
  return changes;
}

function parseArguments(argv) {
  const parsed = {
    manifest: DEFAULT_MANIFEST,
    mode: null,
    live: null,
    repository: process.env.GITHUB_REPOSITORY || null,
    dryRun: false,
    selfTest: false
  };

  for (let index = 0; index < argv.length; index += 1) {
    const argument = argv[index];
    if (argument === "--dry-run") {
      parsed.dryRun = true;
    } else if (argument === "--self-test") {
      parsed.selfTest = true;
    } else if (["--manifest", "--mode", "--live", "--repository"].includes(argument)) {
      const value = argv[index + 1];
      if (!value || value.startsWith("--")) {
        throw new Error(`${argument} requires a value`);
      }
      parsed[argument.slice(2)] = value;
      index += 1;
    } else {
      throw new Error(`Unknown argument: ${argument}`);
    }
  }

  return parsed;
}

async function loadJson(path) {
  let text;
  try {
    text = await readFile(path, "utf8");
  } catch (error) {
    throw new Error(`Cannot read ${path}: ${error.message}`);
  }

  try {
    return JSON.parse(text);
  } catch (error) {
    throw new Error(`Invalid JSON in ${path}: ${error.message}`);
  }
}

function splitRepository(repository) {
  const parts = String(repository ?? "").split("/");
  if (parts.length !== 2 || parts.some(part => part.length === 0)) {
    throw new Error("repository must be in owner/name form");
  }
  return parts.map(encodeURIComponent);
}

function apiUrl(repository, suffix) {
  const [owner, repo] = splitRepository(repository);
  return `https://api.github.com/repos/${owner}/${repo}${suffix}`;
}

async function githubRequest(repository, token, method, suffix, body) {
  const response = await fetch(apiUrl(repository, suffix), {
    method,
    headers: {
      Accept: "application/vnd.github+json",
      Authorization: `Bearer ${token}`,
      "X-GitHub-Api-Version": API_VERSION,
      "User-Agent": "KPR-labels-sync"
    },
    body: body === undefined ? undefined : JSON.stringify(body)
  });

  if (!response.ok) {
    const detail = await response.text();
    throw new Error(`${method} ${suffix} failed with HTTP ${response.status}: ${detail.slice(0, 500)}`);
  }

  if (response.status === 204) return null;
  return response.json();
}

async function listLiveLabels(repository, token) {
  const labels = [];
  for (let page = 1; ; page += 1) {
    const batch = await githubRequest(repository, token, "GET", `/labels?per_page=100&page=${page}`);
    labels.push(...batch.map(normalizedLiveLabel));
    if (batch.length < 100) break;
  }
  return labels;
}

async function applyChanges(repository, token, changes) {
  for (const change of changes) {
    if (change.action === "update") {
      await githubRequest(
        repository,
        token,
        "PATCH",
        `/labels/${encodeURIComponent(change.currentName)}`,
        {
          new_name: change.target.name,
          color: change.target.color,
          description: change.target.description
        }
      );
    } else if (change.action === "create") {
      await githubRequest(repository, token, "POST", "/labels", change.target);
    } else if (change.action === "delete") {
      await githubRequest(repository, token, "DELETE", `/labels/${encodeURIComponent(change.currentName)}`);
    }
  }
}

function markdownEscape(value) {
  return String(value).replaceAll("|", "\\|").replaceAll("\n", " ");
}

function renderSummary({ mode, desiredCount, changes = [], verified = null, dryRun = false, error = null }) {
  const counts = Object.fromEntries(["create", "update", "delete"].map(action => [
    action,
    changes.filter(change => change.action === action).length
  ]));
  const lines = [
    "# KPR label synchronization",
    "",
    `- Mode: \`${mode}\`${dryRun ? " (dry run)" : ""}`,
    `- Manifest labels: **${desiredCount}**`,
    `- Planned changes: **${changes.length}**`,
    `- Create / update / delete: **${counts.create} / ${counts.update} / ${counts.delete}**`
  ];

  if (verified !== null) {
    lines.push(`- Post-run exact match: **${verified ? "yes" : "no"}**`);
  }

  if (error) {
    lines.push("", "## Failure", "", `\`${markdownEscape(error)}\``);
  } else if (changes.length === 0) {
    lines.push("", "No label changes are required.");
  } else {
    lines.push("", "| Action | Label | Detail |", "|---|---|---|");
    for (const change of changes) {
      const detail = change.action === "update" ? change.fields.join(", ") : change.action === "delete" ? "outside manifest" : "missing live label";
      lines.push(`| ${change.action} | \`${markdownEscape(change.name)}\` | ${markdownEscape(detail)} |`);
    }
  }

  return `${lines.join("\n")}\n`;
}

async function publishSummary(summary) {
  process.stdout.write(summary);
  if (process.env.GITHUB_STEP_SUMMARY) {
    await appendFile(process.env.GITHUB_STEP_SUMMARY, summary, "utf8");
  }
}

function expectFailure(operation, pattern) {
  assert.throws(operation, pattern);
}

async function runSelfTest(manifestPath) {
  const baselineDocument = await loadJson(manifestPath);
  const desired = validateManifest(baselineDocument);

  const missingField = structuredClone(baselineDocument);
  delete missingField.labels[0].description;
  expectFailure(() => validateManifest(missingField), /name, color, and description/);

  const invalidColor = structuredClone(baselineDocument);
  invalidColor.labels[0].color = "#c24e00";
  expectFailure(() => validateManifest(invalidColor), /uppercase hexadecimal/);

  const duplicateName = structuredClone(baselineDocument);
  duplicateName.labels[1].name = baselineDocument.labels[0].name.toUpperCase();
  expectFailure(() => validateManifest(duplicateName), /duplicates/);

  const longDescription = structuredClone(baselineDocument);
  longDescription.labels[0].description = "x".repeat(101);
  expectFailure(() => validateManifest(longDescription), /100-character/);

  const wrongCount = structuredClone(baselineDocument);
  wrongCount.labels.pop();
  expectFailure(() => validateManifest(wrongCount), /exactly 23/);

  const unsorted = structuredClone(baselineDocument);
  [unsorted.labels[0], unsorted.labels[1]] = [unsorted.labels[1], unsorted.labels[0]];
  expectFailure(() => validateManifest(unsorted), /sorted/);

  assert.equal(planChanges(desired, desired).length, 0, "idempotent plan must be empty");

  const missing = desired.slice(1);
  assert.deepEqual(planChanges(desired, missing).map(change => change.action), ["create"]);

  const changed = structuredClone(desired);
  changed[0].color = "000000";
  changed[0].description = "changed";
  const changedPlan = planChanges(desired, changed);
  assert.deepEqual(changedPlan.map(change => change.action), ["update"]);
  assert.deepEqual(changedPlan[0].fields, ["color", "description"]);

  const extra = [...structuredClone(desired), { name: "extra", color: "ABCDEF", description: "extra" }];
  assert.deepEqual(planChanges(desired, extra).map(change => change.action), ["delete"]);

  const combined = structuredClone(desired.slice(1));
  combined[0].description = "changed";
  combined.push({ name: "extra", color: "ABCDEF", description: "extra" });
  assert.deepEqual(planChanges(desired, combined).map(change => change.action).sort(), ["create", "delete", "update"]);

  process.stdout.write("SELF-TEST PASS: 12 validation and reconciliation cases\n");
}

async function main() {
  const options = parseArguments(process.argv.slice(2));

  if (options.selfTest) {
    await runSelfTest(options.manifest);
    return;
  }

  const document = await loadJson(options.manifest);
  const desired = validateManifest(document);

  if (options.mode === "validate") {
    await publishSummary(renderSummary({ mode: "validate", desiredCount: desired.length }));
    return;
  }

  let live;
  if (options.mode === "plan") {
    if (!options.live) throw new Error("--mode plan requires --live");
    const liveDocument = await loadJson(options.live);
    live = Array.isArray(liveDocument) ? liveDocument : liveDocument.labels;
    if (!Array.isArray(live)) throw new Error("live fixture must be an array or contain a labels array");
  } else if (options.mode === "reconcile") {
    if (!options.repository) throw new Error("--mode reconcile requires --repository or GITHUB_REPOSITORY");
    if (!process.env.GITHUB_TOKEN) throw new Error("--mode reconcile requires GITHUB_TOKEN");
    live = await listLiveLabels(options.repository, process.env.GITHUB_TOKEN);
  } else {
    throw new Error("--mode must be validate, plan, or reconcile");
  }

  const changes = planChanges(desired, live);

  if (options.mode === "plan" || options.dryRun) {
    await publishSummary(renderSummary({
      mode: options.mode,
      desiredCount: desired.length,
      changes,
      dryRun: options.dryRun
    }));
    return;
  }

  await applyChanges(options.repository, process.env.GITHUB_TOKEN, changes);
  const after = await listLiveLabels(options.repository, process.env.GITHUB_TOKEN);
  const remaining = planChanges(desired, after);
  const verified = remaining.length === 0;
  await publishSummary(renderSummary({ mode: "reconcile", desiredCount: desired.length, changes, verified }));

  if (!verified) {
    throw new Error(`Post-run verification found ${remaining.length} remaining difference(s)`);
  }
}

main().catch(async error => {
  const message = error instanceof Error ? error.message : String(error);
  try {
    await publishSummary(renderSummary({ mode: "failed", desiredCount: 0, error: message }));
  } finally {
    process.stderr.write(`${message}\n`);
    process.exitCode = 1;
  }
});
