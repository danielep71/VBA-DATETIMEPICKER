#!/usr/bin/env node

import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { appendFile, mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import process from "node:process";
import { promisify } from "node:util";
import { planChanges, resolveDesired, resolvePolicySelection, validateManifest } from "./labels-sync.mjs";

const execFileAsync = promisify(execFile);
const TOOL_NAME = "Repository label drift";
const API_VERSION = "2022-11-28";
const DEFAULT_MANIFEST = ".github/labels.json";
const DEFAULT_POLICY = ".github/repository-profile.json";
const LABELS_PER_PAGE = 100;
const MAX_LABEL_PAGES = 100;
const MAX_RATE_LIMIT_RETRIES = 2;
const MAX_RETRY_DELAY_MS = 5000;

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

function parseArguments(argv) {
  const parsed = {
    manifest: DEFAULT_MANIFEST,
    policy: DEFAULT_POLICY,
    live: null,
    repository: process.env.GITHUB_REPOSITORY || null,
    output: null,
    summary: null,
    selfTest: false
  };
  for (let index = 0; index < argv.length; index += 1) {
    const argument = argv[index];
    if (argument === "--self-test") {
      parsed.selfTest = true;
      continue;
    }
    if (["--manifest", "--policy", "--live", "--repository", "--output", "--summary"].includes(argument)) {
      const value = argv[index + 1];
      if (!value || value.startsWith("--")) throw new Error(`${argument} requires a value`);
      parsed[argument.slice(2)] = value;
      index += 1;
      continue;
    }
    throw new Error(`Unknown argument: ${argument}`);
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

function sleep(milliseconds) {
  return new Promise(resolve => setTimeout(resolve, milliseconds));
}

function isRateLimited(response) {
  return response.status === 429
    || (
      response.status === 403
      && (
        response.headers.get("x-ratelimit-remaining") === "0"
        || response.headers.has("retry-after")
      )
    );
}

function retryDelayMs(response, attempt) {
  const header = response.headers.get("retry-after");
  const retryAfter = header !== null && header.trim() !== "" ? Number(header) : NaN;
  if (Number.isFinite(retryAfter) && retryAfter >= 0) {
    return Math.min(retryAfter * 1000, MAX_RETRY_DELAY_MS);
  }
  return Math.min(1000 * (2 ** attempt), MAX_RETRY_DELAY_MS);
}

async function fetchLabelsPage(repository, token, page) {
  const suffix = `/labels?per_page=${LABELS_PER_PAGE}&page=${page}`;
  for (let attempt = 0; ; attempt += 1) {
    const response = await fetch(apiUrl(repository, suffix), {
      method: "GET",
      headers: {
        Accept: "application/vnd.github+json",
        Authorization: `Bearer ${token}`,
        "X-GitHub-Api-Version": API_VERSION,
        "User-Agent": "canonical-vba-label-drift"
      }
    });
    if (response.ok) {
      const batch = await response.json();
      if (!Array.isArray(batch)) throw new Error("GitHub labels response is not an array");
      return batch;
    }

    const detail = await response.text();
    if (!isRateLimited(response) || attempt >= MAX_RATE_LIMIT_RETRIES) {
      throw new Error(`GET ${suffix} failed with HTTP ${response.status}: ${detail.slice(0, 500)}`);
    }
    await sleep(retryDelayMs(response, attempt));
  }
}

async function listLiveLabels(repository, token) {
  const labels = [];
  for (let page = 1; page <= MAX_LABEL_PAGES; page += 1) {
    const batch = await fetchLabelsPage(repository, token, page);
    labels.push(...batch.map(normalizedLiveLabel));
    if (batch.length < LABELS_PER_PAGE) {
      return labels.sort((left, right) => compareNames(left.name, right.name));
    }
  }
  throw new Error(`GitHub label pagination exceeded ${MAX_LABEL_PAGES} pages`);
}

function localPlan(desiredLabels, liveLabels, { prune = true } = {}) {
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
      changes.push({ action: "update", name: target.name, currentName: current.name, fields, target });
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
  return changes.sort((left, right) => (
    order[left.action] - order[right.action] || compareNames(left.name, right.name)
  ));
}

async function verifyAgainstCanonicalPlan(manifestPath, policyPath, liveLabels, changes) {
  const { manifest, desiredLabels } = await prepareContract(manifestPath, policyPath);
  const canonical = planChanges(desiredLabels, liveLabels, { prune: manifest.prune });
  // Compare full structured changes; Markdown is only a human-readable report.
  assert.deepEqual(canonical, changes, "canonical/local change plans disagree");
}

function changeEvidence(change, observedByName) {
  const currentKey = String(change.currentName ?? change.name).toLowerCase();
  return {
    action: change.action,
    name: change.name,
    fields: change.fields ?? [],
    observed: observedByName.get(currentKey) ?? null,
    desired: change.target ?? null
  };
}

function buildReport({ repository, manifestPath, policyPath, manifest, selection, desiredLabels, liveLabels }) {
  const observed = liveLabels
    .map(normalizedLiveLabel)
    .sort((left, right) => compareNames(left.name, right.name));
  const desired = structuredClone(desiredLabels)
    .sort((left, right) => compareNames(left.name, right.name));
  const changes = localPlan(desired, observed, { prune: manifest.prune });
  const observedByName = new Map(observed.map(label => [label.name.toLowerCase(), label]));
  const differences = changes.map(change => changeEvidence(change, observedByName));
  const counts = Object.fromEntries(["create", "update", "delete"].map(action => [
    action,
    differences.filter(item => item.action === action).length
  ]));
  return {
    schema_version: 1,
    tool: TOOL_NAME,
    status: differences.length === 0 ? "pass" : "drift",
    repository,
    manifest: manifestPath,
    policy: policyPath,
    selection: { profile: selection.profile, domains: selection.domains },
    prune: manifest.prune,
    counts: {
      desired: desired.length,
      observed: observed.length,
      differences: differences.length,
      create: counts.create,
      update: counts.update,
      delete: counts.delete
    },
    desired,
    observed,
    differences,
    canonical_plan_verified: true
  };
}

function markdownEscape(value) {
  return String(value).replaceAll("|", "\\|").replaceAll("\n", " ");
}

function compactLabel(label) {
  if (label === null) return "—";
  return `\`${markdownEscape(label.name)}\` / \`${label.color}\` / ${markdownEscape(label.description)}`;
}

function markdownReport(report) {
  const lines = [
    "# Repository label drift",
    "",
    `- Status: **${String(report.status).toUpperCase()}**`,
    `- Repository: \`${report.repository ?? "fixture"}\``,
    `- Manifest: \`${report.manifest}\``,
    `- Policy: \`${report.policy}\``,
    `- Profile overlay: **${report.selection.profile ?? "none"}**`,
    `- Domain overlays: **${report.selection.domains.length > 0 ? report.selection.domains.join(", ") : "none"}**`,
    `- Desired / observed: **${report.counts.desired} / ${report.counts.observed}**`,
    `- Differences: **${report.counts.differences}**`,
    `- Create / update / delete: **${report.counts.create} / ${report.counts.update} / ${report.counts.delete}**`,
    "- Canonical reconciliation plan cross-check: **PASS**",
    ""
  ];
  if (report.differences.length === 0) {
    lines.push("No live label drift detected.");
  } else {
    lines.push(
      "| Action | Label | Fields | Observed | Desired |",
      "| --- | --- | --- | --- | --- |"
    );
    for (const difference of report.differences) {
      lines.push(
        `| ${difference.action} | \`${markdownEscape(difference.name)}\` | `
        + `${difference.fields.length > 0 ? difference.fields.join(", ") : "—"} | `
        + `${compactLabel(difference.observed)} | ${compactLabel(difference.desired)} |`
      );
    }
  }
  return `${lines.join("\n")}\n`;
}

async function writeEvidence(path, content) {
  if (!path) return;
  const slash = path.lastIndexOf("/");
  if (slash > 0) await mkdir(path.slice(0, slash), { recursive: true });
  await writeFile(path, content, "utf8");
}

async function publishSummary(summary) {
  process.stdout.write(summary);
  if (process.env.GITHUB_STEP_SUMMARY) {
    await appendFile(process.env.GITHUB_STEP_SUMMARY, summary, "utf8");
  }
}

async function prepareContract(manifestPath, policyPath) {
  const manifest = validateManifest(await loadJson(manifestPath));
  const selection = resolvePolicySelection(await loadJson(policyPath), manifest);
  const desiredLabels = resolveDesired(manifest, selection);
  return { manifest, selection, desiredLabels };
}

async function checkedReport(parameters) {
  const report = buildReport(parameters);
  const changes = localPlan(report.desired, report.observed, { prune: report.prune });
  await verifyAgainstCanonicalPlan(report.manifest, report.policy, report.observed, changes);
  return report;
}

async function testDelimitedLabels() {
  const directory = await mkdtemp(join(tmpdir(), "label-drift-fixtures-"));
  const manifestPath = join(directory, "manifest.json");
  const policyPath = join(directory, "policy.json");
  const livePath = join(directory, "live.json");
  const outputPath = join(directory, "report.json");
  const summaryPath = join(directory, "report.md");
  const runCli = async expectedCode => {
    await rm(outputPath, { force: true });
    await rm(summaryPath, { force: true });
    let code = 0;
    let stderr = "";
    try {
      await execFileAsync(process.execPath, [
        ".github/scripts/labels-drift.mjs", "--manifest", manifestPath,
        "--policy", policyPath, "--live", livePath,
        "--output", outputPath, "--summary", summaryPath
      ], { env: { ...process.env, GITHUB_TOKEN: "", GITHUB_STEP_SUMMARY: "" } });
    } catch (error) {
      code = error.code;
      stderr = error.stderr;
    }
    assert.equal(code, expectedCode, stderr);
    if (expectedCode === 2) {
      await assert.rejects(readFile(outputPath), { code: "ENOENT" });
      return stderr;
    }
    assert.equal(stderr, "");
    const report = await loadJson(outputPath);
    assert.equal(report.canonical_plan_verified, true);
    assert.equal(await readFile(summaryPath, "utf8"), markdownReport(report));
    return report;
  };
  try {
    await writeFile(policyPath, JSON.stringify({ mode: "template", profile: null, label_domains: [] }));
    for (const delimiter of ["|", "`", "`|\\"]) {
      const desired = ["create", "update"].map(prefix => ({
        name: `${prefix}${delimiter}label`, color: "123ABC", description: `Literal ${delimiter} description`
      }));
      const manifest = {
        schema_version: 1, prune: true, core: desired,
        overlays: { profile: { application: [], library: [], "ui-component": [] }, domain: {} }
      };
      await writeFile(manifestPath, JSON.stringify(manifest));
      await writeFile(livePath, JSON.stringify(desired));
      const clean = await runCli(0);
      assert.equal(clean.status, "pass");
      assert.deepEqual(clean.differences, []);
      assert.deepEqual(clean.desired, desired);

      const live = [
        { ...desired[1], name: desired[1].name.toUpperCase(), color: "000000", description: "changed" },
        { name: `obsolete${delimiter}label`, color: "ABCDEF", description: "extra" }
      ];
      await writeFile(livePath, JSON.stringify(live));
      const drift = await runCli(1);
      assert.equal(drift.status, "drift");
      assert.deepEqual(drift.differences.map(item => [item.action, item.name, item.fields]), [
        ["update", desired[1].name, ["name", "color", "description"]],
        ["create", desired[0].name, []],
        ["delete", live[1].name, []]
      ]);
      assert.deepEqual(drift.desired, desired);
      assert.deepEqual(drift, await runCli(1), "delimiter evidence must be deterministic");
      assert.deepEqual(await loadJson(livePath), live, "CLI must not mutate live fixtures");

      const changes = localPlan(desired, live);
      await assert.rejects(
        verifyAgainstCanonicalPlan(manifestPath, policyPath, live, changes.slice(1)),
        /canonical\/local change plans disagree/
      );
      const wrongTarget = structuredClone(changes);
      wrongTarget[0].target.description = "wrong target";
      await assert.rejects(
        verifyAgainstCanonicalPlan(manifestPath, policyPath, live, wrongTarget),
        /canonical\/local change plans disagree/
      );

      manifest.prune = false;
      await writeFile(manifestPath, JSON.stringify(manifest));
      const noPrune = await runCli(1);
      assert.deepEqual(noPrune.differences.map(item => item.action), ["update", "create"]);

      // Delimiters remain legal, but invalid manifests must still fail closed.
      manifest.core[0].color = "invalid";
      await writeFile(manifestPath, JSON.stringify(manifest));
      assert.match(await runCli(2), /uppercase hexadecimal/);
    }

    // Even reconcile arguments and a token must not run the CLI on import.
    const { stdout, stderr } = await execFileAsync(process.execPath, [
      "--input-type=module", "--eval",
      'process.argv = [process.execPath, "import-fixture.mjs", "--mode", "reconcile"]; '
      + 'globalThis.fetch = () => { throw new Error("unexpected API call"); }; '
      + 'await import("./.github/scripts/labels-sync.mjs");'
    ], { env: { ...process.env, GITHUB_TOKEN: "fixture", GITHUB_STEP_SUMMARY: summaryPath } });
    assert.equal(stdout, "", "import must not publish a CLI summary");
    assert.equal(stderr, "");
    await assert.rejects(readFile(summaryPath), { code: "ENOENT" });
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
}

async function runSelfTest(manifestPath, policyPath) {
  for (const [header, attempt, expected] of [
    [null, 0, 1000], [null, 2, 4000], ["", 1, 2000],
    ["invalid", 1, 2000], ["-1", 1, 2000], ["0", 2, 0],
    ["2", 0, 2000], ["999999", 0, MAX_RETRY_DELAY_MS],
    [null, 20, MAX_RETRY_DELAY_MS]
  ]) {
    const headers = new Headers(header === null ? {} : { "retry-after": header });
    assert.equal(retryDelayMs({ headers }, attempt), expected, `Retry-After: ${header}`);
  }

  const { manifest, selection, desiredLabels } = await prepareContract(manifestPath, policyPath);

  const baseline = structuredClone(desiredLabels);
  const baselineBefore = JSON.stringify(baseline);
  const cleanFirst = await checkedReport({
    repository: "fixture/clean",
    manifestPath,
    policyPath,
    manifest,
    selection,
    desiredLabels,
    liveLabels: baseline
  });
  const cleanSecond = await checkedReport({
    repository: "fixture/clean",
    manifestPath,
    policyPath,
    manifest,
    selection,
    desiredLabels,
    liveLabels: baseline
  });
  assert.equal(cleanFirst.status, "pass");
  assert.equal(cleanFirst.counts.differences, 0);
  assert.equal(JSON.stringify(cleanFirst), JSON.stringify(cleanSecond));
  assert.equal(markdownReport(cleanFirst), markdownReport(cleanSecond));
  assert.equal(JSON.stringify(baseline), baselineBefore, "clean check must not mutate live input");

  const drifted = structuredClone(desiredLabels);
  const missing = drifted.shift();
  assert.ok(missing, "fixture requires at least one desired label");
  drifted[0].color = drifted[0].color === "000000" ? "FFFFFF" : "000000";
  drifted[0].description = "simulated out-of-band change";
  drifted.push({ name: "out-of-band", color: "ABCDEF", description: "simulated extra label" });
  const driftedBefore = JSON.stringify(drifted);
  const driftFirst = await checkedReport({
    repository: "fixture/drift",
    manifestPath,
    policyPath,
    manifest,
    selection,
    desiredLabels,
    liveLabels: drifted
  });
  const driftSecond = await checkedReport({
    repository: "fixture/drift",
    manifestPath,
    policyPath,
    manifest,
    selection,
    desiredLabels,
    liveLabels: drifted
  });
  assert.equal(driftFirst.status, "drift");
  assert.deepEqual(driftFirst.differences.map(item => item.action), ["update", "create", "delete"]);
  assert.equal(JSON.stringify(driftFirst), JSON.stringify(driftSecond));
  assert.equal(markdownReport(driftFirst), markdownReport(driftSecond));
  assert.equal(JSON.stringify(drifted), driftedBefore, "drift check must not mutate live input");
  assert.match(markdownReport(driftFirst), /simulated out-of-band change/);
  assert.match(markdownReport(driftFirst), new RegExp(missing.name));

  await testDelimitedLabels();

  process.stdout.write(
    "SELF-TEST PASS: delimiter CLI/negative/import fixtures, retry delays and deterministic no-drift and create/update/delete drift fixtures are read-only and match the canonical reconciler.\n"
  );
}

async function main() {
  const options = parseArguments(process.argv.slice(2));
  if (options.selfTest) {
    await runSelfTest(options.manifest, options.policy);
    return;
  }

  const { manifest, selection, desiredLabels } = await prepareContract(options.manifest, options.policy);

  let liveLabels;
  if (options.live) {
    const liveDocument = await loadJson(options.live);
    liveLabels = Array.isArray(liveDocument) ? liveDocument : liveDocument.labels;
    if (!Array.isArray(liveLabels)) {
      throw new Error("live fixture must be an array or contain a labels array");
    }
  } else {
    if (!options.repository) throw new Error("live check requires --repository or GITHUB_REPOSITORY");
    if (!process.env.GITHUB_TOKEN) throw new Error("live check requires GITHUB_TOKEN");
    liveLabels = await listLiveLabels(options.repository, process.env.GITHUB_TOKEN);
  }

  const report = await checkedReport({
    repository: options.repository,
    manifestPath: options.manifest,
    policyPath: options.policy,
    manifest,
    selection,
    desiredLabels,
    liveLabels
  });
  const json = `${JSON.stringify(report, null, 2)}\n`;
  const markdown = markdownReport(report);
  await writeEvidence(options.output, json);
  await writeEvidence(options.summary, markdown);
  await publishSummary(markdown);
  process.exitCode = report.status === "pass" ? 0 : 1;
}

main().catch(error => {
  const message = error instanceof Error ? error.message : String(error);
  process.stderr.write(`ERROR: ${message}\n`);
  process.exitCode = 2;
});
