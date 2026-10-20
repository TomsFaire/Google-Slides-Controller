#!/usr/bin/env node
/**
 * Report this app's position in Electron's support window.
 *
 * Prints a JSON report to stdout. Exits 1 when the state is critical, so a CI
 * run turns red. Network failures exit 2 and are NOT treated as a drift signal —
 * a flaky registry must not open a bogus issue.
 *
 * Usage: node scripts/check-electron-support.js
 */
const path = require('path');
const {
  parseMajor,
  assessSupport,
  exitCodeForFailure,
  networkError,
  classifyHttpFailure,
  EXIT
} = require(path.join(__dirname, '..', 'src', 'electron-support'));

const DIST_TAGS_URL = 'https://registry.npmjs.org/-/package/electron/dist-tags';
const BREAKING_CHANGES_URL = (v) => `https://raw.githubusercontent.com/electron/electron/v${v}/docs/breaking-changes.md`;
const REGISTRY_VERSION_URL = (v) => `https://registry.npmjs.org/electron/${v}`;

// Every failure reaching the network is tagged, so exitCodeForFailure can tell a
// transient blip apart from a bad config or a bug. Untagged throws are INTERNAL.
async function getJson(url) {
  let res;
  try {
    res = await fetch(url, { headers: { accept: 'application/json' } });
  } catch (e) {
    throw networkError(`GET ${url} failed: ${e.message}`);
  }
  if (!res.ok) throw classifyHttpFailure(url, res.status);
  try {
    return await res.json();
  } catch (e) {
    throw networkError(`GET ${url} returned unparseable JSON: ${e.message}`);
  }
}

async function getText(url) {
  let res;
  try {
    res = await fetch(url);
  } catch (e) {
    throw networkError(`GET ${url} failed: ${e.message}`);
  }
  if (!res.ok) throw classifyHttpFailure(url, res.status);
  return res.text();
}

/**
 * Pull the macOS-floor-raising entries out of the target's breaking-changes doc.
 * These are the "Removed: macOS N support" headings; the floor is N+1.
 */
function extractMacosFloorNotes(breakingChanges) {
  return (breakingChanges.match(/^### Removed: macOS .*$/gm) || []).map((l) => l.replace(/^### /, ''));
}

async function main() {
  const pkg = require(path.join(__dirname, '..', 'package.json'));
  const currentRange = (pkg.devDependencies && pkg.devDependencies.electron) || '';
  const currentMajor = parseMajor(currentRange);

  const distTags = await getJson(DIST_TAGS_URL);
  const latestVersion = distTags.latest;
  const latestMajor = parseMajor(latestVersion);

  const assessment = assessSupport({ currentMajor, latestMajor });

  let nodeRequirement = 'unknown';
  let macosNotes = [];
  if (assessment.state !== 'supported') {
    // Only spend these calls when we are about to write an issue.
    try {
      const meta = await getJson(REGISTRY_VERSION_URL(latestVersion));
      nodeRequirement = (meta.engines && meta.engines.node) || 'unspecified';
    } catch (e) {
      nodeRequirement = `lookup failed: ${e.message}`;
    }
    try {
      macosNotes = extractMacosFloorNotes(await getText(BREAKING_CHANGES_URL(latestVersion)));
    } catch (e) {
      macosNotes = [`macOS floor lookup failed: ${e.message}`];
    }
  }

  const title = assessment.state === 'critical'
    ? `Electron ${currentMajor} is ${assessment.majorsPastEol} majors past end-of-life (latest: ${latestMajor})`
    : `Electron ${currentMajor} has left the support window (latest: ${latestMajor})`;

  const body = [
    `The app is on **Electron ${currentMajor}**; the latest stable is **${latestVersion}**.`,
    '',
    `- Majors behind: **${assessment.majorsBehind}**`,
    `- Majors past end-of-life: **${assessment.majorsPastEol}**`,
    `- Oldest still-supported major: **${assessment.oldestSupportedMajor}**`,
    `- Target's Node requirement: \`${nodeRequirement}\``,
    '',
    '**macOS floor changes to check:**',
    ...(macosNotes.length ? macosNotes.map((n) => `- ${n}`) : ['- none found in the target\'s breaking-changes doc']),
    '',
    'The bundled Chromium receives no security patches once a major leaves the',
    'support window, and this app loads `docs.google.com` in a session holding live',
    'Google credentials.',
    '',
    // GitHub does not resolve relative links in issue bodies, so these are plain paths.
    'Upgrade procedure: `docs/electron-upgrade-checklist.md`',
    'Cadence policy: `docs/electron-cadence-policy.md`',
    '',
    '<!-- maintained by .github/workflows/electron-support-check.yml -->'
  ].join('\n');

  const report = { ...assessment, currentMajor, latestMajor, latestVersion, nodeRequirement, macosNotes, title, body };
  process.stdout.write(JSON.stringify(report, null, 2) + '\n');
  process.exitCode = assessment.state === 'critical' ? EXIT.DRIFT_CRITICAL : EXIT.OK;
}

main().catch((err) => {
  const code = exitCodeForFailure(err);
  const label = code === EXIT.NETWORK ? 'network failure (retry later)' : 'internal error (needs a human)';
  process.stderr.write(`electron support check failed - ${label}: ${err.message}\n`);
  // 2 = could not reach the registry, safe to retry. 3 = bad config or a bug, which
  // must NOT be silently retried forever, or the alarm never fires.
  process.exitCode = code;
});
