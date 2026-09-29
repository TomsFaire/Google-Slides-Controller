# Electron Stay-Current Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a process that makes falling off Electron's support window impossible to miss, without auto-merging anything.

**Architecture:** A pure-logic support-assessment module (unit tested, no Electron dependency), a script that feeds it live registry data, a weekly workflow that maintains exactly one escalating GitHub issue, Dependabot for routine bumps, and a written cadence policy.

**Tech Stack:** `node:test`, Node 22 `fetch`, GitHub Actions, `gh` CLI, Dependabot.

**Spec:** [docs/plans/electron-upgrade-and-stay-current.md](electron-upgrade-and-stay-current.md) §7 — read it alongside this plan.

**Sibling plan:** [electron-upgrade-implementation.md](electron-upgrade-implementation.md) delivers §7.1 (`.nvmrc` + `engines.node`) as its Task 1, because PR 1's CI depends on it. This plan assumes that already landed.

## Global Constraints

- **Nothing auto-merges.** Every dependency update opens a PR and waits for human review. This is a deliberate choice; parts of this plan exist specifically to stop the resulting queue from rotting.
- **Electron support policy:** the latest **3** majors are supported. A major ships roughly every 8 weeks, giving ~24 weeks of runway.
- **Node floor:** `>=22.20 <23` — the only range satisfying both `electron@44.4.5` and the companion module's `^22.20`.
- **Two Yarn ecosystems:** root is Yarn 1 (`yarn.lock`, `--frozen-lockfile`); `companion-module-gslide-opener` is Yarn 4 (`--immutable`). Dependabot must cover both.
- **`electron` and `electron-builder` are a coupled pair** and must always be grouped into one PR.
- **Silence when healthy.** A workflow that cries wolf gets muted, which defeats the whole design.

---

## File Structure

| File | Change | Responsibility |
|---|---|---|
| `src/electron-support.js` | **create** | Pure support-window arithmetic. No I/O, no Electron import — so it is unit testable |
| `tests/electron-support.test.js` | **create** | Drives the module's behaviour, including the boundary cases |
| `scripts/check-electron-support.js` | **create** | Fetches live registry/release data, calls the module, emits the issue body |
| `.github/workflows/electron-support-check.yml` | **create** | Weekly schedule; maintains exactly one labelled issue; fails the run when critical |
| `.github/dependabot.yml` | **create** | Weekly PRs for both ecosystems, with the Electron pair grouped |
| `docs/electron-cadence-policy.md` | **create** | When to upgrade, and the blackout window |

Splitting logic (`src/`) from I/O (`scripts/`) is what makes the interesting part testable — and it follows the repo's existing pattern, where `src/perfectcue-*.js` and `src/notes-window-adoption.js` hold logic that `tests/` requires directly. `main.js` cannot be required from a test, so nothing in this plan puts logic there.

---

## Task 1: The support-window module

**Files:**
- Create: `src/electron-support.js`
- Test: `tests/electron-support.test.js`

**Interfaces:**
- Produces:
  - `parseMajor(versionRange: string) => number` — accepts `"^44.4.5"`, `"44.4.5"`, `"~44.0.0"`, `">=44"`; throws on unparseable input
  - `assessSupport({ currentMajor: number, latestMajor: number, supportedMajors?: number }) => { state: 'supported'|'unsupported'|'critical', majorsBehind: number, majorsPastEol: number, oldestSupportedMajor: number }`
  - Task 2 consumes both of these.

- [ ] **Step 1: Write the failing test**

Create `tests/electron-support.test.js`:

```javascript
const { test } = require('node:test');
const assert = require('node:assert/strict');

const { parseMajor, assessSupport } = require('../src/electron-support');

// ── parseMajor ───────────────────────────────────────────────────────────────

test('parseMajor reads the major from a caret range', () => {
  assert.equal(parseMajor('^44.4.5'), 44);
});

test('parseMajor reads the major from an exact version', () => {
  assert.equal(parseMajor('44.4.5'), 44);
});

test('parseMajor reads the major from tilde and gte ranges', () => {
  assert.equal(parseMajor('~44.0.0'), 44);
  assert.equal(parseMajor('>=44'), 44);
});

test('parseMajor throws on unparseable input', () => {
  assert.throws(() => parseMajor('not-a-version'), /unparseable/i);
  assert.throws(() => parseMajor(''), /unparseable/i);
});

// ── assessSupport: the healthy case must be silent ───────────────────────────

test('the latest major is supported', () => {
  const r = assessSupport({ currentMajor: 44, latestMajor: 44 });
  assert.equal(r.state, 'supported');
  assert.equal(r.majorsBehind, 0);
  assert.equal(r.majorsPastEol, 0);
});

test('the oldest of the latest three majors is still supported', () => {
  const r = assessSupport({ currentMajor: 42, latestMajor: 44 });
  assert.equal(r.state, 'supported');
  assert.equal(r.majorsBehind, 2);
  assert.equal(r.oldestSupportedMajor, 42);
});

// ── assessSupport: the escalation boundary ───────────────────────────────────

test('one major past end-of-life is unsupported, not yet critical', () => {
  const r = assessSupport({ currentMajor: 41, latestMajor: 44 });
  assert.equal(r.state, 'unsupported');
  assert.equal(r.majorsPastEol, 1);
  assert.equal(r.majorsBehind, 3);
});

test('two majors past end-of-life escalates to critical', () => {
  const r = assessSupport({ currentMajor: 40, latestMajor: 44 });
  assert.equal(r.state, 'critical');
  assert.equal(r.majorsPastEol, 2);
});

test('the situation this whole process exists to prevent', () => {
  // Electron 33 against a latest of 44 — where the app actually was.
  const r = assessSupport({ currentMajor: 33, latestMajor: 44 });
  assert.equal(r.state, 'critical');
  assert.equal(r.majorsBehind, 11);
  assert.equal(r.majorsPastEol, 9);
});

// ── assessSupport: edges ─────────────────────────────────────────────────────

test('running ahead of stable (alpha/beta) counts as supported', () => {
  const r = assessSupport({ currentMajor: 45, latestMajor: 44 });
  assert.equal(r.state, 'supported');
  assert.equal(r.majorsBehind, -1);
  assert.equal(r.majorsPastEol, 0);
});

test('the support window size is configurable', () => {
  const r = assessSupport({ currentMajor: 43, latestMajor: 44, supportedMajors: 1 });
  assert.equal(r.state, 'unsupported');
  assert.equal(r.oldestSupportedMajor, 44);
});

test('assessSupport rejects non-numeric input rather than guessing', () => {
  assert.throws(() => assessSupport({ currentMajor: NaN, latestMajor: 44 }), /must be a number/i);
  assert.throws(() => assessSupport({ currentMajor: 44, latestMajor: undefined }), /must be a number/i);
});
```

- [ ] **Step 2: Run the test to verify it fails**

```bash
node --test tests/electron-support.test.js
```

Expected: FAIL — `Cannot find module '../src/electron-support'`

- [ ] **Step 3: Write the minimal implementation**

Create `src/electron-support.js`:

```javascript
/**
 * Electron support-window arithmetic.
 *
 * Electron supports the latest N majors (N = 3 as of this writing). Deliberately
 * pure: no I/O and no `require('electron')`, so it can be unit tested. Live data
 * fetching lives in scripts/check-electron-support.js.
 */

const DEFAULT_SUPPORTED_MAJORS = 3;

/** Extract the major version from a version or semver range. */
function parseMajor(versionRange) {
  const match = String(versionRange || '').match(/(\d+)\s*\./) || String(versionRange || '').match(/(\d+)\s*$/);
  if (!match) {
    throw new Error(`unparseable Electron version: ${JSON.stringify(versionRange)}`);
  }
  return Number(match[1]);
}

/**
 * Assess a major against the support window.
 *
 * States:
 *   supported   - within the latest N majors; callers should stay silent
 *   unsupported - exactly 1 major past end-of-life; open one issue
 *   critical    - 2 or more majors past end-of-life; escalate and fail the run
 */
function assessSupport({ currentMajor, latestMajor, supportedMajors = DEFAULT_SUPPORTED_MAJORS }) {
  for (const [name, value] of [['currentMajor', currentMajor], ['latestMajor', latestMajor]]) {
    if (typeof value !== 'number' || !Number.isFinite(value)) {
      throw new Error(`${name} must be a number, got ${JSON.stringify(value)}`);
    }
  }

  const oldestSupportedMajor = latestMajor - (supportedMajors - 1);
  const majorsBehind = latestMajor - currentMajor;
  const majorsPastEol = Math.max(0, oldestSupportedMajor - currentMajor);

  let state = 'supported';
  if (majorsPastEol === 1) state = 'unsupported';
  else if (majorsPastEol >= 2) state = 'critical';

  return { state, majorsBehind, majorsPastEol, oldestSupportedMajor };
}

module.exports = { parseMajor, assessSupport, DEFAULT_SUPPORTED_MAJORS };
```

- [ ] **Step 4: Run the test to verify it passes**

```bash
node --test tests/electron-support.test.js
```

Expected: PASS, 12 tests.

- [ ] **Step 5: Confirm the whole suite still passes**

```bash
yarn test
```

Expected: all pass — the existing 103 plus the 12 new ones.

- [ ] **Step 6: Commit**

```bash
git add src/electron-support.js tests/electron-support.test.js
git commit -m "feat: add Electron support-window assessment module

Pure arithmetic over Electron's latest-3-majors support policy, with the
escalation boundary (1 major past EOL vs 2+) under test. No I/O and no
electron import, so it is unit testable."
```

---

## Task 2: The data-fetching script

**Files:**
- Create: `scripts/check-electron-support.js`

**Interfaces:**
- Consumes: `parseMajor`, `assessSupport` from `src/electron-support.js` (Task 1)
- Produces: a CLI that prints a JSON object `{ state, currentMajor, latestMajor, latestVersion, majorsBehind, majorsPastEol, nodeRequirement, macosNotes, title, body }` to stdout, and exits `1` when `state === 'critical'`. Task 3's workflow consumes it.

- [ ] **Step 1: Write the script**

Create `scripts/check-electron-support.js`:

```javascript
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
const { parseMajor, assessSupport } = require(path.join(__dirname, '..', 'src', 'electron-support'));

const DIST_TAGS_URL = 'https://registry.npmjs.org/-/package/electron/dist-tags';
const BREAKING_CHANGES_URL = (v) => `https://raw.githubusercontent.com/electron/electron/v${v}/docs/breaking-changes.md`;
const REGISTRY_VERSION_URL = (v) => `https://registry.npmjs.org/electron/${v}`;

async function getJson(url) {
  const res = await fetch(url, { headers: { accept: 'application/json' } });
  if (!res.ok) throw new Error(`GET ${url} -> ${res.status}`);
  return res.json();
}

async function getText(url) {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`GET ${url} -> ${res.status}`);
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
  process.exitCode = assessment.state === 'critical' ? 1 : 0;
}

main().catch((err) => {
  process.stderr.write(`electron support check failed: ${err.message}\n`);
  // Exit 2, distinct from the critical-drift exit 1: a network blip is not drift.
  process.exitCode = 2;
});
```

- [ ] **Step 2: Run it against the current state**

```bash
node scripts/check-electron-support.js; echo "exit=$?"
```

Expected **before** the upgrade plan lands (Electron 33): `"state": "critical"` and `exit=1`, with `majorsBehind: 11`.

Expected **after** the upgrade plan lands (Electron 44): `"state": "supported"` and `exit=0`.

Either outcome proves the script works; note which you saw.

- [ ] **Step 3: Verify a network failure is not mistaken for drift**

```bash
node -e "
global.fetch = () => Promise.reject(new Error('simulated offline'));
process.argv[1] = require('path').join(process.cwd(),'scripts/check-electron-support.js');
require('./scripts/check-electron-support.js');
" ; echo "exit=$?"
```

Expected: `exit=2` and a message on stderr. Exit 2 must never open an issue — Task 3 depends on this distinction.

- [ ] **Step 4: Commit**

```bash
git add scripts/check-electron-support.js
git commit -m "feat: add Electron support-drift check script

Feeds live npm dist-tags into the support-window module and emits a JSON
report with the actionable delta: majors behind, the target's
engines.node, and its macOS-floor changes. Exits 1 on critical drift and
2 on network failure, so a flaky registry cannot masquerade as drift."
```

---

## Task 3: The weekly drift alarm workflow

**Files:**
- Create: `.github/workflows/electron-support-check.yml`

**Interfaces:**
- Consumes: `scripts/check-electron-support.js` (Task 2) — its JSON stdout and its exit code

The anti-rot property: the job finds an existing open issue **by label** and edits it in place, so there is never more than one. A pile of stale notifications is what people learn to ignore.

- [ ] **Step 1: Write the workflow**

```yaml
name: Electron support check

on:
  schedule:
    # Mondays, 09:00 UTC
    - cron: '0 9 * * 1'
  workflow_dispatch:

permissions:
  contents: read
  issues: write

jobs:
  check:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4

      - uses: actions/setup-node@v4
        with:
          node-version-file: '.nvmrc'

      - name: Assess the support window
        id: assess
        run: |
          set +e
          node scripts/check-electron-support.js > report.json
          code=$?
          set -e
          if [ "$code" = "2" ]; then
            echo "Support check could not reach the network; not treating as drift."
            exit 0
          fi
          echo "state=$(node -p "require('./report.json').state")" >> "$GITHUB_OUTPUT"
          echo "title=$(node -p "require('./report.json').title")" >> "$GITHUB_OUTPUT"
          node -p "require('./report.json').body" > issue-body.md
          echo "exit_code=$code" >> "$GITHUB_OUTPUT"

      - name: Healthy — nothing to report
        if: steps.assess.outputs.state == 'supported'
        run: echo "Electron is within the support window. Staying silent."

      - name: Open or update the single drift issue
        if: steps.assess.outputs.state == 'unsupported' || steps.assess.outputs.state == 'critical'
        env:
          GH_TOKEN: ${{ secrets.GITHUB_TOKEN }}
          STATE: ${{ steps.assess.outputs.state }}
          TITLE: ${{ steps.assess.outputs.title }}
        run: |
          existing=$(gh issue list --label electron-drift --state open --limit 1 --json number --jq '.[0].number // empty')
          labels="electron-drift"
          if [ "$STATE" = "critical" ]; then labels="electron-drift,critical"; fi

          if [ -n "$existing" ]; then
            echo "Updating existing issue #$existing in place."
            gh issue edit "$existing" --title "$TITLE" --body-file issue-body.md --add-label "$labels"
          else
            echo "No open drift issue; creating one."
            gh issue create --title "$TITLE" --body-file issue-body.md --label "$labels"
          fi

      - name: Fail the run on critical drift
        if: steps.assess.outputs.state == 'critical'
        run: |
          echo "::error::Electron is 2+ majors past end-of-life. See the electron-drift issue."
          exit 1
```

- [ ] **Step 2: Create the labels the workflow uses**

```bash
gh label create electron-drift --description "Electron has left the support window" --color B60205 || true
gh label create critical --description "Needs attention now" --color B60205 || true
```

`gh label create` fails harmlessly if the label already exists; `|| true` keeps the step green.

- [ ] **Step 3: Verify the workflow is valid YAML**

```bash
python3 -c "import yaml,sys; yaml.safe_load(open('.github/workflows/electron-support-check.yml')); print('valid YAML')"
```

Expected: `valid YAML`

- [ ] **Step 4: Dry-run the logic locally**

```bash
node scripts/check-electron-support.js > report.json; echo "exit=$?"
node -p "require('./report.json').state"
node -p "require('./report.json').title"
node -p "require('./report.json').body" | head -20
rm -f report.json
```

Expected: the state, a sensible title, and a body containing the majors-behind count and the Node requirement.

- [ ] **Step 5: Trigger it for real once merged**

```bash
gh workflow run "Electron support check"
gh run watch
```

Expected: if the upgrade plan has landed, the run is green and silent. If it has not, the run fails with exit 1 and exactly one `electron-drift` issue exists. Run it **twice** and confirm the second run edits that same issue rather than opening a second one — this is the property the whole design rests on.

- [ ] **Step 6: Commit**

```bash
git add .github/workflows/electron-support-check.yml
git commit -m "ci: weekly Electron support-drift alarm

Silent while within the latest 3 majors. One major past EOL opens a
single labelled issue; 2+ escalates the label and fails the run. The
issue is located by label and edited in place, so it cannot decay into a
pile of stale notifications. Network failure (exit 2) is not drift."
```

---

## Task 4: Dependabot for both ecosystems

**Files:**
- Create: `.github/dependabot.yml`

- [ ] **Step 1: Write the config**

```yaml
version: 2
updates:
  # Root app — Yarn 1
  - package-ecosystem: npm
    directory: "/"
    schedule:
      interval: weekly
      day: monday
    open-pull-requests-limit: 5
    labels:
      - dependencies
    groups:
      # electron and electron-builder are a coupled pair: a major Electron bump
      # generally needs a matching builder. Never let them move separately.
      electron-toolchain:
        patterns:
          - "electron"
          - "electron-builder"
    commit-message:
      prefix: "chore"
      include: scope

  # Companion module — Yarn 4, separate lockfile
  - package-ecosystem: npm
    directory: "/companion-module-gslide-opener"
    schedule:
      interval: weekly
      day: monday
    open-pull-requests-limit: 3
    labels:
      - dependencies
      - companion-module
    commit-message:
      prefix: "chore"
      include: scope

  # Keep the CI actions themselves current
  - package-ecosystem: github-actions
    directory: "/"
    schedule:
      interval: monthly
    labels:
      - dependencies
      - ci
```

Nothing here enables auto-merge — every PR waits for review, as chosen.

- [ ] **Step 2: Validate the YAML**

```bash
python3 -c "import yaml; d=yaml.safe_load(open('.github/dependabot.yml')); print('valid YAML,', len(d['updates']), 'ecosystems')"
```

Expected: `valid YAML, 3 ecosystems`

- [ ] **Step 3: Create the labels**

```bash
gh label create dependencies --description "Dependency updates" --color 0366D6 || true
gh label create companion-module --description "Companion module" --color 5319E7 || true
gh label create ci --description "CI and build" --color 1D76DB || true
```

- [ ] **Step 4: Confirm GitHub accepted the config after merge**

```bash
gh api "repos/{owner}/{repo}/dependabot/alerts" --silent 2>/dev/null || true
gh browse --no-browser --settings 2>/dev/null || echo "Check Insights -> Dependency graph -> Dependabot for parse errors"
```

Expected: no parse error reported on the repo's Dependabot page. A malformed file fails silently in the UI rather than in CI, so this check matters.

- [ ] **Step 5: Commit**

```bash
git add .github/dependabot.yml
git commit -m "ci: add Dependabot for both Yarn ecosystems

Weekly PRs for the root (Yarn 1) and companion module (Yarn 4), plus
monthly GitHub Actions updates. electron and electron-builder are
grouped so the coupled pair always moves together. Nothing auto-merges."
```

---

## Task 5: The cadence policy

**Files:**
- Create: `docs/electron-cadence-policy.md`

**Blocked on one answer:** the blackout window. Write everything else, leave that section's dates as the single explicit question for the owner — do not invent an event season.

- [ ] **Step 1: Write the policy**

```markdown
# Electron Cadence Policy

## The arithmetic

Electron ships a major roughly every **8 weeks** and supports the **latest 3**.
Each major therefore grants about **24 weeks** (~5.5 months) of runway.

Waiting until end-of-life to act guarantees shipping an unpatched Chromium. The
rule is to move *while still supported*.

## The rule

**Upgrade on reaching N-2** — that is, when two newer majors exist. In practice
this means roughly **twice a year**, hopping 2-3 majors at a time.

| Position | Meaning | Action |
|---|---|---|
| N (latest) | Current | Nothing |
| N-1 | Supported | Nothing |
| N-2 | Supported, oldest in window | **Plan the upgrade now** |
| N-3 | Out of support | Overdue — the drift alarm has opened an issue |
| N-4 or older | Unpatched Chromium | Critical — the drift alarm is failing CI |

The weekly `Electron support check` workflow enforces the last two rows. It stays
silent for the first three, so a notification always means something.

## Procedure

Follow [electron-upgrade-checklist.md](electron-upgrade-checklist.md). It carries
both the hardware acceptance gates and the method for auditing breaking changes,
so an upgrade is a checklist run rather than a fresh investigation.

Two constraints to confirm before committing to a target version:

1. **macOS floor.** Each major may raise it. Verify every venue machine clears the
   target's floor first; it can cap the target regardless of anything else.
2. **`engines.node`.** The target's Node requirement must be reconcilable with the
   companion module's (`^22.20` at the time of writing). These have conflicted
   before: Electron 44 wants `>=22.12.0` while the companion module forbids
   Node 23+, leaving Node 22.20 as the only workable choice.

## Blackout window

Chromium jumps must not land close to a live event.

> **TO BE FILLED IN BY THE OWNER.** No upgrade is merged or released between
> `<START>` and `<END>` (the event season). Schedule hardware acceptance passes
> and releases outside that window.

Until this is filled in, treat any week containing a live show as a blackout and
schedule around it manually.

## Dependency review

Nothing auto-merges. Dependabot opens PRs weekly; they wait for review. Because
the queue is entirely manual, the drift alarm exists as the backstop: it is the
one signal that escalates on its own and eventually fails CI.
```

- [ ] **Step 2: Verify the internal link resolves**

```bash
cd docs && ([ -e electron-upgrade-checklist.md ] && echo "OK" || echo "MISS — the upgrade plan's Task 7 must land first"); cd ..
```

Expected: `OK`. If it reports MISS, the sibling upgrade plan's Task 7 has not landed; that is fine, but note the dangling link.

- [ ] **Step 3: Commit**

```bash
git add docs/electron-cadence-policy.md
git commit -m "docs: add Electron cadence policy

Upgrade on reaching N-2 rather than after end-of-life: a major grants
~24 weeks of runway, so roughly twice a year. Records the macOS-floor
and engines.node checks that can cap a target. The blackout window is
left for the owner to fill in."
```

---

## Gate — plan verification

- [ ] `yarn test` green, including the 12 new support-module tests
- [ ] `node scripts/check-electron-support.js` reports the true current state
- [ ] A forced-offline run exits **2**, not 1 — a network blip is not drift
- [ ] `gh workflow run "Electron support check"` behaves correctly for the current state
- [ ] Running it **twice** edits one issue rather than opening two
- [ ] Dependabot's page shows no parse error
- [ ] `docs/electron-cadence-policy.md` blackout window is filled in, or its being open is explicitly accepted

---

## Open parameter

**Blackout window.** Task 5 leaves `<START>`/`<END>` for the owner. Everything else
in this plan is complete and implementable without it.
