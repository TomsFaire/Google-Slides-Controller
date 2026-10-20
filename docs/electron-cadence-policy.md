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
