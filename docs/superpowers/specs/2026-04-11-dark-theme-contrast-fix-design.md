# Dark Theme Contrast Fix

**Date:** 2026-04-11  
**Status:** Approved

## Problem

`theme-dark` and `theme-thumb` both use dark backgrounds but inherit many base CSS text colors designed for light backgrounds. The result is dark-on-dark text that fails WCAG AA contrast (4.5:1 minimum). Affected: labels, section headings, control buttons, slide preview labels, build number, hint text (`<small>`), disabled states, inputs/selects, and dynamically-injected preset labels.

`theme-light`, `theme-max`, and `theme-touch` are unaffected — all pass on their light backgrounds.

## Approach

**Option A (chosen):** Add missing `body.theme-dark` and `body.theme-thumb` CSS overrides for every failing element. Use `!important` only where inline `style=""` attributes cannot be overridden otherwise. No HTML restructuring, no new classes.

## Color Palette

| Role | Value | Contrast on theme-dark (~#252527) | Contrast on theme-thumb (~#222834) |
|------|-------|-----------------------------------|------------------------------------|
| Primary text | `rgba(255,255,255,0.87)` | ~12:1 ✓ | ~11:1 ✓ |
| Secondary text | `rgba(255,255,255,0.65)` | ~8:1 ✓ | ~7.5:1 ✓ |
| Muted text | `rgba(255,255,255,0.45)` | ~4.8:1 ✓ | ~4.6:1 ✓ |

## Changes — `main.js` theme CSS block

All changes are additive overrides appended to the existing `body.theme-dark` and `body.theme-thumb` rule groups in the `<style id="theme-overrides">` block.

### Selectors to add to `body.theme-dark`

| Selector | Property | Value | Note |
|----------|----------|-------|------|
| `body.theme-dark label` | `color` | `rgba(255,255,255,0.87)` | |
| `body.theme-dark .btn-control` | `color` | `rgba(255,255,255,0.87)` | |
| `body.theme-dark .btn-control` | `background` | `rgba(255,255,255,0.08)` | |
| `body.theme-dark .btn-control` | `border-color` | `rgba(255,255,255,0.15)` | |
| `body.theme-dark .tab-btn:hover` | `color` | `rgba(255,255,255,0.9)` | Prevent hover revert to `#333` |
| `body.theme-dark .tab-btn:hover` | `background` | `rgba(255,255,255,0.15)` | |
| `body.theme-dark .slide-preview-label` | `color` | `rgba(255,255,255,0.65)` | |
| `body.theme-dark .build-number` | `color` | `rgba(255,255,255,0.45)` | |
| `body.theme-dark .stagetimer-container.disabled` | `color` | `rgba(255,255,255,0.55)` | |
| `body.theme-dark .stagetimer-container.disabled` | `background` | `rgba(255,255,255,0.08)` | |
| `body.theme-dark .notes-zoom-controls` | `background` | `rgba(44,44,46,0.95)` | Replace hard-coded `white` |
| `body.theme-dark input[type="text"]` | `color` | `rgba(255,255,255,0.87)` | |
| `body.theme-dark input[type="text"]` | `background` | `rgba(255,255,255,0.08)` | |
| `body.theme-dark input[type="text"]` | `border-color` | `rgba(255,255,255,0.2)` | |
| `body.theme-dark input[type="number"]` | same as text | | |
| `body.theme-dark select` | `color` | `rgba(255,255,255,0.87)` | |
| `body.theme-dark select` | `background` | `rgba(44,44,46,0.95)` | |
| `body.theme-dark select` | `border-color` | `rgba(255,255,255,0.2)` | |
| `body.theme-dark small` | `color` | `rgba(255,255,255,0.55) !important` | Override inline `style="color:#888"` |
| `body.theme-dark label` (already listed above) | `!important` | on color rule | Also overrides JS-injected `label.style.cssText = '... color: #333 ...'` at line 6981 |

### Selectors to add to `body.theme-thumb`

Same set as `theme-dark`, plus:

| Selector | Property | Value | Note |
|----------|----------|-------|------|
| `body.theme-thumb h2, body.theme-thumb h3` | `color` | `rgba(255,255,255,0.85)` | Only h1 was overridden |
| `body.theme-thumb .notes-zoom-btn` | `color` | `rgba(255,255,255,0.87)` | Already in theme-dark, missing from thumb |
| `body.theme-thumb .notes-zoom-btn` | `background` | `rgba(255,255,255,0.1)` | |
| `body.theme-thumb .notes-zoom-btn` | `border-color` | `rgba(255,255,255,0.15)` | |
| `body.theme-thumb .notes-zoom-controls` | `background` | `rgba(26,32,44,0.95)` | Match thumb background |

### JS-injected inline styles

CSS `!important` declarations beat inline `style=""` attributes that don't also use `!important`. So:

- `body.theme-dark label { color: rgba(255,255,255,0.87) !important; }` automatically covers the preset label at line 6981 (`label.style.cssText = '... color: #333 ...'`).
- The "no presets" `<div style="color: #999 ...">` (line 7003) and debug console entries (lines 6329, 7669, 7702) sit inside containers with known IDs (`#preset-buttons-container`, `#debug-console`). Target them with descendant selectors + `!important` rather than adding new classes:
  - `body.theme-dark #preset-buttons-container > div { color: rgba(255,255,255,0.45) !important; }`
  - `body.theme-dark #debug-console div, body.theme-dark #debug-console span { color: rgba(255,255,255,0.45) !important; }`

No new CSS classes or changes to JS injection sites are needed.

## Files Changed

| File | Change |
|------|--------|
| `main.js` | Add CSS overrides to `body.theme-dark` and `body.theme-thumb` blocks; add classes to 3 JS-injected inline-style locations |

## Verification

1. Switch to `theme-dark` → all labels, section headings, button text, hint text, build number readable on dark background.
2. Switch to `theme-thumb` → same, plus h2/h3/notes-zoom-btn readable.
3. Hover a `.tab-btn` in dark/thumb → text stays light, does not flash `#333`.
4. Open Settings in dark/thumb → inputs and selects have dark background, light text.
5. Disable Stagetimer → `.stagetimer-container.disabled` text is readable muted light.
6. Add presets → "Presentation 1:" label is light in dark/thumb.
7. Remove all presets → "No preset presentations configured" message is readable.
8. Open Debug Console → log entries are legible.
9. `theme-light`, `theme-max`, `theme-touch` → no visual change.
