# Web Remote V2-C — Cursor handoff bundle

Drop this whole folder anywhere in the `gslide-opener` repo (suggested: `docs/plans/web-remote-v2c/`) and point Cursor at `plan.md`.

## Contents

- `plan.md` — the implementation plan. Self-contained; references in §1 point to files in this same folder.
- `Web Remote UI.html` — open in a browser to see the V2-C target. Scroll to the artboard labeled **V2-C — Collapsed previews + bigger notes + inline timer message**.
- `design-canvas.jsx` — canvas scaffolding used by the HTML.
- `components/web-remote-variants.jsx` — source for the V0/V1/V2/V3 artboards.
- `components/v2-explorations.jsx` — source for the V2-A/B/C artboards (V2-C is the target).

## How Cursor should use this

1. Open `Web Remote UI.html` in a browser once to see the target.
2. Open `plan.md` and walk §0 → §12 in order. Use the search anchors in §2 — line numbers are intentionally not used because they drift.
3. Commit in the order listed in §12.
4. Run the smoke checklist in §10.

## Nothing here is app code

None of the files in this bundle should be loaded by the Electron app at runtime. The `.jsx` and `.html` files are design reference only. The only deliverable is edits to `main.js` as described in the plan.
