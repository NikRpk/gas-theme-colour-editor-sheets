# PROJECT_CONTEXT.md — hf-theme-colour-editor

## Overview

A Google Sheets add-on (Google Apps Script) that lets HelloFresh team members set and persist custom theme colours across all their spreadsheets. It provides a sidebar UI with a native colour picker, hex input, live preview, and an "extra colours" section that injects custom colours into Sheets' built-in paint-bucket palette via a temporary-sheet workaround.

**Status:** Active / deployed as a Sheets add-on  
**GitHub:** https://github.com/NikRpk/gas-theme-colour-editor-sheets  
**Real source dir:** `/Users/niklas.roepke/Developer/theme-colours`  
**Deployed via:** clasp (`scriptId: 1qVB0k9qMKg0N1qoi4J3HSXt2AmeBiXXpj6HRKu7Fig1aTuDvZ7a18NQz`)

---

## Tech Stack

| Layer | Technology |
|---|---|
| Runtime | Google Apps Script (GAS), V8 engine |
| Client UI | GAS HTML Service — `Sidebar.html` (plain HTML/CSS/JS, no build step) |
| Server | `Code.js` — plain GAS JavaScript |
| Persistence | `PropertiesService.getUserProperties()` — per-user, cross-spreadsheet |
| Deployment | `clasp` CLI → script.google.com |
| Fonts | Google Fonts (Roboto, Google Sans) loaded via `<link>` in sidebar |

---

## Key Concepts

### Theme slots vs. extra colours

- **9 theme slots** (`text`, `background`, `accent1`–`accent6`, `hyperlink`) map to native `SpreadsheetApp.ThemeColorType` enum values and are applied with `ss.getSpreadsheetTheme().setConcreteColor(...)`.
- **Extra colours** (`extra1`, `extra2`, … `extraN`) have no native API. They are injected into Sheets' "custom colours" paint-bucket palette via a temp-sheet trick: create a hidden sheet, paint each hex as a cell background, `SpreadsheetApp.flush()`, then delete the sheet. This causes one brief tab-flicker.

### Extra colour count tracking

- `extraCount` user property stores how many rows were visible when last saved, so the sidebar can restore the exact same number of rows on next open.
- Minimum is `DEFAULT_EXTRA_COUNT = 5`.
- When a user removes a row in the sidebar and saves, the removed `extraN` properties must be **deleted** from UserProperties — otherwise they reappear on next open.

### Sidebar row keys

- Extra rows are keyed sequentially as `extra1`, `extra2`, … in the DOM, but rows can be removed in any order (leaving DOM gaps).
- `grabtext()` must collect visible extras by iterating actual DOM elements (not by sequential index) and remap them to contiguous keys `extra1`, `extra2`, … before sending to the server.
- The server always receives a contiguous, gap-free dict of `extra1`…`extraN`, then deletes any old `extraK` properties beyond N.

---

## Configuration

```js
// Code.js
const DEFAULT_SETTINGS = {
  text: "000000", background: "ffffff",
  accent1: "50c846", accent2: "009646",
  accent3: "ff5f64", accent4: "d9d9d9",
  accent5: "ff941a", accent6: "1464ff",
  hyperlink: "009646",
  extra1: "FF941A", extra2: "FFE900", extra3: "FF63AA",
  extra4: "FEF8F0", extra5: "232323"
};
const DEFAULT_EXTRA_COUNT = 5;
```

OAuth scopes required:
- `https://www.googleapis.com/auth/script.container.ui`
- `https://www.googleapis.com/auth/spreadsheets`

---

## Workflow

```bash
# Deploy changes
cd /Users/niklas.roepke/Developer/theme-colours
clasp push

# Create a new version (note the version number printed)
clasp version "v{X.Y} - description"

# Redeploy the versioned deployment to the new version number
clasp deploy \
  --deploymentId AKfycbw0sJxH8dg3Oc_S0cn6-4UCGZT6NGNE5hIvVBt6Q2v9XEshY5F4r_Q5lbdbiPtIMswc \
  --versionNumber {N} \
  --description "v{X.Y} - description"

# !! REQUIRED — update the Marketplace SDK version number !!
# Go to: console.cloud.google.com
# → APIs & Services → Google Workspace Marketplace SDK → App Configuration
# → "Sheets add-on script version" field → set to {N}
# Without this step, the published add-on still serves the old version.
```

The add-on must be opened via **Extensions → Theme Colour Editor → Edit theme colours** in a Google Sheet (not the Apps Script editor) to test the sidebar.

---

## Common Errors

| Error | Root cause | Fix |
|---|---|---|
| Deleted extra rows reappear on next open | `saveColour` only saves submitted keys; orphaned `extraK` properties remain in UserProperties | Delete all `extraK` properties above the new count in `saveColour` |
| Gaps in extra key numbering | `grabtext()` iterates by index and skips removed rows, sending `extra1, extra3` instead of `extra1, extra2` | Iterate DOM elements, remap to contiguous keys |
| Tab flicker on save | `injectExtraColours_` creates and deletes a temp sheet — unavoidable with current API | Expected behaviour; document in UI |
| `ThemeColorType[dtype.toUpperCase()]` returns undefined | Unknown dtype string passed to `setThemeColour_` | Logs to Stackdriver; does not throw |

---

## Recent Changes

- **2026-06-08** — Added `PROJECT_CONTEXT.md`. Fixed delete-extras bug: `saveColour` now deletes orphaned `extraK` properties; `grabtext()` now remaps DOM rows to contiguous keys.
