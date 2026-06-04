# Theme Colour Editor for Google Sheets

A Google Sheets add-on for setting and saving custom theme colours across all your spreadsheets — no more typing hex codes manually every time you style a new sheet.

Colours are saved to your user profile, so once configured they carry through to every sheet you open the add-on in.

---

## Features

- **Theme colours** — set all 9 Sheets theme slots (Text, Background, Accent 1–6, Hyperlink) in one go
- **Extra colours** — inject additional colours directly into the paint-bucket custom palette using a temp-sheet workaround (no limit)
- **Native colour picker** — click the colour tile to open a visual picker, or type a hex code directly
- **Hex input** — accepts values with or without `#`
- **Live preview** — colour tiles update as you type
- **Copy to clipboard** — hover a tile and click the copy icon to grab the hex value
- **Persistent settings** — colours are saved per user and restored every time the sidebar opens
- **Clean UI** — Google-style sidebar with loading indicator and toast notifications

---

## Menu options

| Option | Description |
|---|---|
| **Set theme colours** | Applies your saved colours to the current sheet's theme |
| **Edit theme colours** | Opens the sidebar to view and update your saved colours |
| **Reset to default** | Resets all colours to the built-in defaults |

---

## Installation

This add-on is designed to be deployed within a Google Workspace organisation. It is not published to the Google Workspace Marketplace.

1. Open [script.google.com](https://script.google.com) and create a new project
2. Copy the contents of `Code.js`, `Sidebar.html`, and `appsscript.json` into your project
3. Deploy as an **Editor add-on** (Deploy → New deployment → Editor Add-on)
4. Install it on a Google Sheet via **Extensions → Add-ons → Manage add-ons**

To test without a full deployment, use **Deploy → Test deployments** and link it to a spreadsheet.

---

## Customising the defaults

Edit the `DEFAULT_SETTINGS` object at the top of `Code.js` to set the colours that appear when a user first opens the add-on or resets to default:

```js
const DEFAULT_SETTINGS = {
  "text":       "000000",
  "background": "ffffff",
  "accent1":    "50c846",
  // ... etc
  "extra1":     "FF941A",
  "extra2":     "FFE900",
  // ...
};
```

---

## FAQ

**Nothing happens when I click "Apply theme"**  
Most commonly caused by being signed into multiple Google accounts. Open the sheet from your primary account, or sign out of all others first.

**A new sheet doesn't have the right colours**  
The add-on doesn't apply colours automatically on open. Use **Extensions → Set theme colours** to apply them to the current sheet.

**I want more than 5 extra colours**  
Click **+ Add colour** in the Extra Colours section of the sidebar. There's no limit.

**Issues / feature requests**  
Please [open an issue](https://github.com/NikRpk/gas-theme-colour-editor-sheets/issues/new) and I'll get back to you.
