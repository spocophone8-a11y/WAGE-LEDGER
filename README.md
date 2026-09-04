# Wage Ledger

A labourer wage &amp; debt tracker that connects directly to an Excel file on your
computer — no server, no account, no download/upload cycle. Open the page,
connect (or create) a `.xlsx` file, and every change writes straight back to it.

## Live demo

Once this repo is published on GitHub Pages, your app will be at:

```
https://<your-username>.github.io/<repo-name>/
```

## Requirements

- **Chrome or Edge** — the direct file connection uses the File System Access
  API, which only those two browsers support. Firefox and Safari fall back to
  a manual load/download flow instead (the page detects this automatically
  and shows the right controls).
- HTTPS — GitHub Pages serves over HTTPS automatically, which this feature
  requires.

## Files

- `index.html` — page structure
- `style.css` — the ledger/paper visual style
- `script.js` — all the logic, including reading/writing the Excel file
