# How to update the Payment Coverage explorer

The app is a single `index.html`. All the data lives inside it. To change the
data you edit a spreadsheet, run one script, and re-upload `index.html`.

## One-time setup

1. Install **Python 3.8+**.
2. Install the one dependency:
   ```
   pip install openpyxl
   ```
3. Keep these four files together in one folder:
   - `build_index.py` — the build script
   - `psp_alias.json` — real provider name → `pspN` mapping (locks the numbering)
   - `index.html` — the app (used as the shell; only its data block is replaced)
   - the data workbook, e.g. `Payment_coverage.xlsx`

## Updating the data

1. Edit the workbook (in Google Sheets or Excel), then export it:
   Google Sheets → **File → Download → Microsoft Excel (.xlsx)**.
2. In the folder, run:
   ```
   python build_index.py Payment_coverage.xlsx
   ```
3. The script rewrites `index.html` in place. Upload that file to the GitHub
   repo (**Add file → Upload files → Commit**). GitHub Pages redeploys within
   about a minute at the same URL.

That's the whole loop.

## The spreadsheet

Four tabs; you mostly touch **Offerings** and **License**.

**Offerings** — one row per provider × method × direction:

| Column | Meaning |
|---|---|
| PSP | Provider name (free text; published anonymized as `psp1`, `psp2`, …) |
| Method | Display name, e.g. PIX, Cards, Apple Pay, Mobile Money |
| Type | One of: `bank`, `card`, `ewallet`, `mobile`, `cash`, `crypto` |
| Direction | One of: `in`, `out`, `both` |
| Mode | `include` = supported list · `exclude` = blocked list · `all` = everywhere |
| Countries (ISO codes) | Comma-separated ISO codes, e.g. `BR, MX, CL`. Empty when Mode = `all` |

- **Add a provider:** add its rows. It gets the next free `pspN` automatically.
- **Remove a provider:** delete its rows.
- **Change coverage:** edit the Countries cell. Look up codes on the **Countries** tab.

**License** — countries blocked outright; this overrides all provider
availability. **Un-block a country by deleting its row.** Adding a row (with an
ISO code and status `blocked`) blocks it everywhere.

## What the script does automatically

- Validates every ISO code and dropdown value. **If anything is invalid it prints
  the problems and writes nothing** — a typo can never silently ship.
- Tidies a few method names (`MobileMoney` → `Mobile Money`, `ApplePay` →
  `Apple Pay`, `GooglePay` → `Google Pay`).
- Treats `Mode = all` as "available everywhere."
- Anonymizes provider names using `psp_alias.json` and assigns new numbers to
  new providers.

## Good to know

- **Keep `psp_alias.json`.** It's the private key to which `pspN` is which
  provider, and it keeps the numbering stable. It is not needed by the live page.
- The script uses `index.html` as its template and only swaps the data block, so
  if the app itself is ever updated (new feature or layout), just drop in the new
  `index.html` and keep using the same script.
- A brand-new **method category** (beyond the six Types above) needs a code
  change to add its color — the script will stop and name the unknown type if it
  sees one.
- Options: `python build_index.py --help` (e.g. `--out` to write to a different
  file instead of overwriting `index.html`).

## Troubleshooting

- **"openpyxl is not installed"** → run `pip install openpyxl`.
- **"VALIDATION FAILED"** → read the listed rows; fix the workbook; re-run.
- **"could not find the 'const DATA = {…};' block"** → the `--template` file
  isn't a valid app `index.html`; point it at a good copy.
- **Unknown Type error** → either fix the Type to one of the six, or request a
  new category to be added.
