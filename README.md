# Payment Coverage — Network Explorer

An internal lookup tool for sales: pick a region, choose a country, and see which
payment method **types** are available for pay-ins and pay-outs. License
restrictions always override provider availability.

It is a single self-contained `index.html` — no build step, no server, no
dependencies. Open it locally by double-clicking, or host it (see below).

## Features

- Region tabs (All / EMEA / APAC / Americas) with served-country counts
- Country list, searchable, A–Z, with served / restricted / no-coverage states
- Per-country detail: type chips for Payments (pay-in) and Payouts (pay-out)
- "Method names" toggle to reveal the individual methods behind each chip
- License-restricted countries are blocked outright, regardless of provider data

## Hosting on GitHub Pages

1. Create a new repository on GitHub (e.g. `payment-coverage`).
2. Add `index.html` (and this README) to it — either via `git` or by dragging
   the files into the browser on the repo's **Add file → Upload files** page.
3. In the repo: **Settings → Pages → Build and deployment**, set
   **Source = Deploy from a branch**, **Branch = main**, folder = **/ (root)**, Save.
4. After a minute the page is live at
   `https://<your-username>.github.io/<repo-name>/`. Share that link.

## Updating

The data is embedded in `index.html`. To publish an update, replace `index.html`
with the new version and commit (or re-upload it). Pages redeploys automatically.

## Notes

- Provider names in the data are anonymized (`psp1`, `psp2`, …).
- The page contains no sensitive information; anyone with the link can view it.
  If access needs to be restricted, that requires a host with authentication.
