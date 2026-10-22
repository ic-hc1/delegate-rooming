# Delegate Rooming Monitor

A static monitoring dashboard for delegate hotel rooming, published with GitHub Pages.
It shows aggregate counts only: delegates expected and arrived per hotel, check-ins and
check-outs per day, delegates in-house per night, flight arrival time bands, minors,
SFTS, late-night and early arrivals, and branch / language / airline mix.

No delegate names, delegate codes, group codes, flight numbers, kit boxes or remarks are
ever written to the repository. The source spreadsheets are git-ignored and the Pages
workflow refuses to deploy if one is present.

## Live page

Once GitHub Pages is enabled (see below) the dashboard is served at
`https://mac-milan.github.io/delegate-rooming-monitoring/`.

## Layout

| Path | Purpose |
| --- | --- |
| `index.html` | The dashboard. Plain HTML, CSS and JavaScript, no build step, no external dependencies. |
| `data/summary.js` | Aggregate data the page loads (`window.ROOMING_DATA`). Generated. |
| `data/summary.json` | The same aggregates as plain JSON, for anyone who wants the numbers. Generated. |
| `scripts/build_data.py` | Reads the hotel master-list workbooks and writes the two data files. |
| `.github/workflows/pages.yml` | Deploys the repository root to GitHub Pages on every push to `main`. |

## Updating the numbers

1. Put the latest `*_Hotel_Delegates_Master_List.xlsx` files in a folder outside the repo,
   or in `input/` inside it (that folder is git-ignored).
2. Run the build script (needs Python 3 and `openpyxl`):

   ```bash
   pip install openpyxl
   python3 scripts/build_data.py input/
   ```

   The script reads the `Master List` sheet of each workbook, counts `Arrived` and
   `Welcome Kit` as given wherever the cell says `Yes`, excludes the grey cancelled rows,
   and aborts if any person-level column would reach the output.
3. Commit and push `data/summary.js` and `data/summary.json`. The Pages workflow
   redeploys the site automatically.

## One-time setup: enable GitHub Pages

In the repository on GitHub go to **Settings → Pages** and under **Build and deployment**
set **Source** to **GitHub Actions**. The next push to `main` (or a manual run of the
"Deploy dashboard to GitHub Pages" workflow from the Actions tab) publishes the site.

## Viewing locally

Open `index.html` directly in a browser, or serve the folder:

```bash
python3 -m http.server 8000
```

Then open <http://localhost:8000/>. Add `#<hotel-key>` to the URL (for example
`#okada-manila`) to open the page already focused on one hotel.
