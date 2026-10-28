# Delegate Rooming Monitor

A static monitoring dashboard for delegate hotel rooming, published with GitHub Pages.
It shows aggregate counts only: delegates expected and arrived per hotel, check-ins and
check-outs per day, delegates in-house per night, flight arrival time bands, minors,
SFTS, late-night and early arrivals, and branch / language / airline mix.

No delegate names, delegate codes, group codes, flight numbers, kit boxes or remarks are
ever written to the repository. The source spreadsheets are git-ignored and the Pages
workflow refuses to deploy if one is present.

The page also carries a **Delegates** section with one row per delegate: hotel status
from the Arrived column, welcome kit, HID remarks, check-in and flight arrival. Names
are masked to initials until the viewer searches or presses "Show names". These rows
exist only inside the encrypted data file; delegate codes, group codes and kit boxes
are never published.

Without the password the page looks like a plain "404 Not Found" page. Tapping the
large 404 three times, or typing `open`, reveals the password box. A device that
ticked "Keep me signed in" skips both.

The published figures are **password protected**: the build script encrypts them
(AES-256-GCM, key derived from the password with PBKDF2-SHA256) and the page decrypts
them in the browser after the viewer enters the password. Only the encrypted file
`data/summary.enc.js` is committed; unencrypted data files are git-ignored and the
workflow refuses to deploy if one is present.

## Live page

Once GitHub Pages is enabled (see below) the dashboard is served at
`https://ic-hc1.github.io/delegate-rooming/`.

## Layout

| Path | Purpose |
| --- | --- |
| `index.html` | The dashboard. Plain HTML, CSS and JavaScript, no build step, no external dependencies. |
| `data/summary.enc.js` | Encrypted aggregate data the page loads and decrypts with the password. Generated. |
| `scripts/build_data.py` | Reads the hotel master-list workbooks and writes the encrypted data file. |
| `vendor/crypto-fallback.js` | Pure-JavaScript PBKDF2 and AES-GCM (bundled from the audited `@noble` libraries) used only when a browser's own WebCrypto is unavailable or fails. |
| `.github/workflows/pages.yml` | Deploys the repository root to GitHub Pages on every push to `main`. |

## Updating the numbers

1. Put the latest `*_Hotel_Delegates_Master_List.xlsx` files in a folder outside the repo,
   or in `input/` inside it (that folder is git-ignored).
2. Run the build script with the dashboard password (needs Python 3, `openpyxl` and
   `cryptography`):

   ```bash
   pip install openpyxl cryptography
   python3 scripts/build_data.py input/ --password 'the-dashboard-password'
   ```

   You can also put the password in the `DASHBOARD_PASSWORD` environment variable instead
   of typing it on the command line. The script reads the `Master List` sheet of each
   workbook, counts `Arrived` and `Welcome Kit` as given wherever the cell says `Yes`,
   excludes the grey cancelled rows, and aborts if any person-level column would reach
   the output.
   If the planners' *Change of flights* workbook is available, pass it too so rebooked
   delegates show their new arrival and are not counted as overdue:

   ```bash
   python3 scripts/build_data.py input/ --changes input/Changed_Flights.xlsx --password '...'
   ```

   A planners' cancellation list (any workbook with a `Delegate Code` column) excludes
   those delegates from every count and lists them under a Cancelled tab:

   ```bash
   python3 scripts/build_data.py input/ --cancelled input/CANCELLED.xlsx --password '...'
   ```

   The registration export (`Delegates-YYYY-MM-DD.csv`) can be passed as the registered
   base; the page then reconciles it with the hotel lists and every cancellation on record:

   ```bash
   python3 scripts/build_data.py input/ --registry input/Delegates-2026-10-24.csv --cancelled input/CANCELLED.xlsx --password '...'
   ```

3. Commit and push `data/summary.enc.js`. The Pages workflow redeploys the site
   automatically.

## Changing the password

Rebuild with the new password and push. Viewers who ticked "Keep me signed in on this
device" are signed out automatically because their saved key no longer opens the new
file. There is no password recovery: the figures can only be read with the password the
file was built with, so keep it somewhere safe. Use the **Lock** button in the page header
to sign out on a shared device.

## One-time setup: enable GitHub Pages

In the repository on GitHub go to **Settings → Pages** and under **Build and deployment**
set **Source** to **GitHub Actions**. The next push to `main` (or a manual run of the
"Deploy dashboard to GitHub Pages" workflow from the Actions tab) publishes the site.

## Viewing locally

Open `index.html` directly in a browser and enter the password, or serve the folder:

```bash
python3 -m http.server 8000
```

For an unprotected local preview, build with `--plain` as well; that writes
`data/summary.js`, which the page uses automatically when present and which is
git-ignored so it can never be published.

Then open <http://localhost:8000/>. Add `#<hotel-key>` to the URL (for example
`#okada-manila`) to open the page already focused on one hotel.
