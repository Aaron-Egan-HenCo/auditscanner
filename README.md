# auditscanner

ScanAudit — phone barcode scanner for Barra GPS serials. Hosted on GitHub Pages; saved audits go to a shared Google Sheet via a Google Apps Script web app, and can be downloaded as `.xlsx` from any device.

## How it works

1. On the phone: enter **your name** (+ depot / notes), scan, tap **Save Audit**.
   Saving again after more scans updates the same audit (no duplicate rows).
2. **Start new audit** clears the scanner and gives a new Audit ID.
3. On any device, tap **Saved**:
   - **Recent**: download one audit, a selection, or **All scans** as Excel.
   - **By Auditor**: type a name to see that person's audits grouped into 2-day windows (e.g. 1/11 + 2/11 = one group; 3/11 starts the next). Download a group, or everything for that person.
   Multi-audit downloads list each serial once (first scan kept), and `Missing` only shows serials that weren't scanned in any audit in the download.
   Workbook sheets: `Serials` (copy from here), `Missing`, `Audits` (summary).
4. **Excel (this)** downloads the current session straight from the phone, no backend needed.

Serials that are plain digits are written as numbers in Excel; anything with a leading zero or letters stays text.

## Backend setup (one-off, ~5 min)

1. Create a new Google Sheet, e.g. **ScanAudit DB** (in the Google account that should own the data).
2. In the sheet: **Extensions → Apps Script**. Delete the default code, paste in `apps-script/Code.gs`, save.
3. Choose the `setup` function in the toolbar and click **Run**. Approve the permissions prompt. This creates the `Audits` and `Scans` tabs.
4. **Deploy → New deployment** → type **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
   → **Deploy**, then copy the Web app URL (ends in `/exec`).
5. In `index.html`, set:
   ```js
   const APPS_SCRIPT_URL = "https://script.google.com/macros/s/XXXXXXXX/exec";
   ```
   then commit and push. GitHub Pages redeploys in about a minute.

When `Code.gs` changes later: **Deploy → Manage deployments → Edit (pencil) → Version: New version → Deploy**. That keeps the same URL. Don't use "New deployment" again or the URL changes.

## Notes

- Access is name-only: anyone with the page URL can save and download. The sheet itself stays private to the owner's Google account; only the script reads or writes it.
- Delete or fix bad audits directly in the Google Sheet (remove the row in `Audits` and its rows in `Scans`).
- Scans are still stored in the phone's browser storage until you start a new audit, so nothing is lost if a save fails on bad signal. Just tap Save again.
