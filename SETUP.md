# Property Data Entry Portal — Setup

## 1. Create the Google Sheet
1. Go to [sheets.google.com](https://sheets.google.com) → new spreadsheet.
2. Copy the Sheet ID from the URL:
   `https://docs.google.com/spreadsheets/d/`**`SHEET_ID`**`/edit`
3. The app auto-creates three tabs inside it: **Households**, **Entries**, **Entry Data**.

## 2. Deploy the Apps Script
1. Go to [script.google.com](https://script.google.com) → **New project**.
2. Replace `Code.gs` with the provided file; paste your Sheet ID into `MASTER_SHEET_ID`.
3. Click **+** → **HTML** → name it `Index` → paste the provided `Index.html`.
4. **Deploy** → **New deployment** → **Web app** → Execute as **Me**, Access **Anyone** → **Deploy**.
5. Authorize, then share the web app URL with your team.

## Flow
1. **Households** — landing page shows all houses. Add a new one by name only.
2. **House page** — click a house to see its previous entries and upload a new sheet.
3. **Dashboard** — every upload creates a timestamped entry with stats + full data table.
4. **History** — clicking any past entry re-opens its dashboard. Data never vanishes.

## Notes
- Excel files are parsed in the browser (no Drive API needed).
- Every entry is tagged with an auto-generated Entry ID and timestamp.
- Removing a house deletes its metadata, entry index, and all data rows.
