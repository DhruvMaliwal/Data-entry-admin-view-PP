# Property Data Entry Portal — Setup Guide

## 1. Create the Master Google Sheet

1. Go to [Google Sheets](https://sheets.google.com) and create a new spreadsheet.
2. Name it **Property Master Database** (or whatever you prefer).
3. The app automatically creates two sheets inside it:
   - **Households** — stores the list of houses (House ID, House Name, Address, Created Date)
   - **Data Entries** — stores all uploaded data, tagged by House ID and Upload Date
4. **Get the Sheet ID** from the URL:
   ```
   https://docs.google.com/spreadsheets/d/THIS_IS_YOUR_SHEET_ID/edit
   ```
5. Paste it into `Code.gs` line 16 as `MASTER_SHEET_ID`.

## 2. Set Up the Apps Script Project

1. Go to [Google Apps Script](https://script.google.com) → **New project**.
2. Replace `Code.gs` contents with the provided file.
3. Click **+** → **HTML** → name it `Index` → paste the provided `Index.html`.
4. Enable **Drive API**: click **Services** (+) → add **Drive API** (needed for Excel parsing).
5. Update `MASTER_SHEET_ID` in `Code.gs`.

## 3. Deploy as a Web App

1. Click **Deploy** → **New deployment** → **Web app**.
2. Execute as: **Me** | Who has access: **Anyone** (or your domain).
3. Click **Deploy**, authorize when prompted.
4. Share the web app URL with your data entry team.

## How It Works

1. **Landing page** shows all households — click one to select it, or add a new one.
2. **Upload page** — choose a `.csv` or `.xlsx` data entry file for that household.
3. **Dashboard** — after upload, see a summary with stats and a table of all uploaded rows.
4. All data is stored in the master Google Sheet, tagged by House ID and date.

## Troubleshooting

| Problem | Fix |
|---------|-----|
| "Unsupported file type" | Upload `.csv`, `.xlsx`, or `.xls` only. |
| "No data rows found" | File needs a header row + at least one data row. |
| Excel parsing fails | Enable Drive API advanced service (step 2.4). |
| Permission denied | Re-deploy and re-authorize the app. |
