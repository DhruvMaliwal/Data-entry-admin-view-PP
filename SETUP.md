# Property Data Entry Portal — Setup Guide

## 1. Create the Master Google Sheet

1. Go to [Google Sheets](https://sheets.google.com) and create a new spreadsheet.
2. Name it something like **Property Master Database**.
3. The app will automatically write headers on first run, but you can add them yourself in row 1:

   | A | B | C | D | E | F | G |
   |---|---|---|---|---|---|---|
   | House ID | Address | Price | Beds | Baths | Description | Images |

4. **Get the Sheet ID** — open the sheet and look at the URL:
   ```
   https://docs.google.com/spreadsheets/d/THIS_IS_YOUR_SHEET_ID/edit
   ```
   Copy the long string between `/d/` and `/edit`.

5. Paste it into `Code.gs` as the value of `MASTER_SHEET_ID`.

## 2. Create the Google Drive Folder for Images

1. Go to [Google Drive](https://drive.google.com) and create a new folder (e.g. **Property Images**).
2. Open the folder. The URL will look like:
   ```
   https://drive.google.com/drive/folders/THIS_IS_YOUR_FOLDER_ID
   ```
   Copy the ID after `/folders/`.

3. Paste it into `Code.gs` as the value of `IMAGE_FOLDER_ID`.

## 3. Set Up the Apps Script Project

1. Go to [Google Apps Script](https://script.google.com) and click **New project**.
2. Rename the project to **Property Data Entry Portal**.
3. Replace the contents of the default `Code.gs` with the provided `Code.gs` file.
4. Click **+** next to **Files** → **HTML** → name it `Index` (without the `.html` extension — Apps Script adds it automatically).
5. Replace the contents with the provided `Index.html` file.

### Enable the Drive API (required for Excel parsing)

The Excel-to-Sheet conversion uses the **Drive API v2** (advanced service):

1. In the Apps Script editor, click **Services** (+ icon) on the left sidebar.
2. Find **Drive API** and click **Add**.
3. Leave the identifier as `Drive` and version as `v2`.

## 4. Deploy as a Web App

1. Click **Deploy** → **New deployment**.
2. Click the gear icon next to **Select type** → choose **Web app**.
3. Set:
   - **Description**: Property Data Entry Portal
   - **Execute as**: **Me** (your account — so the script has permission to write to your Sheet and Drive)
   - **Who has access**: **Anyone** (or restrict to your Google Workspace domain)
4. Click **Deploy**.
5. Authorize the app when prompted (it needs access to Sheets and Drive).
6. Copy the **Web app URL** — share this with your data entry team.

## 5. Prepare Your Data File

Create an Excel or CSV file with these columns (header row required):

| House ID | Address | Price | Beds | Baths | Description |
|----------|---------|-------|------|-------|-------------|
| H101 | 123 Main St | 450000 | 3 | 2 | Renovated ranch |
| H102 | 456 Oak Ave | 620000 | 4 | 3 | Colonial with pool |

- **House ID** is the key used to match images to rows.
- Extra columns beyond these six are ignored (not written to the sheet).

## 6. Name Your Image Files

Image filenames must start with the House ID followed by an underscore:

```
H101_front.jpg
H101_kitchen.png
H102_exterior.jpg
```

- Everything before the first `_` is treated as the House ID.
- Images that don't match any House ID in the data file are skipped (a warning is shown).

## Troubleshooting

| Problem | Fix |
|---------|-----|
| "Unsupported file type" | Make sure the data file is `.csv`, `.xlsx`, or `.xls`. |
| "No data rows found" | Check that your file has a header row and at least one data row. |
| Images not matching | Verify filenames start with the exact House ID and an underscore. |
| Excel parsing fails | Ensure the Drive API advanced service is enabled (see step 3 above). |
| Permission denied | Re-deploy and re-authorize. The script must run as your account. |
