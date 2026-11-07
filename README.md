# Dataset to Google Sheets, CSV & Excel Exporter

Export any Apify dataset, or a public CSV / JSON / JSON Lines URL, to:
- **Google Sheets**, using *your own* Google service account (no third-party app gets access to your Google account)
- **CSV** (UTF-8 with BOM, so Excel opens it correctly) and **XLSX**, saved to the run's key-value store

## Why this one
- **Checks before it writes.** Validates your key, checks Google Sheets' 10M-cell limit, and fails early with a clear message instead of half-writing a sheet.
- **Plain-English errors.** "Share the spreadsheet with bot@your-project.iam.gserviceaccount.com as an Editor" instead of `403 PERMISSION_DENIED`.
- **Handles big exports.** Writes in batches, with automatic retry and backoff when Google rate-limits.
- **Flattens nested JSON** into `dot.path` columns, with a column cap.
- **Your key stays secret.** The input is encrypted by Apify, used only for the run, and removed from every log and error message.

## Setup (about 5 minutes, one time)
1. In Google Cloud Console, create a project and enable the **Google Sheets API**.
2. IAM & Admin > Service Accounts > Create. Then Keys > Add key > JSON, and download the file.
3. Open your spreadsheet > **Share**, add the service account's email (`...@...iam.gserviceaccount.com`) as **Editor**.
4. Paste the JSON key and the spreadsheet ID (the part of the URL between `/d/` and `/edit`) into this Actor.

Leave the spreadsheet ID empty to get CSV/XLSX files only (no Google setup needed).

## Pricing
Free for the first 2,000 rows of each run. After that, $0.25 per 1,000 rows exported, plus $0.01 per run that writes to Google Sheets (pay-per-event). You're only charged for rows actually exported, and Apify's per-run spending cap applies.

## Limits (v1)
- Rows are loaded into memory before writing. Very large datasets need a bigger memory setting.
- Google Sheets caps a spreadsheet at 10M cells and a cell at 50k characters (longer text is cut).
- The target spreadsheet must already exist and be shared with the service account. Service accounts can't own Drive files, so this Actor never creates new spreadsheets.
- CSV/XLSX cells that start with `= + - @` are prefixed with `'` to block formula injection. Plain numbers such as `-5` are left alone.
- **Numbers:** in Excel and Google Sheets, a column whose values are all plain numbers (`22`, `-7.25`) is written as real numbers. Mixed columns, leading zeros (`02134`), integers over 11 digits (barcodes, IDs) and decimals over 15 significant digits stay text. Trailing zeros are dropped (`1.50` becomes `1.5`). Turn off **Convert numbers** to keep everything as text. CSV output is never changed. In Append mode the decision is made per run, so a column can be numbers in one run and text in another. Keep the setting the same across runs, or turn it off, if that matters.
- Append mode requires the tab's header row to match the export's columns.
- Replace mode clears the tab before writing, so a failure partway through can leave it half-filled. Use a dedicated tab for exports, not one holding data you edit by hand.
