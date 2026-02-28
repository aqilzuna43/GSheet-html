# User Guide: General Pilot Dashboard

This guide is for the general pilot dashboard only (standard schedule template).

## 1. Prerequisites

1. Google account with access to Google Sheets and Apps Script
2. One or more Google Sheets used as schedule sources
3. Files from this repo:
   - `Code.gs`
   - `Index.html`
   - `templates/standard/schedule_setup.gs` (optional helper)

## 2. Create Apps Script Project

1. Open your host Google Sheet.
2. Go to `Extensions -> Apps Script`.
3. In Apps Script:
   1. Add/replace `Code.gs`
   2. Add `Index.html`
   3. Optional: add `templates/standard/schedule_setup.gs`
4. Save all files.

## 3. Prepare Source Sheet

Option A (recommended):
1. Run `setupScheduleSheet`.
2. Use sheet tab `Schedule`.

Option B (manual):
1. Create sheet tab `Schedule`.
2. Add row-1 headers exactly:
   - `ID`, `Title`, `Start Date`, `End Date`, `Owner`, `Department`, `Status`, `Description`, `Tags`

### Status Standard (Traffic-Light)

Use only these `Status` values:
- `Not Started`
- `In Progress`
- `At Risk`
- `Blocked`
- `Completed`

PMO interpretation:
- `Completed` = Green
- `In Progress` / `Not Started` = Amber-monitoring
- `At Risk` / `Blocked` = Red attention

### Tag Standard (How To Use `Tags`)

Purpose:
- Tags improve search, filtering, and reporting context.

Format:
- Comma-separated labels in a single cell.
- Example: `design, supplier, gate-2, risk-high`

Rules:
- Use lowercase.
- Use short, stable labels.
- Avoid duplicates.
- Keep 3-6 tags per row.

## 4. Configure Default Source (Optional)

You can preconfigure a default source in a `Config` sheet (key/value format):

1. Create sheet tab `Config`.
2. Add:
   - `standard_source_spreadsheet_id` -> `<spreadsheet id>`
   - `standard_source_sheet_name` -> `Schedule` (or your preferred tab)

If not set, the app uses the active bound spreadsheet and `Schedule`.

## 5. Deploy Web App

1. `Deploy -> New deployment`
2. Type: `Web app`
3. Set:
   - Execute as: `Me`
   - Who has access: your target users/domain
4. Click `Deploy` and copy URL.

## 6. Let Users Choose Source by Preference

In the dashboard toolbar:

1. `Source Spreadsheet ID (optional)`:
   - Leave blank to use configured default/active spreadsheet.
   - You can paste either:
     - Spreadsheet ID only, or
     - Full Google Sheets URL (ID is extracted automatically).
2. `Source Sheet`:
   - Enter sheet tab name (for example `Schedule`).
3. Click `Apply Source`.

Behavior:
- Preference is saved in browser local storage per user/browser.
- `Reset Source` clears personal preference and reverts to configured default.

## 7. Embed in Google Sites

1. Open Google Sites.
2. `Insert -> Embed -> By URL`.
3. Paste web app URL.
4. Publish.

## 8. Troubleshooting

1. `Could not find Index.html`
   - Add/save `Index.html` in Apps Script.

2. `Missing required column: ...`
   - Fix row-1 headers to match required schema exactly.

3. `Cannot open source spreadsheet by ID`
   - Check sharing permissions for the account running the web app.
   - Confirm the ID/URL points to a real Google Sheet.
   - Click `Reset Source` to clear saved local preference if needed.

4. `Sheet "X" not found`
   - Verify source sheet tab name.

## 9. Important Deployment Rule

This guide is for general pilot only.  
Do not include `CodeMECOE.gs` / `IndexMECOE.html` in this Apps Script project.
