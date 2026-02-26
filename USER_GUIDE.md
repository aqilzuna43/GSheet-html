# User Guide: Google Sheets + HTML Dashboards

This guide explains how to set up and run both dashboards in Google Apps Script:

- Standard Schedule Dashboard
- ME COE Gantt Dashboard

## 1. Prerequisites

1. Google account with access to Google Sheets and Apps Script
2. One Google Sheet file to host data
3. The project files from this repo:
   - `Code.gs`
   - `Index.html`
   - `IndexMECOE.html`
   - `templates/standard/schedule_setup.gs` (optional helper)
   - `templates/mecoe/mecoe_setup.gs` (optional helper)

## 2. Create Apps Script Project

1. Open your Google Sheet.
2. Go to `Extensions -> Apps Script`.
3. In the Apps Script editor:
   - Create/replace `Code.gs`
   - Add `Index.html`
   - Add `IndexMECOE.html`
   - Optional: add `templates/standard/schedule_setup.gs`
   - Optional: add `templates/mecoe/mecoe_setup.gs`
4. Click `Save`.

## 3. Setup Standard Schedule Template

Option A (recommended): run setup helper

1. In Apps Script, select function `setupScheduleSheet`.
2. Click `Run`.
3. Authorize script permissions if prompted.
4. Open sheet tab `Schedule` and start entering rows under the header.

Option B: manual setup

1. Create sheet named `Schedule`.
2. Add headers exactly:
   - `ID`, `Title`, `Start Date`, `End Date`, `Owner`, `Department`, `Status`, `Description`, `Tags`

## 4. Setup ME COE Template (Date Model)

Option A (recommended): run setup helper

1. In Apps Script, select function `setupMECOEScheduleSheet`.
2. Click `Run`.
3. Authorize script permissions if prompted.
4. Open sheet tab `ME_COE_Schedule`.
5. Fill rows from row 2 using columns:
   - `ID`, `Phase`, `Deliverable / Document`, `Site`, `Docs`, `Priority`, `Status`, `Owner`, `Start Date`, `End Date`, `Milestone Date`, `Notes / Actions`

Option B: migrate legacy wide sheet metadata

1. If you have old matrix format in `ME_COE` (or another tab), run:
   - `migrateLegacyMECOEToDateModel("ME_COE")`
2. Fill `Start Date` / `End Date` / `Milestone Date` manually after migration.

## 5. Deploy Web App

1. In Apps Script, click `Deploy -> New deployment`.
2. Type: `Web app`
3. Set:
   - Execute as: `Me`
   - Who has access: your internal domain users (or as needed)
4. Click `Deploy`.
5. Copy the Web App URL.

## 6. Open Each Dashboard

1. Standard dashboard:
   - `WEB_APP_URL`
2. ME COE Gantt dashboard:
   - `WEB_APP_URL?app=mecoe`

## 7. Embed in Google Sites

1. Open Google Sites page.
2. Insert -> Embed -> `By URL`.
3. Paste either dashboard URL.
4. Publish site.

## 8. Updating Code

After any code change:

1. Save all Apps Script files.
2. `Deploy -> Manage deployments`.
3. Edit deployment and choose `New version`.
4. Deploy.
5. Refresh browser page.

## 9. Troubleshooting

1. `Script function not found: doGet`
   - `Code.gs` is missing or unsaved, or old deployment version is still active.
   - Save files and redeploy a new version.

2. `Identifier ... has already been declared`
   - Two script files use the same global constant/function name.
   - Keep unique names across all `.gs` files.

3. Blank dashboard
   - Verify sheet name exists (`Schedule` or `ME_COE_Schedule`).
   - Verify headers match exactly.
   - Ensure date columns are valid dates (`yyyy-mm-dd` works well).

4. Data not refreshing
   - Click dashboard `Refresh`.
   - Cache TTL is up to 5 minutes unless forced refresh is used.

## 10. Daily Use Tips

1. Keep headers unchanged.
2. Prefer one task per row.
3. Keep dates populated for timeline/Gantt visibility.
4. Use filters/search for faster navigation.
