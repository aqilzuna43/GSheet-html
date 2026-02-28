# User Guide: ME COE Timeline Dashboard

This guide is for the ME COE dashboard only.

## 1. Prerequisites

1. Google account with access to Google Sheets and Apps Script
2. A Google Sheet with ME COE schedule data
3. Files from this repo:
   - `CodeMECOE.gs`
   - `IndexMECOE.html`
   - `templates/mecoe/mecoe_setup.gs` (optional helper)

## 2. Create Apps Script Project

1. Open your ME COE Google Sheet.
2. Go to `Extensions -> Apps Script`.
3. In Apps Script:
   1. Add `CodeMECOE.gs` (you may rename it to `Code.gs` in that project)
   2. Add `IndexMECOE.html`
   3. Optional: add `templates/mecoe/mecoe_setup.gs`
4. Save all files.

## 3. Setup ME COE Sheet

Option A (recommended):
1. Run `setupMECOEScheduleSheet`.
2. Fill `ME_COE_Schedule` rows starting from row 2.

Option B (existing legacy matrix):
1. Keep legacy sheet (`ME_COE`).
2. Run `migrateLegacyMECOEToDateModel("ME_COE")` if needed.

Date-model headers:
- `ID`, `Phase`, `Deliverable / Document`, `Site`, `Docs`, `Priority`, `Status`, `Owner`, `Start Date`, `End Date`, `Milestone Date`, `Notes / Actions`

## 4. Deploy Web App

1. `Deploy -> New deployment`
2. Type: `Web app`
3. Set execute/access based on your internal policy
4. Deploy and copy URL

## 5. Open Dashboard

Use the web app URL directly (no `?app=` switch needed in split deployment).

## 6. Troubleshooting

1. `Could not find IndexMECOE.html`
   - Add/save `IndexMECOE.html`.

2. `ME COE sheet not found`
   - Ensure one of these tabs exists:
     - `ME_COE_Schedule`
     - `ME_COE`
     - `Schedule_ME_COE`

3. Legacy timeline not rendering
   - Verify timeline columns contain recognizable month labels or `now`.

## 7. Important Deployment Rule

This guide is for ME COE only.  
Do not include the general pilot `Code.gs` / `Index.html` in this Apps Script project.
