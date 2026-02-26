# GSheet-html

Visual Schedule Dashboard pilot for Google Sheets + Google Apps Script + Google Sites.

## Files

- `Code.gs`: Backend endpoint (`doGet`) with two app views, plus data APIs and 5-minute caches.
- `Index.html`: Frontend dashboard UI (calendar + table + filters + search + modal).
- `IndexMECOE.html`: Frontend timeline matrix generated from task date ranges.
- `templates/standard/schedule_setup.gs`: Setup/validation helper for the `Schedule` sheet schema.
- `templates/standard/schedule_template.csv`: Minimal starter data aligned to the standard schema.
- `templates/mecoe/mecoe_setup.gs`: Setup helper for normalized `ME_COE_Schedule`.
- `templates/mecoe/ME_COE_Schedule-template.csv`: ME COE template CSV.
- `USER_GUIDE.md`: End-to-end setup, deployment, and troubleshooting guide for users.
- `HTML.export.md`: Pilot PRD.

## Required Sheet Schema

Sheet name: `Schedule`

Headers (exact):

1. `ID`
2. `Title`
3. `Start Date`
4. `End Date`
5. `Owner`
6. `Department`
7. `Status`
8. `Description`
9. `Tags`

## Deploy (Google Apps Script)

1. Open your target Google Sheet.
2. Go to `Extensions -> Apps Script`.
3. Add `Code.gs` and `Index.html` from this repo.
4. Optional: add `templates/standard/schedule_setup.gs` and run `setupScheduleSheet()` once.
5. Click `Deploy -> New deployment -> Web app`.
6. Execute as: `Me`.
7. Who has access: your internal domain users.
8. Copy the Web App URL and embed it in Google Sites.

## App Routes

- Standard schedule dashboard:
  - `WEB_APP_URL`
- ME COE timeline dashboard:
  - `WEB_APP_URL?app=mecoe`

## ME COE Template Notes

- Expected sheet names (first found is used):
  1. `ME_COE_Schedule`
  2. `ME_COE`
  3. `Schedule_ME_COE`
- Recommended schema (exact headers):
  1. `ID`
  2. `Phase`
  3. `Deliverable / Document`
  4. `Site`
  5. `Docs`
  6. `Priority`
  7. `Status`
  8. `Owner`
  9. `Start Date`
  10. `End Date`
  11. `Milestone Date`
  12. `Notes / Actions`
- Timeline buckets are generated automatically from min/max dates.
- `Today` is generated at runtime as the benchmark marker.
- Legacy wide format is still readable as fallback, but date-model is preferred.

### Quick Setup For ME COE Sheet

1. In Apps Script, add `templates/mecoe/mecoe_setup.gs`.
2. Run `setupMECOEScheduleSheet()` once.
3. Fill rows starting from row 2.
4. Open `WEB_APP_URL?app=mecoe`.

## Behavior

- Calendar views: Month + Week.
- Table features: sorting, global search, owner/department/status filters, pagination.
- Modal detail view on click (calendar event or table row).
- Refresh button bypasses cache.
- Server cache TTL: 300 seconds (max 5 minutes).
