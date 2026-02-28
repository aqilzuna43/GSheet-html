# GSheet-html

This repository now supports two separate Google Apps Script web apps:

1. General Pilot Dashboard (mass users)
2. ME COE Timeline Dashboard (internal/specialized users)

Do not deploy both backends in the same Apps Script project.

## General Pilot App

Files to include in Apps Script:
- `Code.gs`
- `Index.html`
- `templates/standard/schedule_setup.gs` (optional)

Primary API:
- `getScheduleItems(forceRefresh, sourcePrefs)`

Sheet model:
- `Schedule`
- Headers: `ID`, `Title`, `Start Date`, `End Date`, `Owner`, `Department`, `Status`, `Description`, `Tags`

Source preference support:
- Users can choose source Spreadsheet ID and source sheet in the UI.
- Optional defaults can be set in a `Config` sheet:
  - `standard_source_spreadsheet_id`
  - `standard_source_sheet_name`

Guide:
- See `USER_GUIDE.md`

## ME COE App

Files to include in Apps Script:
- `CodeMECOE.gs` (rename to `Code.gs` inside the ME COE Apps Script project if desired)
- `IndexMECOE.html`
- `templates/mecoe/mecoe_setup.gs` (optional)

Primary API:
- `getMECOESchedule(forceRefresh)`

Sheet model candidates:
1. `ME_COE_Schedule`
2. `ME_COE`
3. `Schedule_ME_COE`

Guide:
- See `USER_GUIDE_MECOE.md`

## Why Split

- Keeps general pilot experience simple and stable.
- Prevents ME COE-specific logic/routes from leaking into mass-user deployment.
- Allows independent release cadence for pilot vs specialized dashboard.
