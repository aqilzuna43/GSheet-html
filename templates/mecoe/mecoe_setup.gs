/**
 * Dedicated setup script for the ME COE date-based timeline sheet.
 *
 * Usage:
 * 1) Add this file to your Apps Script project.
 * 2) Run setupMECOEScheduleSheet() once.
 */

const MECOE_SETUP_SHEET_NAME = 'ME_COE_Schedule';
const MECOE_SETUP_HEADERS = [
  'ID',
  'Phase',
  'Deliverable / Document',
  'Site',
  'Docs',
  'Priority',
  'Status',
  'Owner',
  'Start Date',
  'End Date',
  'Milestone Date',
  'Notes / Actions',
];
const MECOE_SETUP_PRIORITY_VALUES = ['P0', 'P0*', 'P1', 'P2'];
const MECOE_SETUP_STATUS_VALUES = ['Not Started', 'Draft', 'In Review', 'Completed'];

function setupMECOEScheduleSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateMECOESheet_(ss, MECOE_SETUP_SHEET_NAME);
  sheet.clear({ contentsOnly: false });
  ensureMECOESize_(sheet, 1000, MECOE_SETUP_HEADERS.length);

  sheet.getRange(1, 1, 1, MECOE_SETUP_HEADERS.length).setValues([MECOE_SETUP_HEADERS]);
  styleMECOESheet_(sheet);
  applyMECOEValidation_(sheet);
  ensureMECOEDashboardConfig_(ss);

  SpreadsheetApp.getUi().alert('ME_COE_Schedule date-model setup completed.');
}

function getOrCreateMECOESheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function ensureMECOESize_(sheet, rows, cols) {
  if (sheet.getMaxRows() < rows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), rows - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < cols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), cols - sheet.getMaxColumns());
  }
}

function styleMECOESheet_(sheet) {
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);

  sheet.getRange(1, 1, 1, MECOE_SETUP_HEADERS.length)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBackground('#e8f0fe');

  sheet.setColumnWidth(1, 90);
  sheet.setColumnWidth(2, 70);
  sheet.setColumnWidth(3, 420);
  sheet.setColumnWidth(4, 85);
  sheet.setColumnWidth(5, 55);
  sheet.setColumnWidth(6, 80);
  sheet.setColumnWidth(7, 110);
  sheet.setColumnWidth(8, 120);
  sheet.setColumnWidth(9, 110);
  sheet.setColumnWidth(10, 110);
  sheet.setColumnWidth(11, 120);
  sheet.setColumnWidth(12, 460);

  sheet.getRange(2, 9, 999, 3).setNumberFormat('yyyy-mm-dd').setHorizontalAlignment('center');
  sheet.getRange(2, 1, 999, 8).setVerticalAlignment('top').setWrap(true);
  sheet.getRange(2, 12, 999, 1).setVerticalAlignment('top').setWrap(true);
}

function applyMECOEValidation_(sheet) {
  const maxRows = Math.max(sheet.getMaxRows(), 1000);
  const priorityRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(MECOE_SETUP_PRIORITY_VALUES, true)
    .setAllowInvalid(true)
    .build();
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(MECOE_SETUP_STATUS_VALUES, true)
    .setAllowInvalid(true)
    .build();
  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(true)
    .build();

  sheet.getRange(2, 6, maxRows - 1, 1).setDataValidation(priorityRule);
  sheet.getRange(2, 7, maxRows - 1, 1).setDataValidation(statusRule);
  sheet.getRange(2, 9, maxRows - 1, 3).setDataValidation(dateRule);
}

/**
 * Best-effort migration from a legacy wide timeline sheet to the date model.
 * Copies core metadata fields and leaves date columns blank for manual completion.
 */
function migrateLegacyMECOEToDateModel(sourceSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceName = sourceSheetName || 'ME_COE';
  const source = ss.getSheetByName(sourceName);
  if (!source) throw new Error('Source sheet not found: ' + sourceName);

  const target = getOrCreateMECOESheet_(ss, MECOE_SETUP_SHEET_NAME);
  const srcValues = source.getDataRange().getDisplayValues();
  if (!srcValues.length) throw new Error('Source sheet is empty.');

  const headerRow = findLegacyHeaderRow_(srcValues);
  if (headerRow === -1) {
    throw new Error('Could not detect legacy header row. Expected Phase + Deliverable / Document.');
  }

  const rows = [];
  for (let r = headerRow + 1; r < srcValues.length; r++) {
    const row = srcValues[r];
    const phase = String(row[0] || '').trim();
    const deliverable = String(row[1] || '').trim();
    if (!phase && !deliverable) continue;
    rows.push([
      '',
      phase,
      deliverable,
      String(row[2] || '').trim(),
      String(row[3] || '').trim(),
      String(row[4] || '').trim(),
      String(row[5] || '').trim(),
      '',
      '',
      '',
      '',
      String(row[6] || '').trim(),
    ]);
  }

  setupMECOEScheduleSheet();
  if (rows.length) {
    target.getRange(2, 1, rows.length, MECOE_SETUP_HEADERS.length).setValues(rows);
  }
  SpreadsheetApp.getUi().alert('Migration completed. Please fill Start/End/Milestone dates.');
}

function findLegacyHeaderRow_(values) {
  for (let i = 0; i < values.length; i++) {
    const c0 = String(values[i][0] || '').trim().toLowerCase();
    const c1 = String(values[i][1] || '').trim().toLowerCase();
    if (c0 === 'phase' && c1.indexOf('deliverable') !== -1) return i;
  }
  return -1;
}

function ensureMECOEDashboardConfig_(ss) {
  const sheet = getOrCreateMECOEConfigSheet_(ss);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#f3f4f6');
    sheet.setColumnWidth(1, 170);
    sheet.setColumnWidth(2, 320);
  }
  setMECOEConfigValueIfEmpty_(sheet, 'mecoe_title', 'ME COE Timeline Dashboard');
}

function getOrCreateMECOEConfigSheet_(ss) {
  let sheet = ss.getSheetByName('Config');
  if (!sheet) sheet = ss.insertSheet('Config');
  return sheet;
}

function setMECOEConfigValueIfEmpty_(sheet, key, defaultValue) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const values = sheet.getRange(1, 1, lastRow, 2).getDisplayValues();
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][0] || '').trim().toLowerCase() !== String(key).toLowerCase()) continue;
    if (!String(values[r][1] || '').trim()) {
      sheet.getRange(r + 1, 2).setValue(defaultValue);
    }
    return;
  }
  sheet.appendRow([key, defaultValue]);
}
