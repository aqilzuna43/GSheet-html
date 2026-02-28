/**
 * Google Sheets setup script for the Visual Schedule pilot.
 * Usage:
 * 1) Create/import a sheet and name it "Schedule".
 * 2) Extensions -> Apps Script -> paste this file.
 * 3) Run setupScheduleSheet() once.
 */

const SETUP_SHEET_NAME = 'Schedule';
const SETUP_REQUIRED_HEADERS = [
  'ID',
  'Title',
  'Start Date',
  'End Date',
  'Owner',
  'Department',
  'Status',
  'Description',
  'Tags',
];

const SETUP_STATUS_VALUES = ['Not Started', 'In Progress', 'At Risk', 'Blocked', 'Completed'];
const SETUP_DEPARTMENT_VALUES = ['Engineering', 'Product', 'Mechanical', 'Quality', 'Management'];

function registerScheduleToolsMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('Schedule Tools')
    .addItem('Setup / Repair Sheet', 'setupScheduleSheet')
    .addItem('Validate Data Now', 'validateScheduleData')
    .addToUi();
}

function onEditScheduleValidation_(e) {
  const sheet = e && e.range ? e.range.getSheet() : null;
  if (!sheet || sheet.getName() !== SETUP_SHEET_NAME) return;

  const editedRow = e.range.getRow();
  const editedCol = e.range.getColumn();
  if (editedRow === 1) {
    SpreadsheetApp.getUi().alert('Header row is protected. Please do not modify column names.');
    setupScheduleSheet();
    return;
  }

  // Re-check date integrity only when Start/End Date changes.
  if (editedCol === 3 || editedCol === 4) {
    validateScheduleRow_(sheet, editedRow);
  }
}

function setupScheduleSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, SETUP_SHEET_NAME);

  ensureHeaders_(sheet);
  applyHeaderStyle_(sheet);
  applyColumnFormats_(sheet);
  applyDataValidation_(sheet);
  protectHeaderRow_(sheet);
  ensureScheduleDashboardConfig_(ss);
  validateScheduleData();

  SpreadsheetApp.getUi().alert('Schedule sheet setup completed.');
}

function validateScheduleData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETUP_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Sheet "Schedule" not found.');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  let hasErrors = false;
  for (let row = 2; row <= lastRow; row++) {
    const rowHasError = !validateScheduleRow_(sheet, row);
    if (rowHasError) hasErrors = true;
  }

  if (hasErrors) {
    SpreadsheetApp.getUi().alert('Validation finished with issues. Check red highlighted date cells.');
  } else {
    SpreadsheetApp.getUi().alert('Validation passed. No date issues found.');
  }
}

function validateScheduleRow_(sheet, row) {
  const startCell = sheet.getRange(row, 3);
  const endCell = sheet.getRange(row, 4);
  const startDate = startCell.getValue();
  const endDate = endCell.getValue();

  // Clear prior highlights if row is now valid/empty.
  startCell.setBackground(null);
  endCell.setBackground(null);

  if (!startDate || !endDate) return true;
  if (!(startDate instanceof Date) || !(endDate instanceof Date)) {
    startCell.setBackground('#fce8e6');
    endCell.setBackground('#fce8e6');
    return false;
  }

  if (endDate < startDate) {
    startCell.setBackground('#fce8e6');
    endCell.setBackground('#fce8e6');
    return false;
  }

  return true;
}

function getOrCreateSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function ensureHeaders_(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, SETUP_REQUIRED_HEADERS.length);
  headerRange.setValues([SETUP_REQUIRED_HEADERS]);
  if (sheet.getMaxColumns() > SETUP_REQUIRED_HEADERS.length) {
    sheet.hideColumns(SETUP_REQUIRED_HEADERS.length + 1, sheet.getMaxColumns() - SETUP_REQUIRED_HEADERS.length);
  }
}

function applyHeaderStyle_(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, SETUP_REQUIRED_HEADERS.length);
  headerRange
    .setFontWeight('bold')
    .setBackground('#e8f0fe')
    .setHorizontalAlignment('center')
    .setWrap(true);
  sheet.setFrozenRows(1);
}

function applyColumnFormats_(sheet) {
  sheet.setColumnWidths(1, 9, 140);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(8, 260);
  sheet.getRange('C:D').setNumberFormat('yyyy-mm-dd');
}

function applyDataValidation_(sheet) {
  const maxRows = Math.max(sheet.getMaxRows(), 1000);
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(SETUP_STATUS_VALUES, true)
    .setAllowInvalid(false)
    .build();
  const deptRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(SETUP_DEPARTMENT_VALUES, true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange(2, 7, maxRows - 1, 1).setDataValidation(statusRule); // Status
  sheet.getRange(2, 6, maxRows - 1, 1).setDataValidation(deptRule); // Department
}

function protectHeaderRow_(sheet) {
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  const existing = protections.find((p) => p.getRange().getA1Notation() === '1:1');
  if (existing) return;

  const protection = sheet.getRange('1:1').protect();
  protection.setDescription('Protect Schedule header row');
  const me = Session.getEffectiveUser();
  protection.addEditor(me);
  protection.removeEditors(protection.getEditors().filter((u) => u.getEmail() !== me.getEmail()));
  if (protection.canDomainEdit()) protection.setDomainEdit(false);
}

function ensureScheduleDashboardConfig_(ss) {
  const sheet = getOrCreateScheduleConfigSheet_(ss);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#f3f4f6');
    sheet.setColumnWidth(1, 170);
    sheet.setColumnWidth(2, 320);
  }
  setScheduleConfigValueIfEmpty_(sheet, 'standard_title', 'Visual Schedule Dashboard');
}

function getOrCreateScheduleConfigSheet_(ss) {
  let sheet = ss.getSheetByName('Config');
  if (!sheet) sheet = ss.insertSheet('Config');
  return sheet;
}

function setScheduleConfigValueIfEmpty_(sheet, key, defaultValue) {
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
